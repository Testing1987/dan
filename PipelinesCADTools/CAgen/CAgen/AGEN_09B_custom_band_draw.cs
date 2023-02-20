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
    public partial class AGEN_custom_band_draw : Form
    {


        double custom_band_separation = 0;

        public AGEN_custom_band_draw()
        {
            InitializeComponent();
        }


        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_draw_custom_band);
            lista_butoane.Add(button_read_band_to_xl);
            lista_butoane.Add(button_show_custom_scan);
            lista_butoane.Add(comboBox_custom_atr_distance);
            lista_butoane.Add(comboBox_custom_atr_field1);
            lista_butoane.Add(comboBox_custom_atr_field2);
            lista_butoane.Add(comboBox_custom_atr_sta1);
            lista_butoane.Add(comboBox_custom_atr_sta2);
            lista_butoane.Add(comboBox_custom_block);
            lista_butoane.Add(comboBox_end);
            lista_butoane.Add(comboBox_custom_excel_name);
            lista_butoane.Add(comboBox_start);
            lista_butoane.Add(button_draw_rectangles);

            foreach (Control bt1 in lista_butoane)
            {
                bt1.Enabled = false;
            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_draw_custom_band);
            lista_butoane.Add(button_read_band_to_xl);
            lista_butoane.Add(button_show_custom_scan);
            lista_butoane.Add(comboBox_custom_atr_distance);
            lista_butoane.Add(comboBox_custom_atr_field1);
            lista_butoane.Add(comboBox_custom_atr_field2);
            lista_butoane.Add(comboBox_custom_atr_sta1);
            lista_butoane.Add(comboBox_custom_atr_sta2);
            lista_butoane.Add(comboBox_custom_block);
            lista_butoane.Add(comboBox_end);
            lista_butoane.Add(comboBox_custom_excel_name);
            lista_butoane.Add(comboBox_start);
            lista_butoane.Add(button_draw_rectangles);


            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }



        private void button_show_custom_scan_Click(object sender, EventArgs e)
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

            _AGEN_mainform.tpage_cust_draw.Hide();
            _AGEN_mainform.tpage_sheet_gen.Hide();


            _AGEN_mainform.tpage_cust_scan.Show();
        }

        private void button_custom_refresh_Click(object sender, EventArgs e)
        {


            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        comboBox_custom_atr_field1.Items.Clear();
                        comboBox_custom_atr_field2.Items.Clear();
                        comboBox_custom_atr_distance.Items.Clear();
                        comboBox_custom_atr_sta1.Items.Clear();
                        comboBox_custom_atr_sta2.Items.Clear();

                        Functions.Incarca_existing_Blocks_with_attributes_to_combobox(comboBox_custom_block);
                        this.Refresh();
                    }
                }

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }



        }

        private void comboBox_custom_block_SelectedIndexChanged(object sender, EventArgs e)
        {

            Functions.Incarca_existing_Atributes_to_combobox(comboBox_custom_block.Text, comboBox_custom_atr_field1);
            Functions.Incarca_existing_Atributes_to_combobox(comboBox_custom_block.Text, comboBox_custom_atr_field2);
            Functions.Incarca_existing_Atributes_to_combobox(comboBox_custom_block.Text, comboBox_custom_atr_distance);
            Functions.Incarca_existing_Atributes_to_combobox(comboBox_custom_block.Text, comboBox_custom_atr_sta1);
            Functions.Incarca_existing_Atributes_to_combobox(comboBox_custom_block.Text, comboBox_custom_atr_sta2);

        }





        private Point3d get_band_insertion_point_and_band_height(string band_name)
        {
            Point3d pt_ins = new Point3d();

            if (_AGEN_mainform.Data_Table_custom_bands != null)
            {
                if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count > 0)
                {
                    for (int i = 0; i < _AGEN_mainform.Data_Table_custom_bands.Rows.Count; ++i)
                    {
                        if (_AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"] != DBNull.Value)
                        {
                            string bn = Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"]);
                            if (bn != null)
                            {
                                if (bn == band_name)
                                {
                                    double x0 = 0;
                                    double y0 = 0;
                                    double bh = 0;


                                    if (_AGEN_mainform.Data_Table_custom_bands.Rows[i]["viewport_ms_x"] != DBNull.Value &&
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["viewport_ms_y"] != DBNull.Value &&
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["viewport_height"] != DBNull.Value &&
                                        Functions.IsNumeric(Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["viewport_ms_x"])) == true &&
                                         Functions.IsNumeric(Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["viewport_ms_y"])) == true &&
                                        Functions.IsNumeric(Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["viewport_height"])) == true)
                                    {

                                        bh = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["viewport_height"]);
                                        x0 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["viewport_ms_x"]);
                                        y0 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["viewport_ms_y"]);
                                    }


                                    if (_AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_separation"] != DBNull.Value)
                                    {
                                        custom_band_separation = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_separation"]);
                                    }

                                    pt_ins = new Point3d(x0, y0, bh);

                                    if (_AGEN_mainform.Data_Table_custom_bands.Rows[i]["Custom_scale"] != DBNull.Value)
                                    {
                                        _AGEN_mainform.custom_band_scale = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["Custom_scale"]);
                                    }

                                    if (comboBox_custom_block.Text != "")
                                    {
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["block_name"] = comboBox_custom_block.Text;
                                    }
                                    else
                                    {
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["block_name"] = DBNull.Value;
                                    }


                                    if (comboBox_custom_atr_sta1.Text != "")
                                    {
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["block_sta_atr1"] = comboBox_custom_atr_sta1.Text;
                                    }
                                    else
                                    {
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["block_sta_atr1"] = DBNull.Value;
                                    }


                                    if (comboBox_custom_atr_sta2.Text != "")
                                    {
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["block_sta_atr2"] = comboBox_custom_atr_sta2.Text;
                                    }
                                    else
                                    {
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["block_sta_atr2"] = DBNull.Value;
                                    }


                                    if (comboBox_custom_atr_distance.Text != "")
                                    {
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["block_len_atr"] = comboBox_custom_atr_distance.Text;
                                    }
                                    else
                                    {
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["block_len_atr"] = DBNull.Value;
                                    }

                                    if (comboBox_custom_atr_field1.Text != "")
                                    {
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["block_field1"] = comboBox_custom_atr_field1.Text;
                                    }
                                    else
                                    {
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["block_field1"] = DBNull.Value;
                                    }


                                    if (comboBox_custom_atr_field2.Text != "")
                                    {
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["block_field2"] = comboBox_custom_atr_field2.Text;
                                    }
                                    else
                                    {
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["block_field2"] = DBNull.Value;
                                    }

                                }
                                string segment1 = _AGEN_mainform.tpage_setup.Get_segment_name1();
                                if (segment1 == "not defined") segment1 = "";
                                if (bn + "_" + segment1 == band_name)
                                {
                                    double x0 = 0;
                                    double y0 = 0;
                                    double bh = 0;


                                    if (_AGEN_mainform.Data_Table_custom_bands.Rows[i]["viewport_ms_x"] != DBNull.Value &&
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["viewport_ms_y"] != DBNull.Value &&
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["viewport_height"] != DBNull.Value &&
                                        Functions.IsNumeric(Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["viewport_ms_x"])) == true &&
                                         Functions.IsNumeric(Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["viewport_ms_y"])) == true &&
                                        Functions.IsNumeric(Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["viewport_height"])) == true)
                                    {

                                        bh = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["viewport_height"]);
                                        x0 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["viewport_ms_x"]);
                                        y0 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["viewport_ms_y"]);
                                    }

                                    if (_AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_separation"] != DBNull.Value)
                                    {
                                        custom_band_separation = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_separation"]);
                                    }
                                    pt_ins = new Point3d(x0, y0, bh);

                                    if (comboBox_custom_block.Text != "")
                                    {
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["block_name"] = comboBox_custom_block.Text;
                                    }
                                    else
                                    {
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["block_name"] = DBNull.Value;
                                    }


                                    if (comboBox_custom_atr_sta1.Text != "")
                                    {
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["block_sta_atr1"] = comboBox_custom_atr_sta1.Text;
                                    }
                                    else
                                    {
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["block_sta_atr1"] = DBNull.Value;
                                    }


                                    if (comboBox_custom_atr_sta2.Text != "")
                                    {
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["block_sta_atr2"] = comboBox_custom_atr_sta2.Text;
                                    }
                                    else
                                    {
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["block_sta_atr2"] = DBNull.Value;
                                    }


                                    if (comboBox_custom_atr_distance.Text != "")
                                    {
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["block_len_atr"] = comboBox_custom_atr_distance.Text;
                                    }
                                    else
                                    {
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["block_len_atr"] = DBNull.Value;
                                    }

                                    if (comboBox_custom_atr_field1.Text != "")
                                    {
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["block_field1"] = comboBox_custom_atr_field1.Text;
                                    }
                                    else
                                    {
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["block_field1"] = DBNull.Value;
                                    }


                                    if (comboBox_custom_atr_field2.Text != "")
                                    {
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["block_field2"] = comboBox_custom_atr_field2.Text;
                                    }
                                    else
                                    {
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["block_field2"] = DBNull.Value;
                                    }

                                    if (_AGEN_mainform.Data_Table_custom_bands.Rows[i]["Custom_scale"] != DBNull.Value)
                                    {
                                        _AGEN_mainform.custom_band_scale = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["Custom_scale"]);
                                    }


                                }

                            }
                        }
                    }
                }
            }


            return pt_ins;
        }



        private void button_draw_custom_band_Click(object sender, EventArgs e)
        {

            string lnp = "Agen_no_plot_" + comboBox_custom_excel_name.Text.Replace(" ", "");
            int lr = 1;
            if (_AGEN_mainform.Left_to_Right == false) lr = -1;



            string cfg1 = System.IO.Path.GetFileName(_AGEN_mainform.config_path);
            if (Functions.Get_if_workbook_is_open_in_Excel(cfg1) == true)
            {
                MessageBox.Show("Please close the " + cfg1 + " file");
                return;
            }


            if (comboBox_custom_excel_name.Text == "")
            {
                MessageBox.Show("you did not specified the excel file name");
                return;
            }




            _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;
            _AGEN_mainform.tpage_processing.Show();
            // Ag.WindowState = FormWindowState.Minimized;

            if (comboBox_custom_block.Text == "")
            {
                _AGEN_mainform.tpage_processing.Hide();
                MessageBox.Show("you did not specified custom block name");
                Ag.WindowState = FormWindowState.Normal;
                set_enable_true();
                return;
            }

            if (_AGEN_mainform.Vw_height <= 0)

            {
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


                set_enable_true();
                return;
            }

            Point3d cust_point0 = get_band_insertion_point_and_band_height(comboBox_custom_excel_name.Text);

            double custom_band_width = _AGEN_mainform.Vw_width;
            double custom_band_height = cust_point0.Z;
            cust_point0 = new Point3d(cust_point0.X, cust_point0.Y, 0);

            if (custom_band_height <= 0)
            {
                _AGEN_mainform.tpage_processing.Hide();
                MessageBox.Show("you did not specified the band height in the config file");
                Ag.WindowState = FormWindowState.Normal;
                set_enable_true();
                return;
            }

            if (custom_band_separation == 0) custom_band_separation = _AGEN_mainform.Band_Separation;

            System.Data.DataTable dt_cus_data = null;

            set_enable_false();

            if (System.IO.Directory.Exists(_AGEN_mainform.tpage_setup.Get_project_database_folder()) == true)
            {
                string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                {
                    ProjF = ProjF + "\\";
                }


                string fisier_custom = ProjF + comboBox_custom_excel_name.Text + ".xlsx";

                if (System.IO.File.Exists(fisier_custom) == false)
                {
                    _AGEN_mainform.tpage_processing.Hide();
                    set_enable_true();
                    MessageBox.Show("the custom band data file does not exist");
                    return;
                }
                else
                {
                    Functions.create_backup(fisier_custom);
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
                    MessageBox.Show("the sheet index data file does not exist");
                    _AGEN_mainform.dt_station_equation = null;
                    return;
                }

                dt_cus_data = Load_existing_custom_data(fisier_custom);
                if (dt_cus_data == null)
                {
                    _AGEN_mainform.tpage_processing.Hide();
                    set_enable_true();
                    MessageBox.Show("no custom data found");
                    return;
                }



                _AGEN_mainform.dt_sheet_index = _AGEN_mainform.tpage_setup.Load_existing_sheet_index(fisier_si, comboBox_custom_excel_name.Text);


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

            if (dt_cus_data.Rows.Count == 0)
            {
                _AGEN_mainform.tpage_processing.Hide();
                set_enable_true();
                MessageBox.Show("custom file does not have any data");
                return;
            }

            if (_AGEN_mainform.dt_sheet_index.Rows.Count == 0)
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

            string Sta1 = "Sta1";
            string Sta2 = "Sta2";


            string f1 = "field1";
            string f2 = "field2";

            string Pageno = "Page";
            string Rect_len = "RectangleML";

            string stretch_val = "StrechVal";
            string BandL = "BandL";
            string DeltaX_col = "DeltaX";
            string stretch_val_orig = "StrechValoriginal";
            string Sta1CSF = "Sta1CSF";
            string Sta2CSF = "Sta2CSF";

            string colm1 = "M1";
            string colm2 = "M2";
            string colm1csf = "M1csf";
            string colm2csf = "M2csf";






            System.Data.DataTable Data_table_compiled = new System.Data.DataTable();
            Data_table_compiled.Columns.Add(_AGEN_mainform.Col_dwg_name, typeof(string));
            Data_table_compiled.Columns.Add(Sta1, typeof(double));
            Data_table_compiled.Columns.Add(Sta2, typeof(double));
            Data_table_compiled.Columns.Add(f1, typeof(string));
            Data_table_compiled.Columns.Add(f2, typeof(string));
            Data_table_compiled.Columns.Add(Pageno, typeof(int));
            Data_table_compiled.Columns.Add(Rect_len, typeof(double));
            Data_table_compiled.Columns.Add(BandL, typeof(double));
            Data_table_compiled.Columns.Add(DeltaX_col, typeof(double));
            Data_table_compiled.Columns.Add(colm1, typeof(double));
            Data_table_compiled.Columns.Add(colm2, typeof(double));
            Data_table_compiled.Columns.Add(stretch_val, typeof(double));
            Data_table_compiled.Columns.Add(stretch_val_orig, typeof(double));
            Data_table_compiled.Columns.Add("min_stretch", typeof(double));
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
                        string Block_name = comboBox_custom_block.Text;
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

                        #region USA station eq
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
                            _AGEN_mainform.dt_station_equation = null;
                        }
                        #endregion



                        double Min_dist = 0;
                        BlockReference BR1 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "", Block_name, new Point3d(0, 0, 0), 1 / _AGEN_mainform.custom_band_scale, 0, "0", new System.Collections.Specialized.StringCollection(), new System.Collections.Specialized.StringCollection());
                        Min_dist = Functions.Get_distance1_block(BR1);
                        BR1.Erase();



                        string lname1 = "Agen_band_" + comboBox_custom_excel_name.Text.Replace(" ", "");
                        Functions.Creaza_layer(lname1, 7, true);
                        Functions.Creaza_layer(lnp, 30, false);





                        List<int> lista_bands_for_generation = new List<int>();

                        if (comboBox_start.Text != "" & comboBox_end.Text != "")
                        {
                            lista_bands_for_generation = _AGEN_mainform.tpage_setup.create_band_list_of_dwg(comboBox_start.Text, comboBox_end.Text);
                        }

                        if (comboBox_start.Text == "" || comboBox_end.Text == "")
                        {
                            lista_bands_for_generation = _AGEN_mainform.tpage_setup.create_band_list_of_dwg("", "");
                        }

                        if (lista_bands_for_generation.Count == 0)
                        {
                            _AGEN_mainform.tpage_processing.Hide();
                            set_enable_true();
                            MessageBox.Show("please check your input");
                            return;
                        }



                        for (int i = 0; i < dt_cus_data.Rows.Count; ++i)
                        {
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


                            if (dt_cus_data.Rows[i]["2DStaBeg"] != DBNull.Value && dt_cus_data.Rows[i]["2DStaEnd"] != DBNull.Value)
                            {
                                Station1 = Convert.ToDouble(dt_cus_data.Rows[i]["2DStaBeg"]);
                                Station2 = Convert.ToDouble(dt_cus_data.Rows[i]["2DStaEnd"]);
                            }

                            if (_AGEN_mainform.Project_type == "3D")
                            {
                                if (dt_cus_data.Rows[i]["3DStaBeg"] != DBNull.Value && dt_cus_data.Rows[i]["3DStaEnd"] != DBNull.Value)
                                {
                                    Station1 = Convert.ToDouble(dt_cus_data.Rows[i]["3DStaBeg"]);
                                    Station2 = Convert.ToDouble(dt_cus_data.Rows[i]["3DStaEnd"]);
                                }
                            }


                            Station1 = Math.Round(Station1, _AGEN_mainform.round1);
                            Station2 = Math.Round(Station2, _AGEN_mainform.round1);

                            if (Station1 > poly_length)
                            {
                                Station1 = poly_length;
                            }

                            if (Station2 > poly_length)
                            {
                                Station2 = poly_length;
                            }

                            if (Station1 < 0) Station1 = 0;
                            if (Station2 < 0) Station2 = 0;

                            if (_AGEN_mainform.COUNTRY == "USA")
                            {
                                Station1_CSF = Functions.Station_equation_ofV2(Station1, _AGEN_mainform.dt_station_equation);
                                Station2_CSF = Functions.Station_equation_ofV2(Station2, _AGEN_mainform.dt_station_equation);
                            }
                            else if (_AGEN_mainform.COUNTRY == "CANADA")
                            {
                                if (dt_cus_data.Rows[i]["X_Beg"] != DBNull.Value && dt_cus_data.Rows[i]["Y_Beg"] != DBNull.Value &&
                                                                dt_cus_data.Rows[i]["X_End"] != DBNull.Value && dt_cus_data.Rows[i]["Y_End"] != DBNull.Value)
                                {

                                    px_start = Convert.ToDouble(dt_cus_data.Rows[i]["X_Beg"]);
                                    py_start = Convert.ToDouble(dt_cus_data.Rows[i]["Y_Beg"]);
                                    px_end = Convert.ToDouble(dt_cus_data.Rows[i]["X_End"]);
                                    py_end = Convert.ToDouble(dt_cus_data.Rows[i]["Y_End"]);

                                    double param1 = poly2d.GetParameterAtPoint(poly2d.GetClosestPointTo(new Point3d(px_start, py_start, 0), Vector3d.ZAxis, false));
                                    Station1 = poly3d.GetDistanceAtParameter(param1);
                                    double param2 = poly2d.GetParameterAtPoint(poly2d.GetClosestPointTo(new Point3d(px_end, py_end, 0), Vector3d.ZAxis, false));
                                    Station2 = poly3d.GetDistanceAtParameter(param2);

                                    Station1 = Math.Round(Station1, _AGEN_mainform.round1);
                                    Station2 = Math.Round(Station2, _AGEN_mainform.round1);

                                    if (Station1 >= poly3d.Length) Station1 = poly3d.Length - 0.0001;
                                    if (Station2 >= poly3d.Length) Station2 = poly3d.Length - 0.0001;

                                    double d1 = poly2d.GetDistanceAtParameter(param1);
                                    double d2 = poly2d.GetDistanceAtParameter(param2);
                                    double b1 = -1.23456;
                                    double b2 = -1.23456;
                                    Station1_CSF = Functions.get_stationCSF_from_point(poly2d, new Point3d(px_start, py_start, 0), d1, _AGEN_mainform.dt_centerline, ref b1);
                                    Station2_CSF = Functions.get_stationCSF_from_point(poly2d, new Point3d(px_end, py_end, 0), d2, _AGEN_mainform.dt_centerline, ref b2);

                                    dt_cus_data.Rows[i]["3DStaBeg"] = Station1_CSF;
                                    dt_cus_data.Rows[i]["3DStaEnd"] = Station2_CSF;
                                    dt_cus_data.Rows[i]["2DStaBeg"] = DBNull.Value;
                                    dt_cus_data.Rows[i]["2DStaEnd"] = DBNull.Value;






                                }
                                else
                                {
                                    Station1_CSF = Math.Round(Station1, _AGEN_mainform.round1);
                                    Station2_CSF = Math.Round(Station2, _AGEN_mainform.round1);

                                    bool is_found1 = false;
                                    bool is_found2 = false;

                                    for (int k = 0; k < _AGEN_mainform.dt_centerline.Rows.Count - 1; ++k)
                                    {


                                        if (_AGEN_mainform.dt_centerline.Rows[k]["3DSta"] != DBNull.Value && _AGEN_mainform.dt_centerline.Rows[k]["X"] != DBNull.Value && _AGEN_mainform.dt_centerline.Rows[k]["Y"] != DBNull.Value &&
                                            _AGEN_mainform.dt_centerline.Rows[k + 1]["3DSta"] != DBNull.Value && _AGEN_mainform.dt_centerline.Rows[k + 1]["X"] != DBNull.Value && _AGEN_mainform.dt_centerline.Rows[k + 1]["Y"] != DBNull.Value)
                                        {
                                            double s1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[k]["3DSta"]);
                                            double s2 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[k + 1]["3DSta"]);

                                            double x1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[k]["X"]);
                                            double x2 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[k + 1]["X"]);

                                            double y1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[k]["Y"]);
                                            double y2 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[k + 1]["Y"]);

                                            if (is_found1 == false && s1 <= Station1 && s2 > Station1)
                                            {
                                                Polyline poly1 = new Polyline();
                                                poly1.AddVertexAt(0, new Point2d(x1, y1), 0, 0, 0);
                                                poly1.AddVertexAt(1, new Point2d(x2, y2), 0, 0, 0);
                                                poly1.Elevation = poly2d.Elevation;
                                                double dif1 = Station1 - s1;
                                                double dist1 = s2 - s1;
                                                double len1 = Math.Pow((Math.Pow((x1 - x2), 2) + Math.Pow((y1 - y2), 2)), 0.5);
                                                double diferenta1 = dif1 * len1 / dist1;

                                                Point3d p1a = poly1.GetPointAtDist(diferenta1);
                                                Point3d pt_on_poly1 = poly2d.GetClosestPointTo(p1a, Vector3d.ZAxis, false);
                                                px_start = pt_on_poly1.X;
                                                py_start = pt_on_poly1.Y;

                                                if (_AGEN_mainform.Project_type == "3D")
                                                {
                                                    double param1 = poly2d.GetParameterAtPoint(pt_on_poly1);
                                                    if (param1 > poly3d.EndParam)
                                                    {
                                                        param1 = poly3d.EndParam;
                                                    }
                                                    Station1 = poly3d.GetDistanceAtParameter(param1);
                                                }
                                                else
                                                {
                                                    Station1 = poly3d.GetDistAtPoint(pt_on_poly1);
                                                }

                                                is_found1 = true;
                                            }

                                            if (is_found2 == false && s1 <= Station2 && s2 > Station2)
                                            {
                                                Polyline poly1 = new Polyline();
                                                poly1.AddVertexAt(0, new Point2d(x1, y1), 0, 0, 0);
                                                poly1.AddVertexAt(1, new Point2d(x2, y2), 0, 0, 0);
                                                poly1.Elevation = poly2d.Elevation;
                                                double dif2 = Station2 - s1;

                                                double dist2 = s2 - s1;
                                                double len1 = Math.Pow((Math.Pow((x1 - x2), 2) + Math.Pow((y1 - y2), 2)), 0.5);
                                                double diferenta2 = dif2 * len1 / dist2;


                                                Point3d p2a = poly1.GetPointAtDist(diferenta2);

                                                Point3d pt_on_poly2 = poly2d.GetClosestPointTo(p2a, Vector3d.ZAxis, false);
                                                px_end = pt_on_poly2.X;
                                                py_end = pt_on_poly2.Y;

                                                if (_AGEN_mainform.Project_type == "3D")
                                                {
                                                    double param2 = poly2d.GetParameterAtPoint(pt_on_poly2);
                                                    if (param2 > poly3d.EndParam)
                                                    {
                                                        param2 = poly3d.EndParam;
                                                    }
                                                    Station2 = poly3d.GetDistanceAtParameter(param2);
                                                }
                                                else
                                                {
                                                    Station2 = poly3d.GetDistAtPoint(pt_on_poly2);
                                                }
                                                is_found2 = true;
                                            }

                                        }

                                        if (is_found1 == true && is_found2 == true)
                                        {
                                            k = _AGEN_mainform.dt_centerline.Rows.Count;
                                        }
                                    }


                                    if (is_found1 == false || is_found2 == false)
                                    {
                                        MessageBox.Show("no station and poinys calculated\r\noperation aborted");

                                        _AGEN_mainform.tpage_processing.Hide();

                                        set_enable_true();

                                        Ag.WindowState = FormWindowState.Normal;
                                    }
                                }


                            }





                            if (Station1 != -1.123 && Station2 != -1.123)
                            {
                                string val1 = "";
                                string val2 = "";
                                if (dt_cus_data.Rows[i][7] != DBNull.Value)
                                {
                                    val1 = dt_cus_data.Rows[i][7].ToString();
                                }
                                if (dt_cus_data.Rows[i][8] != DBNull.Value)
                                {
                                    val2 = dt_cus_data.Rows[i][8].ToString();
                                }

                                double mx_start = -1.2345;
                                double my_start = -1.2345;
                                double mx_end = -1.2345;
                                double my_end = -1.2345;

                                double min_stretch_from_custom = 0;
                                if (dt_cus_data.Rows[i][0] != DBNull.Value)
                                {
                                    double ms = Convert.ToDouble(dt_cus_data.Rows[i][0]);
                                    if (ms > 0) min_stretch_from_custom = ms;
                                }

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

                                        double M1_DISPLAY = M1;
                                        double M2_DISPLAY = M2;


                                        if (_AGEN_mainform.COUNTRY == "CANADA")
                                        {


                                            double x1 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[j]["X_Beg"]);
                                            double Y1 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[j]["Y_Beg"]);

                                            Point3d pt1 = new Point3d(x1, Y1, poly2d.Elevation);
                                            Point3d pt_on_poly1 = poly2d.GetClosestPointTo(pt1, Vector3d.ZAxis, false);
                                            double param1 = poly2d.GetParameterAtPoint(pt_on_poly1);
                                            if (poly3d.EndParam < param1) param1 = poly3d.EndParam;
                                            M1 = poly3d.GetDistanceAtParameter(param1);



                                            double x2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[j]["X_End"]);
                                            double Y2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[j]["Y_End"]);

                                            Point3d pt2 = new Point3d(x2, Y2, poly2d.Elevation);
                                            Point3d pt_on_poly2 = poly2d.GetClosestPointTo(pt2, Vector3d.ZAxis, false);
                                            double param2 = poly2d.GetParameterAtPoint(pt_on_poly2);
                                            if (poly3d.EndParam < param2) param2 = poly3d.EndParam;
                                            M2 = poly3d.GetDistanceAtParameter(param2);

                                            if (_AGEN_mainform.dt_sheet_index.Columns.Contains("M1_CANADA") == true)
                                            {
                                                if (_AGEN_mainform.dt_sheet_index.Rows[j]["M1_CANADA"] != DBNull.Value)
                                                {
                                                    M1_DISPLAY = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[j]["M1_CANADA"]);
                                                }
                                            }

                                            if (_AGEN_mainform.dt_sheet_index.Columns.Contains("M2_CANADA") == true)
                                            {
                                                if (_AGEN_mainform.dt_sheet_index.Rows[j]["M2_CANADA"] != DBNull.Value)
                                                {
                                                    M2_DISPLAY = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[j]["M2_CANADA"]);
                                                }
                                            }

                                        }



                                        if (_AGEN_mainform.COUNTRY == "USA")
                                        {
                                            M1_DISPLAY = Functions.Station_equation_ofV2(M1, _AGEN_mainform.dt_station_equation);
                                            M2_DISPLAY = Functions.Station_equation_ofV2(M2, _AGEN_mainform.dt_station_equation);
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

                                        if (M1 > poly_length) M1 = poly_length - 0.0001;
                                        if (M2 > poly_length) M2 = poly_length - 0.0001;

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
                                            if (_AGEN_mainform.Project_type == "3D")
                                            {
                                                pm1 = poly3d.EndPoint;
                                            }
                                            else
                                            {
                                                pm1 = poly2d.EndPoint;
                                            }

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
                                            if (_AGEN_mainform.Project_type == "3D")
                                            {
                                                pm2 = poly3d.EndPoint;
                                            }
                                            else
                                            {
                                                pm2 = poly2d.EndPoint;
                                            }

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

                                        #region sta1>M1, sta2>M2 (m2 between sta1 and sta2)

                                        if (Math.Round(M1, 2) <= Math.Round(Station1, 2) && Math.Round(M2, 2) <= Math.Round(Station2, 2) && Math.Round(M1, 2) <= Math.Round(Station2, 2) && Math.Round(M2, 2) > Math.Round(Station1, 2))
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

                                                if (_AGEN_mainform.Project_type == "3D")
                                                {
                                                    ppt1 = poly3d.EndPoint;
                                                }
                                                else

                                                {
                                                    ppt1 = poly2d.EndPoint;
                                                }
                                            }
                                            Point3d Pt1 = Linie_M1_M2.GetClosestPointTo(ppt1, Vector3d.ZAxis, false);

                                            double stretch01 = Pt1.DistanceTo(pm2) * _AGEN_mainform.Vw_scale;
                                            double deltax1 = Pt1.DistanceTo(pm1) * _AGEN_mainform.Vw_scale;

                                            double rec_len = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                            double ban_len = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                            if (_AGEN_mainform.custom_band_scale != 1)
                                            {
                                                stretch01 = Pt1.DistanceTo(pm2);
                                                deltax1 = Pt1.DistanceTo(pm1);
                                                rec_len = Linie_M1_M2.Length;
                                                ban_len = Linie_M1_M2.Length;
                                            }

                                            if (checkBox_ignore_match_with_plan_view.Checked == true)
                                            {
                                                stretch01 = M2 - Station1;
                                                deltax1 = Station1 - M1;
                                                rec_len = M2 - M1;
                                                ban_len = M2 - M1;
                                                if (_AGEN_mainform.custom_band_scale == 1)
                                                {
                                                    stretch01 = stretch01 * _AGEN_mainform.Vw_scale;
                                                    deltax1 = deltax1 * _AGEN_mainform.Vw_scale;
                                                    rec_len = rec_len * _AGEN_mainform.Vw_scale;
                                                    ban_len = ban_len * _AGEN_mainform.Vw_scale;
                                                }

                                            }

                                            Data_table_compiled.Rows.Add();
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][_AGEN_mainform.Col_dwg_name] = _AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_dwg_name];
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta1] = Station1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta2] = M2;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta1CSF] = Station1_CSF;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta2CSF] = M2_DISPLAY;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][f1] = val1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][f2] = val2;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Pageno] = j + 1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm1] = M1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm2] = M2;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm1csf] = M1_DISPLAY;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm2csf] = M2_DISPLAY;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Rect_len] = rec_len;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][stretch_val] = stretch01;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][stretch_val_orig] = stretch01;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][BandL] = ban_len;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][DeltaX_col] = deltax1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["dt_row"] = i;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["xbeg"] = px_start;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["ybeg"] = py_start;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["xend"] = mx_end;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["yend"] = my_end;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["min_stretch"] = min_stretch_from_custom;

                                            Station1 = M2;
                                            Station1_CSF = M2_DISPLAY;

                                            px_start = mx_end;
                                            py_start = my_end;

                                            m_start = j + 1;
                                            Boolean_go_to_check_s1_s2 = true;
                                            goto L123;
                                        }

                                        #endregion

                                        #region sta1<=M1, sta2>=M2 (M2 AND M1 between sta1 and sta2)

                                        if (Math.Round(M1, 2) >= Math.Round(Station1, 2) && Math.Round(M2, 2) <= Math.Round(Station2, 2))
                                        {
                                            Point3d ppt1 = new Point3d();
                                            try
                                            {
                                                if (_AGEN_mainform.Project_type == "3D")
                                                {
                                                    ppt1 = poly3d.GetPointAtDist(M1);
                                                }
                                                else

                                                {
                                                    ppt1 = poly2d.GetPointAtDist(M1);
                                                }

                                            }
                                            catch (Autodesk.AutoCAD.Runtime.Exception ex)
                                            {

                                                if (_AGEN_mainform.Project_type == "3D")
                                                {
                                                    ppt1 = poly3d.EndPoint;
                                                }
                                                else

                                                {
                                                    ppt1 = poly2d.EndPoint;
                                                }
                                            }
                                            Point3d Pt1 = Linie_M1_M2.GetClosestPointTo(ppt1, Vector3d.ZAxis, false);

                                            double stretch01 = Pt1.DistanceTo(pm2) * _AGEN_mainform.Vw_scale;
                                            double deltax1 = Pt1.DistanceTo(pm1) * _AGEN_mainform.Vw_scale;

                                            double rec_len = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                            double ban_len = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                            if (_AGEN_mainform.custom_band_scale != 1)
                                            {
                                                stretch01 = Pt1.DistanceTo(pm2);
                                                deltax1 = Pt1.DistanceTo(pm1);
                                                rec_len = Linie_M1_M2.Length;
                                                ban_len = Linie_M1_M2.Length;
                                            }

                                            if (checkBox_ignore_match_with_plan_view.Checked == true)
                                            {
                                                stretch01 = M2 - M1;
                                                deltax1 = M1 - M1;
                                                rec_len = M2 - M1;
                                                ban_len = M2 - M1;
                                                if (_AGEN_mainform.custom_band_scale == 1)
                                                {
                                                    stretch01 = stretch01 * _AGEN_mainform.Vw_scale;
                                                    deltax1 = deltax1 * _AGEN_mainform.Vw_scale;
                                                    rec_len = rec_len * _AGEN_mainform.Vw_scale;
                                                    ban_len = ban_len * _AGEN_mainform.Vw_scale;
                                                }

                                            }

                                            Data_table_compiled.Rows.Add();
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][_AGEN_mainform.Col_dwg_name] = _AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_dwg_name];
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta1] = M1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta2] = M2;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta1CSF] = M1_DISPLAY;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta2CSF] = M2_DISPLAY;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][f1] = val1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][f2] = val2;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Pageno] = j + 1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm1] = M1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm2] = M2;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm1csf] = M1_DISPLAY;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm2csf] = M2_DISPLAY;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Rect_len] = rec_len;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][stretch_val] = stretch01;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][stretch_val_orig] = stretch01;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][BandL] = ban_len;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][DeltaX_col] = deltax1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["dt_row"] = i;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["xbeg"] = px_start;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["ybeg"] = py_start;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["xend"] = mx_end;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["yend"] = my_end;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["min_stretch"] = min_stretch_from_custom;

                                            Station1 = M2;
                                            Station1_CSF = M2_DISPLAY;

                                            px_start = mx_end;
                                            py_start = my_end;

                                            m_start = j + 1;
                                            Boolean_go_to_check_s1_s2 = true;
                                            goto L123;
                                        }

                                        #endregion

                                        #region m1<sta1, m2>sta2 (sta1 and sta2 between M1 and M2)

                                        if (Math.Round(Station1, 2) >= Math.Round(M1, 2) && Math.Round(Station2, 2) <= Math.Round(M2, 2) && Math.Round(Station1, 2) < Math.Round(M2, 2))
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

                                                if (_AGEN_mainform.Project_type == "3D")
                                                {
                                                    ppt1 = poly3d.EndPoint;
                                                }
                                                else
                                                {
                                                    ppt1 = poly2d.EndPoint;
                                                }
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
                                                if (_AGEN_mainform.Project_type == "3D")
                                                {
                                                    ppt2 = poly3d.EndPoint;
                                                }
                                                else
                                                {
                                                    ppt2 = poly2d.EndPoint;
                                                }
                                            }

                                            Point3d Pt1 = Linie_M1_M2.GetClosestPointTo(ppt1, Vector3d.ZAxis, false);
                                            Point3d Pt2 = Linie_M1_M2.GetClosestPointTo(ppt2, Vector3d.ZAxis, false);
                                            double stretch01 = Pt1.DistanceTo(Pt2) * _AGEN_mainform.Vw_scale;
                                            double deltax1 = Pt1.DistanceTo(pm1) * _AGEN_mainform.Vw_scale;
                                            double rec_len = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                            double ban_len = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                            if (_AGEN_mainform.custom_band_scale != 1)
                                            {
                                                stretch01 = Pt1.DistanceTo(Pt2);
                                                deltax1 = Pt1.DistanceTo(pm1);
                                                rec_len = Linie_M1_M2.Length;
                                                ban_len = Linie_M1_M2.Length;
                                            }


                                            if (checkBox_ignore_match_with_plan_view.Checked == true)
                                            {
                                                stretch01 = Station2 - Station1;
                                                deltax1 = Station1 - M1;
                                                rec_len = M2 - M1;
                                                ban_len = M2 - M1;
                                                if (_AGEN_mainform.custom_band_scale == 1)
                                                {
                                                    stretch01 = stretch01 * _AGEN_mainform.Vw_scale;
                                                    deltax1 = deltax1 * _AGEN_mainform.Vw_scale;
                                                    rec_len = rec_len * _AGEN_mainform.Vw_scale;
                                                    ban_len = ban_len * _AGEN_mainform.Vw_scale;
                                                }

                                            }

                                            Data_table_compiled.Rows.Add();
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][_AGEN_mainform.Col_dwg_name] = _AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_dwg_name];
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta1] = Station1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta2] = Station2;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta1CSF] = Station1_CSF;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta2CSF] = Station2_CSF;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][f1] = val1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][f2] = val2;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Pageno] = j + 1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm1] = M1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm2] = M2;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm1csf] = M1_DISPLAY;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm2csf] = M2_DISPLAY;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Rect_len] = rec_len;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][stretch_val] = stretch01;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][stretch_val_orig] = stretch01;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][BandL] = ban_len;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][DeltaX_col] = deltax1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["dt_row"] = i;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["xbeg"] = px_start;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["ybeg"] = py_start;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["xend"] = px_end;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["yend"] = py_end;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["min_stretch"] = min_stretch_from_custom;


                                            j = _AGEN_mainform.dt_sheet_index.Rows.Count;
                                            goto LS12end;
                                        }
                                    #endregion

                                    LS1S2:

                                        #region sta1 and sta2 between m1 and m2

                                        if (Math.Round(Station1, 2) >= Math.Round(M1, 2) && Math.Round(Station2, 2) <= Math.Round(M2, 2) && Math.Round(Station1, 2) < Math.Round(M2, 2))
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

                                                if (_AGEN_mainform.Project_type == "3D")
                                                {
                                                    ppt1 = poly3d.EndPoint;
                                                }
                                                else
                                                {
                                                    ppt1 = poly2d.EndPoint;
                                                }
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
                                                if (_AGEN_mainform.Project_type == "3D")
                                                {
                                                    ppt2 = poly3d.EndPoint;
                                                }
                                                else
                                                {
                                                    ppt2 = poly2d.EndPoint;
                                                }
                                            }


                                            Point3d Pt1 = Linie_M1_M2.GetClosestPointTo(ppt1, Vector3d.ZAxis, false);
                                            Point3d Pt2 = Linie_M1_M2.GetClosestPointTo(ppt2, Vector3d.ZAxis, false);
                                            double stretch01 = Pt1.DistanceTo(Pt2) * _AGEN_mainform.Vw_scale;
                                            double deltax1 = Pt1.DistanceTo(pm1) * _AGEN_mainform.Vw_scale;
                                            double rec_len = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                            double ban_len = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                            if (_AGEN_mainform.custom_band_scale != 1)
                                            {
                                                stretch01 = Pt1.DistanceTo(Pt2);
                                                deltax1 = Pt1.DistanceTo(pm1);
                                                rec_len = Linie_M1_M2.Length;
                                                ban_len = Linie_M1_M2.Length;
                                            }


                                            if (checkBox_ignore_match_with_plan_view.Checked == true)
                                            {
                                                stretch01 = Station2 - Station1;
                                                deltax1 = Station1 - M1;
                                                rec_len = M2 - M1;
                                                ban_len = M2 - M1;
                                                if (_AGEN_mainform.custom_band_scale == 1)
                                                {
                                                    stretch01 = stretch01 * _AGEN_mainform.Vw_scale;
                                                    deltax1 = deltax1 * _AGEN_mainform.Vw_scale;
                                                    rec_len = rec_len * _AGEN_mainform.Vw_scale;
                                                    ban_len = ban_len * _AGEN_mainform.Vw_scale;
                                                }

                                            }

                                            Data_table_compiled.Rows.Add();
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][_AGEN_mainform.Col_dwg_name] = _AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_dwg_name];
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta1] = Station1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta2] = Station2;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta1CSF] = Station1_CSF;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta2CSF] = Station2_CSF;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][f1] = val1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][f2] = val2;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Pageno] = j + 1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm1] = M1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm2] = M2;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm1csf] = M1_DISPLAY;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm2csf] = M2_DISPLAY;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Rect_len] = rec_len;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][stretch_val] = stretch01;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][stretch_val_orig] = stretch01;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][BandL] = ban_len;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][DeltaX_col] = deltax1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["dt_row"] = i;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["xbeg"] = px_start;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["ybeg"] = py_start;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["xend"] = px_end;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["yend"] = py_end;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["min_stretch"] = min_stretch_from_custom;


                                            j = _AGEN_mainform.dt_sheet_index.Rows.Count;
                                            goto LS12end;

                                        }

                                        #endregion
                                        //else if
                                        #region m2 between sta1 and sta2

                                        else if (Math.Round(Station1, 2) < Math.Round(M2, 2) && Math.Round(Station1, 2) >= Math.Round(M1, 2))
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

                                                if (_AGEN_mainform.Project_type == "3D")
                                                {
                                                    ppt1 = poly3d.EndPoint;
                                                }
                                                else
                                                {
                                                    ppt1 = poly2d.EndPoint;
                                                }
                                            }

                                            Point3d Pt1 = Linie_M1_M2.GetClosestPointTo(ppt1, Vector3d.ZAxis, false);
                                            double stretch01 = Pt1.DistanceTo(pm2) * _AGEN_mainform.Vw_scale;
                                            double deltax1 = Pt1.DistanceTo(pm1) * _AGEN_mainform.Vw_scale;
                                            double rec_len = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                            double ban_len = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                            if (_AGEN_mainform.custom_band_scale != 1)
                                            {
                                                stretch01 = Pt1.DistanceTo(pm2);
                                                deltax1 = Pt1.DistanceTo(pm1);
                                                rec_len = Linie_M1_M2.Length;
                                                ban_len = Linie_M1_M2.Length;
                                            }

                                            if (checkBox_ignore_match_with_plan_view.Checked == true)
                                            {
                                                stretch01 = M2 - Station1;
                                                deltax1 = Station1 - M1;
                                                rec_len = M2 - M1;
                                                ban_len = M2 - M1;
                                                if (_AGEN_mainform.custom_band_scale == 1)
                                                {
                                                    stretch01 = stretch01 * _AGEN_mainform.Vw_scale;
                                                    deltax1 = deltax1 * _AGEN_mainform.Vw_scale;
                                                    rec_len = rec_len * _AGEN_mainform.Vw_scale;
                                                    ban_len = ban_len * _AGEN_mainform.Vw_scale;
                                                }

                                            }
                                            Data_table_compiled.Rows.Add();
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][_AGEN_mainform.Col_dwg_name] = _AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_dwg_name];
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta1] = Station1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta2] = M2;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta1CSF] = Station1_CSF;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta2CSF] = M2_DISPLAY;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][f1] = val1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][f2] = val2;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Pageno] = j + 1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm1] = M1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm2] = M2;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm1csf] = M1_DISPLAY;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm2csf] = M2_DISPLAY;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Rect_len] = rec_len;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][stretch_val] = stretch01;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][stretch_val_orig] = stretch01;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][BandL] = ban_len;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][DeltaX_col] = deltax1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["dt_row"] = i;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["xbeg"] = px_start;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["ybeg"] = py_start;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["xend"] = mx_end;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["yend"] = my_end;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["min_stretch"] = min_stretch_from_custom;


                                            Station1 = M2;
                                            Station1_CSF = M2_DISPLAY;
                                            px_start = mx_end;
                                            py_start = my_end;



                                            m_start = j + 1;
                                            Boolean_go_to_check_s1_s2 = true;
                                            goto L123;
                                        }

                                        #endregion
                                    }
                                }
                            LS12end:
                                string xx = "";
                            }

                        }


                        // Alignment_generator.Functions.Transfer_datatable_to_new_excel_spreadsheet(Data_table_compiled);

                        int Pagep = -1;

                        if (Data_table_compiled != null)
                        {
                            Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                            Functions.Create_custom_od_table();

                            for (int i = 0; i < Data_table_compiled.Rows.Count; ++i)
                            {

                                double sta1 = Convert.ToDouble(Data_table_compiled.Rows[i][Sta1]);
                                double sta2 = Convert.ToDouble(Data_table_compiled.Rows[i][Sta2]);



                                int Page1 = Convert.ToInt32(Data_table_compiled.Rows[i][Pageno]);
                                double ml_len = Convert.ToDouble(Data_table_compiled.Rows[i][Rect_len]);
                                string dwg_name = Data_table_compiled.Rows[i][_AGEN_mainform.Col_dwg_name].ToString();
                                double strech1 = Convert.ToDouble(Data_table_compiled.Rows[i][stretch_val]);
                                double min_stretch_custom = Convert.ToDouble(Data_table_compiled.Rows[i]["min_stretch"]);



                                double Diff = Min_dist - strech1;
                                double Diff_from_custom = min_stretch_custom - strech1;

                                if (checkBox_ignore_match_with_plan_view.Checked == true)
                                {
                                    Diff = 0;
                                    Diff_from_custom = 0;
                                }


                                if (Diff_from_custom > 0)
                                {
                                    Data_table_compiled.Rows[i][stretch_val] = min_stretch_custom;
                                    for (int j = 0; j < Data_table_compiled.Rows.Count; ++j)
                                    {
                                        int Page2 = Convert.ToInt32(Data_table_compiled.Rows[j][Pageno]);
                                        double deltax2 = Convert.ToDouble(Data_table_compiled.Rows[j][DeltaX_col]);
                                        double band_len2 = Convert.ToDouble(Data_table_compiled.Rows[j][BandL]);

                                        if (Page1 == Page2)
                                        {
                                            Data_table_compiled.Rows[j][BandL] = band_len2 + Diff_from_custom;

                                            if (i < j)
                                            {
                                                Data_table_compiled.Rows[j][DeltaX_col] = deltax2 + Diff_from_custom;
                                            }
                                        }
                                    }
                                }
                                else if (Diff > 0)
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
                                        vp_vw1.AddVertexAt(0, new Point2d(cust_point0.X - custom_band_width / 2, cust_point0.Y - (Page1 - 1) * custom_band_separation), 0, 0, 0);
                                        vp_vw1.AddVertexAt(1, new Point2d(cust_point0.X + custom_band_width / 2, cust_point0.Y - (Page1 - 1) * custom_band_separation), 0, 0, 0);
                                        vp_vw1.AddVertexAt(2, new Point2d(cust_point0.X + custom_band_width / 2, cust_point0.Y - custom_band_height - (Page1 - 1) * custom_band_separation), 0, 0, 0);
                                        vp_vw1.AddVertexAt(3, new Point2d(cust_point0.X - custom_band_width / 2, cust_point0.Y - custom_band_height - (Page1 - 1) * custom_band_separation), 0, 0, 0);

                                        vp_vw1.Closed = true;
                                        vp_vw1.Layer = lnp;
                                        vp_vw1.ColorIndex = 3;
                                        BTrecord.AppendEntity(vp_vw1);
                                        Trans1.AddNewlyCreatedDBObject(vp_vw1, true);



                                        Polyline vp_vw2 = new Polyline();

                                        vp_vw2.AddVertexAt(0, new Point2d(cust_point0.X - ml_len / 2, cust_point0.Y - (Page1 - 1) * custom_band_separation), 0, 0, 0);
                                        vp_vw2.AddVertexAt(1, new Point2d(cust_point0.X + ml_len / 2, cust_point0.Y - (Page1 - 1) * custom_band_separation), 0, 0, 0);
                                        vp_vw2.AddVertexAt(2, new Point2d(cust_point0.X + ml_len / 2, cust_point0.Y - custom_band_height - (Page1 - 1) * custom_band_separation), 0, 0, 0);
                                        vp_vw2.AddVertexAt(3, new Point2d(cust_point0.X - ml_len / 2, cust_point0.Y - custom_band_height - (Page1 - 1) * custom_band_separation), 0, 0, 0);

                                        vp_vw2.Closed = true;
                                        vp_vw2.Layer = lnp;
                                        vp_vw2.ColorIndex = 1;
                                        BTrecord.AppendEntity(vp_vw2);
                                        Trans1.AddNewlyCreatedDBObject(vp_vw2, true);

                                        MText Band_label = new MText();
                                        Band_label.Contents = dwg_name;
                                        Band_label.Rotation = 0;
                                        Band_label.Attachment = AttachmentPoint.MiddleLeft;
                                        Band_label.Location = new Point3d(cust_point0.X - custom_band_width / 2, cust_point0.Y - custom_band_height / 2 - (Page1 - 1) * custom_band_separation, 0);
                                        Band_label.Layer = lnp;

                                        double gap1 = (_AGEN_mainform.Vw_width - ml_len * _AGEN_mainform.Vw_scale) / 2;
                                        Extents3d gerect = Band_label.GeometricExtents;
                                        Point3d p2 = gerect.MaxPoint;
                                        Point3d p1 = gerect.MinPoint;
                                        bool repeat1 = false;
                                        do
                                        {
                                            if (p2.X - p1.X > gap1 - custom_band_height / 3 && Band_label.TextHeight >= 2)
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



                            string visib1 = "";

                            for (int i = 0; i < Data_table_compiled.Rows.Count; ++i)
                            {

                                int Page1 = Convert.ToInt32(Data_table_compiled.Rows[i][Pageno]);

                                if (lista_bands_for_generation.Contains(Page1 - 1) == true)
                                {
                                    double Station1 = Convert.ToDouble(Data_table_compiled.Rows[i][Sta1]);
                                    double Station2 = Convert.ToDouble(Data_table_compiled.Rows[i][Sta2]);
                                    double M1 = Convert.ToDouble(Data_table_compiled.Rows[i][colm1]);
                                    double M2 = Convert.ToDouble(Data_table_compiled.Rows[i][colm2]);
                                    double Station1_CSF = Convert.ToDouble(Data_table_compiled.Rows[i][Sta1CSF]);
                                    double Station2_CSF = Convert.ToDouble(Data_table_compiled.Rows[i][Sta2CSF]);
                                    double M1_CSF = Convert.ToDouble(Data_table_compiled.Rows[i][colm1csf]);
                                    double M2_CSF = Convert.ToDouble(Data_table_compiled.Rows[i][colm2csf]);

                                    string val1 = Data_table_compiled.Rows[i][f1].ToString();
                                    string val2 = Data_table_compiled.Rows[i][f2].ToString();

                                    double ml_len = Convert.ToDouble(Data_table_compiled.Rows[i][Rect_len]);
                                    double band_len = Convert.ToDouble(Data_table_compiled.Rows[i][BandL]);
                                    double Diff = (band_len - ml_len) / 2;
                                    double deltax = Convert.ToDouble(Data_table_compiled.Rows[i][DeltaX_col]);

                                    string sta1_string = Functions.Get_chainage_from_double(Station1_CSF, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);
                                    string sta2_string = Functions.Get_chainage_from_double(Station2_CSF, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);

                                    string Suff = "'";
                                    if (_AGEN_mainform.units_of_measurement == "m") Suff = "";

                                    string len1 = Convert.ToString(Math.Round(Station2, _AGEN_mainform.round1) - Math.Round(Station1, _AGEN_mainform.round1)) + Suff;

                                    double strech1 = Convert.ToDouble(Data_table_compiled.Rows[i][stretch_val]);





                                    double x = cust_point0.X - lr * ml_len / 2 + lr * deltax;
                                    double y = cust_point0.Y - custom_band_height - (Page1 - 1) * custom_band_separation;
                                    Point3d InsPt = new Point3d(x, y, 0);


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



                                    System.Collections.Specialized.StringCollection Colectie_nume_atribute = new System.Collections.Specialized.StringCollection();
                                    System.Collections.Specialized.StringCollection Colectie_valori = new System.Collections.Specialized.StringCollection();

                                    Colectie_nume_atribute.Add(comboBox_custom_atr_sta1.Text);
                                    Colectie_valori.Add(sta1_string);

                                    Colectie_nume_atribute.Add(comboBox_custom_atr_sta2.Text);
                                    Colectie_valori.Add(sta2_string);

                                    Colectie_nume_atribute.Add(comboBox_custom_atr_sta1.Text + "1");
                                    Colectie_valori.Add(sta1_string);

                                    Colectie_nume_atribute.Add(comboBox_custom_atr_sta2.Text + "1");
                                    Colectie_valori.Add(sta2_string);

                                    Colectie_nume_atribute.Add(comboBox_custom_atr_field1.Text);
                                    Colectie_valori.Add(val1);

                                    Colectie_nume_atribute.Add(comboBox_custom_atr_field2.Text);
                                    Colectie_valori.Add(val2);

                                    Colectie_nume_atribute.Add(comboBox_custom_atr_distance.Text);
                                    Colectie_valori.Add(len1);

                                    BlockReference Block1 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "", Block_name, InsPt, 1 / _AGEN_mainform.custom_band_scale, 0,
                                                                                                                  lname1, Colectie_nume_atribute, Colectie_valori);



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

                                    Lista_val.Add(comboBox_custom_excel_name.Text);
                                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                    Functions.Populate_object_data_table_from_objectid(Tables1, Block1.ObjectId, "Agen_Custom", Lista_val, Lista_type);




                                }
                            }
                        }


                        if (_AGEN_mainform.Project_type == "3D" && poly3d.IsErased == false) poly3d.Erase();
                        Trans1.Commit();
                        _AGEN_mainform.custom_band_scale = 1;
                        custom_band_separation = 0;
                        write_custom_settings_to_excel(_AGEN_mainform.config_path);

                    }
                }
            }
            catch (System.Exception ex)
            {
                Functions.Transfer_datatable_to_new_excel_spreadsheet(Data_table_compiled);
                _AGEN_mainform.custom_band_scale = 1;
                custom_band_separation = 0;
                MessageBox.Show(ex.Message);
            }
            _AGEN_mainform.tpage_processing.Hide();

            set_enable_true();

            Ag.WindowState = FormWindowState.Normal;
        }

        public System.Data.DataTable Load_existing_custom_data(string File1)
        {

            if (System.IO.File.Exists(File1) == false)
            {
                MessageBox.Show("the custom band data file does not exist");
                return null;
            }


            System.Data.DataTable dt2 = new System.Data.DataTable();

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
                    bool is3d = false;
                    if (_AGEN_mainform.Project_type == "3D")
                    {
                        is3d = true;
                    }

                    dt2 = Functions.Build_Data_table_custom_from_excel(W1, _AGEN_mainform.Start_row_custom + 1, is3d);
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
            return dt2;

        }





        public void button_refresh_bands_Click(object sender, EventArgs e)
        {
            if (_AGEN_mainform.Data_Table_custom_bands != null)
            {
                if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count > 0)
                {
                    comboBox_custom_excel_name.Items.Clear();
                    comboBox_custom_excel_name.Items.Add("");

                    for (int i = 0; i < _AGEN_mainform.Data_Table_custom_bands.Rows.Count; ++i)
                    {
                        if (_AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"] != DBNull.Value)
                        {
                            string bn = Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"]);

                            comboBox_custom_excel_name.Items.Add(bn);


                        }
                    }

                    comboBox_custom_excel_name.SelectedIndex = 0;
                }
            }
        }

        private void comboBox_excel_name_SelectedIndexChanged(object sender, EventArgs e)
        {

            string band_name = comboBox_custom_excel_name.Text;

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
                                    if (_AGEN_mainform.dt_settings_custom.Rows[i][4] != DBNull.Value)
                                    {
                                        add_to_combobox(comboBox_custom_block, Convert.ToString(_AGEN_mainform.dt_settings_custom.Rows[i][4]));
                                    }
                                    if (_AGEN_mainform.dt_settings_custom.Rows[i][5] != DBNull.Value)
                                    {
                                        add_to_combobox(comboBox_custom_atr_sta1, Convert.ToString(_AGEN_mainform.dt_settings_custom.Rows[i][5]));
                                    }
                                    if (_AGEN_mainform.dt_settings_custom.Rows[i][6] != DBNull.Value)
                                    {
                                        add_to_combobox(comboBox_custom_atr_sta2, Convert.ToString(_AGEN_mainform.dt_settings_custom.Rows[i][6]));
                                    }
                                    if (_AGEN_mainform.dt_settings_custom.Rows[i][7] != DBNull.Value)
                                    {
                                        add_to_combobox(comboBox_custom_atr_distance, Convert.ToString(_AGEN_mainform.dt_settings_custom.Rows[i][7]));
                                    }
                                    if (_AGEN_mainform.dt_settings_custom.Rows[i][8] != DBNull.Value)
                                    {
                                        add_to_combobox(comboBox_custom_atr_field1, Convert.ToString(_AGEN_mainform.dt_settings_custom.Rows[i][8]));
                                    }
                                    if (_AGEN_mainform.dt_settings_custom.Rows[i][9] != DBNull.Value)
                                    {
                                        add_to_combobox(comboBox_custom_atr_field2, Convert.ToString(_AGEN_mainform.dt_settings_custom.Rows[i][9]));
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
            comboBox_custom_excel_name.Items.Clear();
            comboBox_custom_block.Items.Clear();
            comboBox_custom_atr_distance.Items.Clear();
            comboBox_custom_atr_field1.Items.Clear();
            comboBox_custom_atr_field2.Items.Clear();
            comboBox_custom_atr_sta1.Items.Clear();
            comboBox_custom_atr_sta2.Items.Clear();
        }

        private void button_read_band_to_xl_Click(object sender, EventArgs e)
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
                        BlockTable BlockTable1 = Trans1.GetObject(ThisDrawing.Database.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as BlockTable;

                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez.MessageForAdding = "\nSelect all objects:";
                        Prompt_rez.SingleOnly = false;
                        Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            Editor1.SetImpliedSelection(Empty_array);
                            Editor1.WriteMessage("\nCommand:");
                            set_enable_true();
                            return;
                        }

                        System.Data.DataTable dt1 = new System.Data.DataTable();
                        dt1.Columns.Add("blockname", typeof(string));
                        dt1.Columns.Add("sta1", typeof(double));
                        dt1.Columns.Add("sta2", typeof(double));
                        dt1.Columns.Add("desc", typeof(string));
                        dt1.Columns.Add("x", typeof(double));
                        dt1.Columns.Add("y", typeof(double));
                        dt1.Columns.Add("visibility", typeof(string));
                        dt1.Columns.Add("stretch", typeof(double));

                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {
                            BlockReference block1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as BlockReference;
                            if (block1 != null)
                            {
                                if (block1.AttributeCollection.Count > 0)
                                {


                                    string blockname = Functions.get_block_name(block1);

                                    dt1.Rows.Add();
                                    dt1.Rows[dt1.Rows.Count - 1][0] = blockname;
                                    dt1.Rows[dt1.Rows.Count - 1][4] = block1.Position.X;
                                    dt1.Rows[dt1.Rows.Count - 1][5] = block1.Position.Y;

                                    Autodesk.AutoCAD.DatabaseServices.AttributeCollection attColl = block1.AttributeCollection;

                                    foreach (ObjectId id1 in attColl)
                                    {
                                        DBObject ent = Trans1.GetObject(id1, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                                        if (ent is AttributeReference)
                                        {
                                            AttributeReference atr1 = ent as AttributeReference;

                                            if (atr1.Tag == "STA1")
                                            {
                                                string valoare = atr1.TextString;
                                                if (Functions.IsNumeric(valoare.Replace("+", "")) == true)
                                                {
                                                    dt1.Rows[dt1.Rows.Count - 1][1] = Convert.ToDouble(valoare.Replace("+", ""));
                                                }
                                            }


                                            else if (atr1.Tag == "STA2")
                                            {
                                                string valoare = atr1.TextString;
                                                if (Functions.IsNumeric(valoare.Replace("+", "")) == true)
                                                {
                                                    dt1.Rows[dt1.Rows.Count - 1][2] = Convert.ToDouble(valoare.Replace("+", ""));
                                                }
                                            }


                                            else
                                            {
                                                string valoare = atr1.TextString;
                                                dt1.Rows[dt1.Rows.Count - 1][3] = Convert.ToString(valoare);
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
                                                    dt1.Rows[dt1.Rows.Count - 1][6] = Convert.ToString(prop.Value);
                                                }
                                                if (prop.PropertyName == "Distance1")
                                                {
                                                    dt1.Rows[dt1.Rows.Count - 1][7] = Convert.ToDouble(prop.Value);
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                        }

                        Functions.Transfer_datatable_to_new_excel_spreadsheet(dt1);
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

        public string get_comboBox_custom_excel_name()
        {
            return comboBox_custom_excel_name.Text;
        }
        public void set_comboBox_custom_excel_name(string txt)
        {
            if (txt != "")
            {
                if (comboBox_custom_excel_name.Items.Contains(txt) == false)
                {
                    comboBox_custom_excel_name.Items.Add(txt);
                }
                comboBox_custom_excel_name.SelectedIndex = comboBox_custom_excel_name.Items.IndexOf(txt);
            }
        }
        public string get_comboBox_custom_block()
        {
            return comboBox_custom_block.Text;

        }
        public void set_comboBox_custom_block(string txt)
        {
            if (txt != "")
            {
                if (comboBox_custom_block.Items.Contains(txt) == false)
                {
                    comboBox_custom_block.Items.Add(txt);
                }
                comboBox_custom_block.SelectedIndex = comboBox_custom_block.Items.IndexOf(txt);
            }
        }
        public string get_comboBox_custom_atr_sta1()
        {
            return comboBox_custom_atr_sta1.Text;

        }
        public void set_comboBox_custom_atr_sta1(string txt)
        {
            if (txt != "")
            {
                if (comboBox_custom_atr_sta1.Items.Contains(txt) == false)
                {
                    comboBox_custom_atr_sta1.Items.Add(txt);
                }
                comboBox_custom_atr_sta1.SelectedIndex = comboBox_custom_atr_sta1.Items.IndexOf(txt);
            }
        }
        public string get_comboBox_custom_atr_sta2()
        {
            return comboBox_custom_atr_sta2.Text;

        }
        public void set_comboBox_custom_atr_sta2(string txt)
        {
            if (txt != "")
            {
                if (comboBox_custom_atr_sta2.Items.Contains(txt) == false)
                {
                    comboBox_custom_atr_sta2.Items.Add(txt);
                }
                comboBox_custom_atr_sta2.SelectedIndex = comboBox_custom_atr_sta2.Items.IndexOf(txt);
            }
        }
        public string get_comboBox_custom_atr_distance()
        {
            return comboBox_custom_atr_distance.Text;
        }
        public void set_comboBox_custom_atr_distance(string txt)
        {
            if (txt != "")
            {
                if (comboBox_custom_atr_distance.Items.Contains(txt) == false)
                {
                    comboBox_custom_atr_distance.Items.Add(txt);
                }
                comboBox_custom_atr_distance.SelectedIndex = comboBox_custom_atr_distance.Items.IndexOf(txt);
            }
        }
        public string get_comboBox_custom_atr_field1()
        {
            return comboBox_custom_atr_field1.Text;
        }
        public void set_comboBox_custom_atr_field1(string txt)
        {
            if (txt != "")
            {
                if (comboBox_custom_atr_field1.Items.Contains(txt) == false)
                {
                    comboBox_custom_atr_field1.Items.Add(txt);
                }
                comboBox_custom_atr_field1.SelectedIndex = comboBox_custom_atr_field1.Items.IndexOf(txt);
            }
        }
        public string get_comboBox_custom_atr_field2()
        {
            return comboBox_custom_atr_field2.Text;
        }
        public void set_comboBox_custom_atr_field2(string txt)
        {
            if (txt != "")
            {
                if (comboBox_custom_atr_field2.Items.Contains(txt) == false)
                {
                    comboBox_custom_atr_field2.Items.Add(txt);
                }
                comboBox_custom_atr_field2.SelectedIndex = comboBox_custom_atr_field2.Items.IndexOf(txt);
            }
        }

        public void write_custom_settings_to_excel(string cfg1)
        {

            string band_name = get_comboBox_custom_excel_name();
            string band_block = get_comboBox_custom_block();
            string atr_sta1 = get_comboBox_custom_atr_sta1();
            string atr_sta2 = get_comboBox_custom_atr_sta2();
            string atr_len = get_comboBox_custom_atr_distance();
            string atr_tag1 = get_comboBox_custom_atr_field1();
            string atr_tag2 = get_comboBox_custom_atr_field2();




            if (band_name != "" || band_block != "" || atr_sta1 != "" || atr_sta2 != "" || atr_len != "" || atr_tag1 != "" || atr_tag2 != "")
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



                foreach (Microsoft.Office.Interop.Excel.Worksheet wsh2 in Workbook1.Worksheets)
                {
                    if (wsh2.Name == comboBox_custom_excel_name.Text + "_cfg_" + segment1)
                    {
                        W1 = wsh2;
                    }
                }

                if (W1 == null)
                {
                    W1 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[Workbook1.Worksheets.Count], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                    if ((comboBox_custom_excel_name.Text + "_cfg_" + segment1).Length > 31)
                    {
                        MessageBox.Show(comboBox_custom_excel_name.Text + "_cfg_" + segment1 + "is bigger than 31 charcaters\r\nor you rename the custom band to have less characters\r\nor/and rename the segment");
                        return;
                    }
                    W1.Name = comboBox_custom_excel_name.Text + "_cfg_" + segment1;
                }


                try
                {
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

                    values[0, 1] = band_name;
                    values[4, 1] = band_block;
                    values[5, 1] = atr_sta1;
                    values[6, 1] = atr_sta2;
                    values[7, 1] = atr_len;
                    values[8, 1] = atr_tag1;
                    values[9, 1] = atr_tag2;


                    Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A1:B10"];
                    range1.Cells.NumberFormat = "General";
                    range1.Value2 = values;
                    Functions.Color_border_range_inside(range1, 0);

                    Workbook1.Save();

                    _AGEN_mainform.dt_settings_custom = null;

                    foreach (Microsoft.Office.Interop.Excel.Worksheet W3 in Workbook1.Worksheets)
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

        private void panel7_Click(object sender, EventArgs e)
        {
            if (checkBox_use_vw_scale.Visible == false)
            {
                checkBox_use_vw_scale.Visible = true;
                checkBox_ignore_match_with_plan_view.Visible = true;
                button_draw_rectangles.Visible = true;
                button_read_band_to_xl.Visible = true;
            }
            else
            {
                checkBox_use_vw_scale.Visible = false;
                checkBox_ignore_match_with_plan_view.Visible = false;

                button_draw_rectangles.Visible = false;
                button_read_band_to_xl.Visible = false;


            }
        }

        private void button_draw_rectangles_Click(object sender, EventArgs e)
        {

            string lnp = "Agen_no_plot_" + comboBox_custom_excel_name.Text.Replace(" ", "");
            int lr = 1;
            if (_AGEN_mainform.Left_to_Right == false) lr = -1;


            Functions.Kill_excel();

            string cfg1 = System.IO.Path.GetFileName(_AGEN_mainform.config_path);
            if (Functions.Get_if_workbook_is_open_in_Excel(cfg1) == true)
            {
                MessageBox.Show("Please close the " + cfg1 + " file");
                return;
            }






            _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;
            _AGEN_mainform.tpage_processing.Show();
            // Ag.WindowState = FormWindowState.Minimized;



            if (_AGEN_mainform.Vw_height <= 0)

            {
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


                set_enable_true();
                return;
            }

            Point3d cust_point0 = get_band_insertion_point_and_band_height(comboBox_custom_excel_name.Text);

            double custom_band_width = _AGEN_mainform.Vw_width;
            double custom_band_height = cust_point0.Z;
            cust_point0 = new Point3d(cust_point0.X, cust_point0.Y, 0);

            if (custom_band_height <= 0)
            {
                _AGEN_mainform.tpage_processing.Hide();
                MessageBox.Show("you did not specified the band height in the config file");
                Ag.WindowState = FormWindowState.Normal;
                set_enable_true();
                return;
            }


            set_enable_false();

            if (System.IO.Directory.Exists(_AGEN_mainform.tpage_setup.Get_project_database_folder()) == true)
            {
                string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                {
                    ProjF = ProjF + "\\";
                }

                string fisier_si = ProjF + _AGEN_mainform.sheet_index_excel_name;

                if (System.IO.File.Exists(fisier_si) == false)
                {
                    _AGEN_mainform.tpage_processing.Hide();
                    set_enable_true();
                    MessageBox.Show("the sheet index data file does not exist");
                    return;
                }



                _AGEN_mainform.dt_sheet_index = _AGEN_mainform.tpage_setup.Load_existing_sheet_index(fisier_si);




            }
            else
            {
                _AGEN_mainform.tpage_processing.Hide();
                set_enable_true();
                MessageBox.Show("the project folder does not exist");
                return;
            }


            if (_AGEN_mainform.dt_sheet_index.Rows.Count == 0)
            {
                _AGEN_mainform.tpage_processing.Hide();
                set_enable_true();
                MessageBox.Show("the sheet index file does not have any data");
                return;
            }

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


                        if (_AGEN_mainform.dt_sheet_index == null || _AGEN_mainform.dt_sheet_index.Rows.Count == 0)
                        {
                            _AGEN_mainform.tpage_processing.Hide();
                            set_enable_true();
                            MessageBox.Show("the data of sheet index file is not complete");
                            return;
                        }


                        double scale_cust = 1;
                        if (checkBox_use_vw_scale.Checked == true) scale_cust = _AGEN_mainform.Vw_scale;
                        Functions.Creaza_layer(lnp, 30, false);

                        for (int i = 0; i < _AGEN_mainform.dt_sheet_index.Rows.Count; ++i)
                        {
                            string dwg_name = _AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_dwg_name].ToString();
                            Polyline vp_vw1 = new Polyline();
                            vp_vw1.AddVertexAt(0, new Point2d(cust_point0.X - custom_band_width / (2 * scale_cust), cust_point0.Y - i * _AGEN_mainform.Band_Separation / scale_cust), 0, 0, 0);///top
                            vp_vw1.AddVertexAt(1, new Point2d(cust_point0.X + custom_band_width / (2 * scale_cust), cust_point0.Y - i * _AGEN_mainform.Band_Separation / scale_cust), 0, 0, 0);
                            vp_vw1.AddVertexAt(2, new Point2d(cust_point0.X + custom_band_width / (2 * scale_cust), cust_point0.Y - i * _AGEN_mainform.Band_Separation / scale_cust - custom_band_height / scale_cust), 0, 0, 0);//bottom
                            vp_vw1.AddVertexAt(3, new Point2d(cust_point0.X - custom_band_width / (2 * scale_cust), cust_point0.Y - i * _AGEN_mainform.Band_Separation / scale_cust - custom_band_height / scale_cust), 0, 0, 0);

                            vp_vw1.Closed = true;
                            vp_vw1.Layer = lnp;
                            vp_vw1.ColorIndex = 3;
                            BTrecord.AppendEntity(vp_vw1);
                            Trans1.AddNewlyCreatedDBObject(vp_vw1, true);

                            MText Band_label = new MText();
                            Band_label.Contents = dwg_name;
                            Band_label.Rotation = 0;
                            if (lr == 1)
                            {
                                Band_label.Attachment = AttachmentPoint.BottomLeft;
                            }
                            else
                            {
                                Band_label.Attachment = AttachmentPoint.BottomRight;
                            }

                            Band_label.Location = new Point3d(cust_point0.X - lr * custom_band_width / (2 * scale_cust), cust_point0.Y - i * _AGEN_mainform.Band_Separation / scale_cust, 0);
                            Band_label.Layer = lnp;
                            Band_label.TextHeight = 2;

                            BTrecord.AppendEntity(Band_label);
                            Trans1.AddNewlyCreatedDBObject(Band_label, true);
                        }

                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            _AGEN_mainform.tpage_processing.Hide();

            set_enable_true();

            Ag.WindowState = FormWindowState.Normal;
        }

        private void panel7_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
