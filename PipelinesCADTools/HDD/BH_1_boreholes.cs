﻿using System;
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
    public partial class HDD_Boreholes : Form
    {



        System.Data.DataTable dt_point;
        System.Data.DataTable dt_rock;
        System.Data.DataTable dt_soil;
        System.Data.DataTable dt_core_s;
        System.Data.DataTable dt_core_r;


        string um = "f";


        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();

            lista_butoane.Add(button_load_borehole_spreadsheet);
            lista_butoane.Add(button_all_pts_l);
            lista_butoane.Add(button_all_pts_nl);

            lista_butoane.Add(button_draw_borehole);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = false;
            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_load_borehole_spreadsheet);
            lista_butoane.Add(button_all_pts_l);
            lista_butoane.Add(button_all_pts_nl);

            lista_butoane.Add(button_draw_borehole);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }

        public HDD_Boreholes()
        {
            InitializeComponent();
            comboBox_scales.SelectedItem = "1:60";
            //comboBox_scales_plan_view.SelectedItem = "1:100";

        }

        private void button_load_borehole_info_Click(object sender, EventArgs e)
        {
            string col_north = "North";
            string col_east = "East";
            string col_lat = "Lat";
            string col_long = "Long";

            try
            {
                if (comboBox_xl1.Text != "")
                {
                    string file1 = comboBox_xl1.Text;
                    if (file1.Length > 0)
                    {
                        set_enable_false();

                        Microsoft.Office.Interop.Excel.Workbook Workbook1 = Functions.Get_opened_workbook_from_Excel(file1);
                        Microsoft.Office.Interop.Excel.Worksheet W1 = null;
                        Microsoft.Office.Interop.Excel.Worksheet W2 = null;
                        Microsoft.Office.Interop.Excel.Worksheet W3 = null;
                        Microsoft.Office.Interop.Excel.Worksheet W4 = null;
                        Microsoft.Office.Interop.Excel.Worksheet W5 = null;

                        if (Workbook1 != null)
                        {
                            for (int i = 1; i <= Workbook1.Worksheets.Count; ++i)
                            {
                                Microsoft.Office.Interop.Excel.Worksheet Wx = Workbook1.Worksheets[i];
                                if (Wx.Name == "POINT")
                                {
                                    W1 = Wx;
                                }
                                if (Wx.Name == "CORE SAMPLE")
                                {
                                    W2 = Wx;
                                }
                                if (Wx.Name == "ROCK LITHOLOGY")
                                {
                                    W3 = Wx;
                                }
                                if (Wx.Name == "SOIL LITHOLOGY")
                                {
                                    W4 = Wx;
                                }
                                if (Wx.Name == "SOIL SAMPLE")
                                {
                                    W5 = Wx;
                                }
                            }
                        }



                        if (W1 != null)
                        {
                            dt_point = null;
                            dt_rock = null;
                            dt_soil = null;
                            dt_core_r = null;
                            dt_core_s = null;
                            build_data_tables(W1, W2, W3, W4, W5);
                            comboBox1.Items.Clear();
                            comboBox2.Items.Clear();

                            //Functions.Transfer_datatable_to_new_excel_spreadsheet(dt_soil);


                            if (dt_point != null && dt_point.Rows.Count > 0 && ((dt_soil != null && dt_soil.Rows.Count > 0) || (dt_rock != null && dt_rock.Rows.Count > 0)))
                            {
                                ObjectId[] Empty_array = null;
                                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;

                                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                                {
                                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                                    {
                                        Autodesk.Gis.Map.Platform.AcMapMap Acmap = Autodesk.Gis.Map.Platform.AcMapMap.GetCurrentMap();
                                        string string_current = Acmap.GetMapSRS();
                                        if (string.IsNullOrEmpty(string_current) == true)
                                        {
                                            MessageBox.Show("Please set your coordinate system");
                                            Editor1.SetImpliedSelection(Empty_array);
                                            Editor1.WriteMessage("\nCommand:");
                                            set_enable_true();
                                            return;
                                        }


                                        OSGeo.MapGuide.MgCoordinateSystemFactory Coord_factory1 = new OSGeo.MapGuide.MgCoordinateSystemFactory();
                                        OSGeo.MapGuide.MgCoordinateSystemCatalog Catalog1 = Coord_factory1.GetCatalog();
                                        OSGeo.MapGuide.MgCoordinateSystemDictionary Dictionary1 = Catalog1.GetCoordinateSystemDictionary();
                                        OSGeo.MapGuide.MgCoordinateSystemEnum Enum1 = Dictionary1.GetEnum();

                                        int count1 = Dictionary1.GetSize();
                                        OSGeo.MapGuide.MgStringCollection Colectie_names_cs = Enum1.NextName(count1);
                                        string coord_sys_name = null;

                                        OSGeo.MapGuide.MgCoordinateSystem coord_sys = null;

                                        int coord_sys_type = 0;

                                        for (int i = 0; i < count1; i++)

                                        {

                                            coord_sys_name = Colectie_names_cs.GetItem(i);

                                            coord_sys = Dictionary1.GetCoordinateSystem(coord_sys_name);
                                            string code1 = coord_sys.GetCsCode();

                                            comboBox1.Items.Add(code1);
                                            comboBox2.Items.Add(code1);

                                            #region learning
                                            coord_sys_type = coord_sys.GetType();

                                            if (coord_sys_type == OSGeo.MapGuide.MgCoordinateSystemType.Arbitrary)

                                            {



                                            }

                                            else if (coord_sys_type == OSGeo.MapGuide.MgCoordinateSystemType.Geographic)

                                            {


                                            }

                                            else if (coord_sys_type == OSGeo.MapGuide.MgCoordinateSystemType.Projected)

                                            {



                                            }
                                            #endregion

                                        }

                                        OSGeo.MapGuide.MgCoordinateSystem current_system = Coord_factory1.Create(string_current);
                                        if (dt_point.Rows[0][col_lat] != DBNull.Value)
                                        {
                                            if (comboBox1.Items.Contains("LL84") == true)
                                            {
                                                comboBox1.SelectedIndex = comboBox1.Items.IndexOf("LL84");
                                            }
                                        }
                                        else
                                        {
                                            if (comboBox1.Items.Contains(current_system.GetCsCode()) == true)
                                            {
                                                comboBox1.SelectedIndex = comboBox1.Items.IndexOf(current_system.GetCsCode());
                                            }
                                        }
                                        if (comboBox2.Items.Contains(current_system.GetCsCode()) == true)
                                        {
                                            comboBox2.SelectedIndex = comboBox1.Items.IndexOf(current_system.GetCsCode());
                                        }

                                    }
                                }
                                Editor1.SetImpliedSelection(Empty_array);
                                Editor1.WriteMessage("\nCommand:");

                            }
                        }
                    }

                }
            }

            catch (System.Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            set_enable_true();
        }




        private void build_data_tables(Microsoft.Office.Interop.Excel.Worksheet W1,
            Microsoft.Office.Interop.Excel.Worksheet W2, Microsoft.Office.Interop.Excel.Worksheet W3,
            Microsoft.Office.Interop.Excel.Worksheet W4, Microsoft.Office.Interop.Excel.Worksheet W5)
        {
            string units1 = "Feet";
            label_um.Visible = false;
            object[,] values1 = new object[300000, 53];
            Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A1:BA30000"];
            values1 = range1.Value2;
            System.Data.DataTable dt1 = new System.Data.DataTable();
            dt1.Columns.Add("A", typeof(string));
            dt1.Columns.Add("B", typeof(string));
            dt1.Columns.Add("C", typeof(string));
            dt1.Columns.Add("D", typeof(string));
            dt1.Columns.Add("E", typeof(string));
            dt1.Columns.Add("F", typeof(string));
            dt1.Columns.Add("G", typeof(string));
            dt1.Columns.Add("H", typeof(string));
            dt1.Columns.Add("I", typeof(string));
            dt1.Columns.Add("J", typeof(string));
            dt1.Columns.Add("K", typeof(string));
            dt1.Columns.Add("L", typeof(string));
            dt1.Columns.Add("M", typeof(string));
            dt1.Columns.Add("N", typeof(string));
            dt1.Columns.Add("O", typeof(string));
            dt1.Columns.Add("P", typeof(string));
            dt1.Columns.Add("Q", typeof(string));
            dt1.Columns.Add("R", typeof(string));
            dt1.Columns.Add("S", typeof(string));
            dt1.Columns.Add("T", typeof(string));
            dt1.Columns.Add("U", typeof(string));
            dt1.Columns.Add("V", typeof(string));
            dt1.Columns.Add("W", typeof(string));
            dt1.Columns.Add("X", typeof(string));
            dt1.Columns.Add("Y", typeof(string));
            dt1.Columns.Add("Z", typeof(string));
            dt1.Columns.Add("AA", typeof(string));
            dt1.Columns.Add("AB", typeof(string));
            dt1.Columns.Add("AC", typeof(string));
            dt1.Columns.Add("AD", typeof(string));
            dt1.Columns.Add("AE", typeof(string));
            dt1.Columns.Add("AF", typeof(string));
            dt1.Columns.Add("AG", typeof(string));
            dt1.Columns.Add("AH", typeof(string));
            dt1.Columns.Add("AI", typeof(string));
            dt1.Columns.Add("AJ", typeof(string));
            dt1.Columns.Add("AK", typeof(string));
            dt1.Columns.Add("AL", typeof(string));
            dt1.Columns.Add("AM", typeof(string));
            dt1.Columns.Add("AN", typeof(string));
            dt1.Columns.Add("AO", typeof(string));
            dt1.Columns.Add("AP", typeof(string));
            dt1.Columns.Add("AQ", typeof(string));
            dt1.Columns.Add("AR", typeof(string));
            dt1.Columns.Add("AS", typeof(string));
            dt1.Columns.Add("AT", typeof(string));
            dt1.Columns.Add("AU", typeof(string));
            dt1.Columns.Add("AV", typeof(string));
            dt1.Columns.Add("AW", typeof(string));
            dt1.Columns.Add("AX", typeof(string));
            dt1.Columns.Add("AY", typeof(string));
            dt1.Columns.Add("AZ", typeof(string));
            dt1.Columns.Add("BA", typeof(string));

            for (int i = 1; i <= 30000; ++i)
            {
                dt1.Rows.Add();
                for (int j = 1; j <= 53; ++j)
                {
                    object Valoare1 = values1[i, j];
                    if (Valoare1 != null)
                    {
                        dt1.Rows[i - 1][j - 1] = Valoare1;
                    }
                    else
                    {
                        if (j == 1)
                        {
                            dt1.Rows[dt1.Rows.Count - 1].Delete();
                            i = 30001;
                            j = 54;
                        }
                    }
                }
            }

            if (dt1.Rows.Count > 1)
            {

                string col_pt = "PointID";
                string col_hole = "HoleDepth";
                string col_elev = "Elevation";
                string col_north = "North";
                string col_east = "East";
                string col_lat = "Lat";
                string col_long = "Long";


                string A = "A";
                string B = "B";
                string C = "C";
                string M = "M";
                string N = "N";
                string O = "O";
                string P = "P";


                dt_point = new System.Data.DataTable();
                dt_point.Columns.Add(col_pt, typeof(string));
                dt_point.Columns.Add(col_hole, typeof(double));
                dt_point.Columns.Add(col_elev, typeof(double));
                dt_point.Columns.Add(col_north, typeof(double));
                dt_point.Columns.Add(col_east, typeof(double));
                dt_point.Columns.Add(col_lat, typeof(double));
                dt_point.Columns.Add(col_long, typeof(double));
                for (int i = 1; i < dt1.Rows.Count; i++)
                {
                    if (dt1.Rows[i][A] != DBNull.Value && dt1.Rows[i][B] != DBNull.Value &&
                        dt1.Rows[i][C] != DBNull.Value &&
                        ((dt1.Rows[i][M] != DBNull.Value && dt1.Rows[i][N] != DBNull.Value) ||
                        (dt1.Rows[i][O] != DBNull.Value && dt1.Rows[i][P] != DBNull.Value)) &&
                        Functions.IsNumeric(Convert.ToString(dt1.Rows[i][B]).Replace(" ", "")) == true &&
                        Functions.IsNumeric(Convert.ToString(dt1.Rows[i][C]).Replace(" ", "")) == true &&
                       ((Functions.IsNumeric(Convert.ToString(dt1.Rows[i][M]).Replace(" ", "")) == true &&
                        Functions.IsNumeric(Convert.ToString(dt1.Rows[i][N]).Replace(" ", "")) == true) ||
                       (Functions.IsNumeric(Convert.ToString(dt1.Rows[i][O]).Replace(" ", "")) == true &&
                        Functions.IsNumeric(Convert.ToString(dt1.Rows[i][P]).Replace(" ", "")) == true)))
                    {
                        dt_point.Rows.Add();
                        dt_point.Rows[dt_point.Rows.Count - 1][col_pt] = dt1.Rows[i][A];
                        dt_point.Rows[dt_point.Rows.Count - 1][col_hole] = Convert.ToDouble(Convert.ToString(dt1.Rows[i][B]).Replace(" ", ""));
                        dt_point.Rows[dt_point.Rows.Count - 1][col_elev] = Convert.ToDouble(Convert.ToString(dt1.Rows[i][C]).Replace(" ", ""));
                        if (dt1.Rows[i][M] != DBNull.Value && dt1.Rows[i][N] != DBNull.Value)
                        {
                            dt_point.Rows[dt_point.Rows.Count - 1][col_north] = Convert.ToDouble(Convert.ToString(dt1.Rows[i][M]).Replace(" ", ""));
                            dt_point.Rows[dt_point.Rows.Count - 1][col_east] = Convert.ToDouble(Convert.ToString(dt1.Rows[i][N]).Replace(" ", ""));
                        }

                        if (dt1.Rows[i][O] != DBNull.Value && dt1.Rows[i][P] != DBNull.Value)
                        {
                            dt_point.Rows[dt_point.Rows.Count - 1][col_lat] = Convert.ToDouble(Convert.ToString(dt1.Rows[i][O]).Replace(" ", ""));
                            dt_point.Rows[dt_point.Rows.Count - 1][col_long] = Convert.ToDouble(Convert.ToString(dt1.Rows[i][P]).Replace(" ", ""));
                        }

                    }
                }
            }

            if (W2 != null)
            {
                object[,] values2 = new object[300000, 53];
                Microsoft.Office.Interop.Excel.Range range2 = W2.Range["A1:BA30000"];
                values2 = range2.Value2;
                System.Data.DataTable dt2 = dt1.Clone();
                for (int i = 1; i <= 30000; ++i)
                {
                    dt2.Rows.Add();
                    for (int j = 1; j <= 53; ++j)
                    {
                        object Valoare2 = values2[i, j];
                        if (Valoare2 != null)
                        {
                            dt2.Rows[i - 1][j - 1] = Valoare2;
                        }
                        else
                        {
                            if (j == 1)
                            {
                                dt2.Rows[dt2.Rows.Count - 1].Delete();
                                i = 30001;
                                j = 54;
                            }
                        }
                    }
                }

                if (dt2.Rows.Count > 1)
                {
                    string col_pt = "PointID";
                    string col_depth = "Depth";
                    string col_len = "Length";
                    string col_type = "Type";
                    string col_number = "Number";
                    string col_recovery = "Recovery";
                    string col_rqd = "RQD";

                    string A = "A";
                    string B = "B";
                    string C = "C";
                    string D = "D";
                    string E = "E";
                    string G = "G";
                    string H = "H";

                    dt_core_r = new System.Data.DataTable();
                    dt_core_r.Columns.Add(col_pt, typeof(string));
                    dt_core_r.Columns.Add(col_depth, typeof(double));
                    dt_core_r.Columns.Add(col_len, typeof(double));
                    dt_core_r.Columns.Add(col_type, typeof(string));
                    dt_core_r.Columns.Add(col_number, typeof(string));
                    dt_core_r.Columns.Add(col_recovery, typeof(double));
                    dt_core_r.Columns.Add(col_rqd, typeof(double));

                    for (int i = 1; i < dt2.Rows.Count; i++)
                    {
                        if (dt2.Rows[i][A] != DBNull.Value &&
                            dt2.Rows[i][B] != DBNull.Value &&
                            dt2.Rows[i][C] != DBNull.Value &&
                            dt2.Rows[i][D] != DBNull.Value &&
                            dt2.Rows[i][E] != DBNull.Value &&
                            dt2.Rows[i][G] != DBNull.Value &&
                            dt2.Rows[i][H] != DBNull.Value &&
                            Functions.IsNumeric(Convert.ToString(dt2.Rows[i][B]).Replace(" ", "")) == true &&
                            Functions.IsNumeric(Convert.ToString(dt2.Rows[i][C]).Replace(" ", "")) == true &&
                            Functions.IsNumeric(Convert.ToString(dt2.Rows[i][G]).Replace(" ", "")) == true &&
                            Functions.IsNumeric(Convert.ToString(dt2.Rows[i][H]).Replace(" ", "")) == true)
                        {
                            dt_core_r.Rows.Add();
                            dt_core_r.Rows[dt_core_r.Rows.Count - 1][col_pt] = dt2.Rows[i][A];
                            dt_core_r.Rows[dt_core_r.Rows.Count - 1][col_depth] = Convert.ToDouble(Convert.ToString(dt2.Rows[i][B]).Replace(" ", ""));
                            dt_core_r.Rows[dt_core_r.Rows.Count - 1][col_len] = Convert.ToDouble(Convert.ToString(dt2.Rows[i][C]).Replace(" ", ""));
                            dt_core_r.Rows[dt_core_r.Rows.Count - 1][col_type] = dt2.Rows[i][D];
                            dt_core_r.Rows[dt_core_r.Rows.Count - 1][col_number] = dt2.Rows[i][E];
                            dt_core_r.Rows[dt_core_r.Rows.Count - 1][col_recovery] = Convert.ToDouble(Convert.ToString(dt2.Rows[i][G]).Replace(" ", ""));
                            dt_core_r.Rows[dt_core_r.Rows.Count - 1][col_rqd] = Convert.ToDouble(Convert.ToString(dt2.Rows[i][H]).Replace(" ", ""));
                        }
                    }
                }
            }

            if (W3 != null)
            {
                object[,] values3 = new object[300000, 53];
                Microsoft.Office.Interop.Excel.Range range3 = W3.Range["A1:BA30000"];
                values3 = range3.Value2;
                System.Data.DataTable dt3 = dt1.Clone();
                for (int i = 1; i <= 30000; ++i)
                {
                    dt3.Rows.Add();
                    for (int j = 1; j <= 53; ++j)
                    {
                        object Valoare3 = values3[i, j];
                        if (Valoare3 != null)
                        {
                            dt3.Rows[i - 1][j - 1] = Valoare3;
                        }
                        else
                        {
                            if (j == 1)
                            {
                                dt3.Rows[dt3.Rows.Count - 1].Delete();
                                i = 30001;
                                j = 54;
                            }
                        }
                    }
                }
                if (dt3.Rows.Count > 1)
                {
                    string col_pt = "PointID";
                    string col_depth = "Depth";
                    string col_bottom = "Bottom";
                    string col_graphic = "Graphic";


                    string A = "A";
                    string B = "B";
                    string C = "C";
                    string D = "D";


                    dt_rock = new System.Data.DataTable();
                    dt_rock.Columns.Add(col_pt, typeof(string));
                    dt_rock.Columns.Add(col_depth, typeof(double));
                    dt_rock.Columns.Add(col_bottom, typeof(double));
                    dt_rock.Columns.Add(col_graphic, typeof(string));


                    for (int i = 1; i < dt3.Rows.Count; i++)
                    {
                        if (dt3.Rows[i][A] != DBNull.Value &&
                            dt3.Rows[i][B] != DBNull.Value &&
                            dt3.Rows[i][C] != DBNull.Value &&
                            dt3.Rows[i][D] != DBNull.Value &&
                            Functions.IsNumeric(Convert.ToString(dt3.Rows[i][B]).Replace(" ", "")) == true &&
                            Functions.IsNumeric(Convert.ToString(dt3.Rows[i][C]).Replace(" ", "")) == true)
                        {

                            dt_rock.Rows.Add();
                            dt_rock.Rows[dt_rock.Rows.Count - 1][col_pt] = dt3.Rows[i][A];
                            dt_rock.Rows[dt_rock.Rows.Count - 1][col_depth] = Convert.ToDouble(Convert.ToString(dt3.Rows[i][B]).Replace(" ", ""));
                            dt_rock.Rows[dt_rock.Rows.Count - 1][col_bottom] = Convert.ToDouble(Convert.ToString(dt3.Rows[i][C]).Replace(" ", ""));
                            dt_rock.Rows[dt_rock.Rows.Count - 1][col_graphic] = dt3.Rows[i][D];
                        }
                    }
                }
            }

            if (W4 != null)
            {


                object[,] values4 = new object[300000, 53];
                Microsoft.Office.Interop.Excel.Range range4 = W4.Range["A1:BA30000"];
                values4 = range4.Value2;
                System.Data.DataTable dt4 = dt1.Clone();

                for (int i = 1; i <= 30000; ++i)
                {
                    dt4.Rows.Add();
                    for (int j = 1; j <= 53; ++j)
                    {
                        object Valoare4 = values4[i, j];
                        if (Valoare4 != null)
                        {
                            dt4.Rows[i - 1][j - 1] = Valoare4;
                        }
                        else
                        {
                            if (j == 1)
                            {
                                dt4.Rows[dt4.Rows.Count - 1].Delete();
                                i = 30001;
                                j = 54;
                            }
                        }
                    }
                }


                if (dt4.Rows.Count > 1)
                {
                    string col_pt = "PointID";
                    string col_depth = "Depth";
                    string col_bottom = "Bottom";
                    string col_graphic = "Graphic";


                    string A = "A";
                    string B = "B";
                    string C = "C";
                    string D = "D";


                    dt_soil = new System.Data.DataTable();
                    dt_soil.Columns.Add(col_pt, typeof(string));
                    dt_soil.Columns.Add(col_depth, typeof(double));
                    dt_soil.Columns.Add(col_bottom, typeof(double));
                    dt_soil.Columns.Add(col_graphic, typeof(string));


                    for (int i = 1; i < dt4.Rows.Count; i++)
                    {
                        if (dt4.Rows[i][A] != DBNull.Value &&
                            dt4.Rows[i][B] != DBNull.Value &&
                            dt4.Rows[i][C] != DBNull.Value &&
                            dt4.Rows[i][D] != DBNull.Value &&
                            Functions.IsNumeric(Convert.ToString(dt4.Rows[i][B]).Replace(" ", "")) == true &&
                            Functions.IsNumeric(Convert.ToString(dt4.Rows[i][C]).Replace(" ", "")) == true)
                        {

                            dt_soil.Rows.Add();
                            dt_soil.Rows[dt_soil.Rows.Count - 1][col_pt] = dt4.Rows[i][A];
                            dt_soil.Rows[dt_soil.Rows.Count - 1][col_depth] = Convert.ToDouble(Convert.ToString(dt4.Rows[i][B]).Replace(" ", ""));
                            dt_soil.Rows[dt_soil.Rows.Count - 1][col_bottom] = Convert.ToDouble(Convert.ToString(dt4.Rows[i][C]).Replace(" ", ""));
                            dt_soil.Rows[dt_soil.Rows.Count - 1][col_graphic] = dt4.Rows[i][D];
                        }
                    }
                }
            }

            if (W5 != null)
            {
                object[,] values5 = new object[300000, 53];
                Microsoft.Office.Interop.Excel.Range range5 = W5.Range["A1:BA30000"];
                values5 = range5.Value2;
                System.Data.DataTable dt5 = dt1.Clone();

                for (int i = 1; i <= 30000; ++i)
                {
                    dt5.Rows.Add();
                    for (int j = 1; j <= 53; ++j)
                    {
                        object Valoare5 = values5[i, j];
                        if (Valoare5 != null)
                        {
                            dt5.Rows[i - 1][j - 1] = Valoare5;
                        }
                        else
                        {
                            if (j == 1)
                            {
                                dt5.Rows[dt5.Rows.Count - 1].Delete();
                                i = 30001;
                                j = 55;
                            }
                        }
                    }
                }

                if (dt5.Rows.Count > 1)
                {


                    string col_pt = "PointID";
                    string col_depth = "Depth";
                    string col_len = "Length";
                    string col_type = "Type";
                    string col_number = "Number";

                    string col_blow1 = "Blows 1st 6in";
                    string col_blow2 = "Blows 2nd 6in";
                    string col_blow3 = "Blows 3rd 6in";
                    string col_blow4 = "Blows 4th 6in";


                    string A = "A";
                    string B = "B";
                    string C = "C";

                    string D = "D";
                    string E = "E";
                    string F = "F";
                    string G = "G";
                    string H = "H";
                    string I = "I";


                    dt_core_s = new System.Data.DataTable();
                    dt_core_s.Columns.Add(col_pt, typeof(string));
                    dt_core_s.Columns.Add(col_depth, typeof(double));
                    dt_core_s.Columns.Add(col_len, typeof(double));
                    dt_core_s.Columns.Add(col_type, typeof(string));
                    dt_core_s.Columns.Add(col_number, typeof(string));
                    dt_core_s.Columns.Add(col_blow1, typeof(string));
                    dt_core_s.Columns.Add(col_blow2, typeof(string));
                    dt_core_s.Columns.Add(col_blow3, typeof(string));
                    dt_core_s.Columns.Add(col_blow4, typeof(string));


                    for (int i = 1; i < dt5.Rows.Count; i++)
                    {
                        if (dt5.Rows[i][A] != DBNull.Value &&
                            dt5.Rows[i][B] != DBNull.Value &&
                            dt5.Rows[i][C] != DBNull.Value &&
                            dt5.Rows[i][D] != DBNull.Value &&
                            dt5.Rows[i][E] != DBNull.Value &&
                            dt5.Rows[i][F] != DBNull.Value &&
                            dt5.Rows[i][G] != DBNull.Value &&
                            dt5.Rows[i][H] != DBNull.Value &&
                            dt5.Rows[i][I] != DBNull.Value &&
                            Functions.IsNumeric(Convert.ToString(dt5.Rows[i][B]).Replace(" ", "")) == true &&
                            Functions.IsNumeric(Convert.ToString(dt5.Rows[i][C]).Replace(" ", "")) == true)
                        {

                            dt_core_s.Rows.Add();
                            dt_core_s.Rows[dt_core_s.Rows.Count - 1][col_pt] = dt5.Rows[i][A];
                            dt_core_s.Rows[dt_core_s.Rows.Count - 1][col_depth] = Convert.ToDouble(Convert.ToString(dt5.Rows[i][B]).Replace(" ", ""));
                            dt_core_s.Rows[dt_core_s.Rows.Count - 1][col_len] = Convert.ToDouble(Convert.ToString(dt5.Rows[i][C]).Replace(" ", ""));
                            dt_core_s.Rows[dt_core_s.Rows.Count - 1][col_type] = dt5.Rows[i][D];
                            dt_core_s.Rows[dt_core_s.Rows.Count - 1][col_number] = dt5.Rows[i][E];
                            dt_core_s.Rows[dt_core_s.Rows.Count - 1][col_blow1] = dt5.Rows[i][F];
                            dt_core_s.Rows[dt_core_s.Rows.Count - 1][col_blow2] = dt5.Rows[i][G];
                            dt_core_s.Rows[dt_core_s.Rows.Count - 1][col_blow3] = dt5.Rows[i][H];
                            dt_core_s.Rows[dt_core_s.Rows.Count - 1][col_blow4] = dt5.Rows[i][I];

                        }
                    }
                }
            }
            label_um.Visible = true;
            label_um.Text = units1;


            // Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt_point);
            //Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt_rock);
            // Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt_soil);
            // Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt_core_s);
            // Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt_core_r);
        }






        private void combo_load_excel_dropdown(object sender, EventArgs e)
        {
            Functions.Load_opened_workbooks_to_combobox(comboBox_xl1);
        }



        private void button_draw_borehole_Click(object sender, EventArgs e)
        {
            try
            {
                if (dt_point != null && dt_point.Rows.Count > 0 && ((dt_soil != null && dt_soil.Rows.Count > 0) || (dt_rock != null && dt_rock.Rows.Count > 0)))
                {
                    this.MdiParent.WindowState = FormWindowState.Minimized;
                    set_enable_false();
                    ObjectId[] Empty_array = null;
                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                    Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                    Editor1.SetImpliedSelection(Empty_array);

                    double Texth = 0.08;
                    double Texth_plan_view = 0.08;

                    if (Functions.IsNumeric(textBox_th.Text) == true)
                    {
                        Texth = Convert.ToDouble(textBox_th.Text);
                    }

                    double scale1 = 1;

                    if (comboBox_scales.Text.Length > 0)
                    {
                        string sc = comboBox_scales.Text.Replace("1:", "");
                        if (Functions.IsNumeric(sc) == true)
                        {
                            scale1 = Convert.ToDouble(sc);
                        }
                    }

                    double scale2 = scale1;

                    Texth = Texth * scale1;
                    Texth_plan_view = Texth_plan_view * scale2;

                    double stick_vexag = 1;


                    List<string> lista_legend = new List<string>();


                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {

                            BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                            LayerTable LayerTable1 = Trans1.GetObject(ThisDrawing.Database.LayerTableId, OpenMode.ForRead) as LayerTable;
                            TextStyleTable Text_style_table = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                            DBDictionary Leader_style_table = Trans1.GetObject(ThisDrawing.Database.MLeaderStyleDictionaryId, OpenMode.ForRead) as DBDictionary;
                            MLeaderStyle HDD_mleader = new MLeaderStyle();
                            TextStyleTableRecord HDD_textstyle = null;

                            ObjectId Arrowid_dot = Functions.Get_Arrow_dimension_ID("DIMBLK2", "_DotSmall");


                            string mleaderstyle_name = "HDD_borehole";
                            string textstyle_name = "HDD_borehole";



                            string f = "'";
                            string f1 = "ft.";
                            if (um != "f")
                            {
                                f = "";
                                f1 = "m";
                            }
                            foreach (ObjectId TextStyle_id in Text_style_table)
                            {
                                TextStyleTableRecord TextStyle1 = Trans1.GetObject(TextStyle_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as TextStyleTableRecord;
                                if (TextStyle1.Name.ToLower() == textstyle_name.ToLower())
                                {
                                    HDD_textstyle = TextStyle1;
                                    //if (TextStyle1.TextSize > 0) Texth = TextStyle1.TextSize;
                                }
                            }

                            if (HDD_textstyle == null)
                            {
                                Text_style_table.UpgradeOpen();
                                HDD_textstyle = new TextStyleTableRecord();
                                HDD_textstyle.Name = textstyle_name;
                                HDD_textstyle.TextSize = Texth;
                                HDD_textstyle.ObliquingAngle = 0;
                                HDD_textstyle.FileName = "arial.ttf";
                                HDD_textstyle.XScale = 1.0;
                                Text_style_table.Add(HDD_textstyle);
                                Trans1.AddNewlyCreatedDBObject(HDD_textstyle, true);
                            }


                            ObjectId Arrowid = ObjectId.Null;
                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("DIMBLK", "_Dot");
                            if (BlockTable1.Has("_Dot") == true)
                            {
                                Arrowid = BlockTable1["_Dot"];
                            }

                            Leader_style_table.UpgradeOpen();


                            Leader_style_table.SetAt(mleaderstyle_name, HDD_mleader);

                            HDD_mleader.ArrowSize = 0;
                            HDD_mleader.BreakSize = Texth;
                            HDD_mleader.DoglegLength = 0;
                            HDD_mleader.LeaderLineColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 256);
                            HDD_mleader.TextColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 256);
                            HDD_mleader.TextHeight = Texth;
                            HDD_mleader.TextStyleId = Text_style_table.ObjectId;
                            HDD_mleader.ArrowSymbolId = Arrowid;
                            HDD_mleader.BlockColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 0);
                            HDD_mleader.BlockRotation = 0;
                            HDD_mleader.BlockScale = new Autodesk.AutoCAD.Geometry.Scale3d(1, 1, 1);
                            HDD_mleader.ContentType = ContentType.MTextContent;
                            HDD_mleader.DrawLeaderOrderType = DrawLeaderOrderType.DrawLeaderHeadFirst;
                            HDD_mleader.DrawMLeaderOrderType = DrawMLeaderOrderType.DrawLeaderFirst;
                            HDD_mleader.EnableBlockRotation = true;
                            HDD_mleader.EnableBlockScale = true;
                            HDD_mleader.EnableDogleg = true;
                            HDD_mleader.EnableFrameText = false;
                            HDD_mleader.EnableLanding = true;
                            HDD_mleader.ExtendLeaderToText = false;
                            HDD_mleader.TextAlignAlwaysLeft = true;
                            HDD_mleader.LandingGap = 0;
                            HDD_mleader.LeaderLineType = LeaderType.StraightLeader;
                            HDD_mleader.LeaderLineWeight = LineWeight.ByBlock;
                            HDD_mleader.MaxLeaderSegmentsPoints = 2;
                            HDD_mleader.Scale = 1;
                            HDD_mleader.TextAlignAlwaysLeft = false;
                            HDD_mleader.TextAlignmentType = TextAlignmentType.LeftAlignment;
                            HDD_mleader.TextAngleType = TextAngleType.HorizontalAngle;
                            Trans1.AddNewlyCreatedDBObject(HDD_mleader, true);


                            string hdd_boreholes = "HDD_boreholes";
                            Functions.Creaza_layer(hdd_boreholes, 1, true);
                            Functions.Creaza_layer("TEXT", 2, true);
                            string no_plot = "NO PLOT";
                            Functions.Creaza_layer(no_plot, 40, false);

                            Autodesk.AutoCAD.EditorInput.PromptEntityResult cl_res;
                            Autodesk.AutoCAD.EditorInput.PromptEntityOptions cl_prompt;
                            cl_prompt = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the centerline:");
                            cl_prompt.SetRejectMessage("\nSelect a polyline!");
                            cl_prompt.AllowNone = true;
                            cl_prompt.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);

                            cl_res = ThisDrawing.Editor.GetEntity(cl_prompt);

                            if (cl_res.Status != PromptStatus.OK)
                            {
                                this.MdiParent.WindowState = FormWindowState.Normal;
                                Editor1.SetImpliedSelection(Empty_array);
                                Editor1.WriteMessage("\nCommand:");
                                set_enable_true();
                                return;
                            }

                            Polyline poly_cl = Trans1.GetObject(cl_res.ObjectId, OpenMode.ForRead) as Polyline;

                            if (poly_cl != null)
                            {
                                double hexag = 1;
                                double graph_vexag = 1;
                                double known_x1 = -123.1234567;
                                double known_y1 = -123.1234567;
                                double known_sta1 = -123.1234567;
                                double known_el1 = -123.1234567;







                                #region 2 STATION AND 2 ELEVATION SELECTION

                                Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_hor1;
                                Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rezh1 = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                                Prompt_rezh1.MessageForAdding = "\nSelect first vertical line (STATION) and the label for it:";
                                Prompt_rezh1.SingleOnly = false;
                                Rezultat_hor1 = ThisDrawing.Editor.GetSelection(Prompt_rezh1);

                                if (Rezultat_hor1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                {
                                    this.MdiParent.WindowState = FormWindowState.Normal;
                                    Editor1.SetImpliedSelection(Empty_array);
                                    set_enable_true();
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    return;
                                }



                                if (Rezultat_hor1.Value.Count != 2)
                                {
                                    MessageBox.Show("the selection has to contain 2 objects. You have selected only" + Rezultat_hor1.Value.Count.ToString() + " object(s).\r\nOperation aborted");
                                    this.MdiParent.WindowState = FormWindowState.Normal;
                                    Editor1.SetImpliedSelection(Empty_array);
                                    set_enable_true();
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    return;
                                }


                                Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_hor2;
                                Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rezh2 = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                                Prompt_rezh2.MessageForAdding = "\nSelect the second vertical line (STATION) and the label for it:";
                                Prompt_rezh2.SingleOnly = false;
                                Rezultat_hor2 = ThisDrawing.Editor.GetSelection(Prompt_rezh2);

                                if (Rezultat_hor2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                {
                                    this.MdiParent.WindowState = FormWindowState.Normal;
                                    Editor1.SetImpliedSelection(Empty_array);
                                    set_enable_true();
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    return;
                                }



                                if (Rezultat_hor2.Value.Count != 2)
                                {
                                    MessageBox.Show("the selection has to contain 2 objects. You have selected only" + Rezultat_hor2.Value.Count.ToString() + " object(s).\r\nOperation aborted");
                                    this.MdiParent.WindowState = FormWindowState.Normal;
                                    Editor1.SetImpliedSelection(Empty_array);
                                    set_enable_true();
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    return;
                                }


                                Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_ver1;
                                Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rezv1 = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                                Prompt_rezv1.MessageForAdding = "\nSelect first horizontal line (ELEVATION) and the label for it:";
                                Prompt_rezv1.SingleOnly = false;
                                Rezultat_ver1 = ThisDrawing.Editor.GetSelection(Prompt_rezv1);

                                if (Rezultat_ver1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                {
                                    this.MdiParent.WindowState = FormWindowState.Normal;
                                    Editor1.SetImpliedSelection(Empty_array);
                                    set_enable_true();
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    return;
                                }

                                if (Rezultat_ver1.Value.Count != 2)
                                {
                                    MessageBox.Show("the selection has to contain 2 objects. You have selected only" + Rezultat_ver1.Value.Count.ToString() + " object(s).\r\nOperation aborted");
                                    this.MdiParent.WindowState = FormWindowState.Normal;
                                    Editor1.SetImpliedSelection(Empty_array);
                                    set_enable_true();
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    return;
                                }


                                Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_ver2;
                                Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rezv2 = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                                Prompt_rezv2.MessageForAdding = "\nSelect second horizontal line (ELEVATION) and the label for it:";
                                Prompt_rezv2.SingleOnly = false;
                                Rezultat_ver2 = ThisDrawing.Editor.GetSelection(Prompt_rezv2);

                                if (Rezultat_ver2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                {
                                    this.MdiParent.WindowState = FormWindowState.Normal;
                                    Editor1.SetImpliedSelection(Empty_array);
                                    set_enable_true();
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    return;
                                }

                                if (Rezultat_ver2.Value.Count != 2)
                                {
                                    MessageBox.Show("the selection has to contain 2 objects. You have selected only" + Rezultat_ver2.Value.Count.ToString() + " object(s).\r\nOperation aborted");
                                    this.MdiParent.WindowState = FormWindowState.Normal;
                                    Editor1.SetImpliedSelection(Empty_array);
                                    set_enable_true();
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    return;
                                }




                                double known_x2 = -123.1234567;
                                double known_y2 = -123.1234567;
                                double known_sta2 = -123.1234567;
                                double known_el2 = -123.1234567;



                                Entity Ent1 = Trans1.GetObject(Rezultat_hor1.Value[0].ObjectId, OpenMode.ForRead) as Entity;
                                Entity Ent2 = Trans1.GetObject(Rezultat_hor1.Value[1].ObjectId, OpenMode.ForRead) as Entity;

                                Entity Ent3 = Trans1.GetObject(Rezultat_ver1.Value[0].ObjectId, OpenMode.ForRead) as Entity;
                                Entity Ent4 = Trans1.GetObject(Rezultat_ver1.Value[1].ObjectId, OpenMode.ForRead) as Entity;

                                Entity Ent11 = Trans1.GetObject(Rezultat_hor2.Value[0].ObjectId, OpenMode.ForRead) as Entity;
                                Entity Ent12 = Trans1.GetObject(Rezultat_hor2.Value[1].ObjectId, OpenMode.ForRead) as Entity;

                                Entity Ent13 = Trans1.GetObject(Rezultat_ver2.Value[0].ObjectId, OpenMode.ForRead) as Entity;
                                Entity Ent14 = Trans1.GetObject(Rezultat_ver2.Value[1].ObjectId, OpenMode.ForRead) as Entity;



                                if (((Ent1 is Polyline || Ent1 is Line) && (Ent2 is MText || Ent2 is DBText)) || ((Ent2 is Polyline || Ent2 is Line) && (Ent1 is MText || Ent1 is DBText)) ||
                                    ((Ent11 is Polyline || Ent11 is Line) && (Ent12 is MText || Ent12 is DBText)) || ((Ent12 is Polyline || Ent12 is Line) && (Ent11 is MText || Ent11 is DBText)))
                                {
                                    #region ent1


                                    if (Ent1 is Polyline)
                                    {
                                        Polyline P1 = Ent1 as Polyline;
                                        if (P1 != null)
                                        {
                                            double x1 = P1.StartPoint.X;
                                            double x2 = P1.EndPoint.X;
                                            if (Math.Round(x1, 2) == Math.Round(x2, 2))
                                            {
                                                known_x1 = x1;

                                            }


                                        }

                                    }

                                    if (Ent1 is Line)
                                    {
                                        Line L1 = Ent1 as Line;
                                        if (L1 != null)
                                        {
                                            double x1 = L1.StartPoint.X;
                                            double x2 = L1.EndPoint.X;
                                            if (Math.Round(x1, 2) == Math.Round(x2, 2))
                                            {
                                                known_x1 = x1;

                                            }


                                        }

                                    }

                                    if (Ent1 is MText)
                                    {
                                        MText M1 = Ent1 as MText;
                                        if (M1 != null)
                                        {
                                            string Continut = M1.Text.Replace("+", "");
                                            if (Functions.IsNumeric(Continut) == true)
                                            {
                                                known_sta1 = Convert.ToDouble(Continut);

                                            }


                                        }

                                    }

                                    if (Ent1 is DBText)
                                    {
                                        DBText T1 = Ent1 as DBText;
                                        if (T1 != null)
                                        {
                                            string Continut = T1.TextString.Replace("+", "");
                                            if (Functions.IsNumeric(Continut) == true)
                                            {
                                                known_sta1 = Convert.ToDouble(Continut);

                                            }


                                        }

                                    }


                                    if (Ent2 is Polyline)
                                    {
                                        Polyline P1 = Ent2 as Polyline;
                                        if (P1 != null)
                                        {
                                            double x1 = P1.StartPoint.X;
                                            double x2 = P1.EndPoint.X;
                                            if (Math.Round(x1, 2) == Math.Round(x2, 2))
                                            {
                                                known_x1 = x1;

                                            }


                                        }

                                    }

                                    if (Ent2 is Line)
                                    {
                                        Line L1 = Ent2 as Line;
                                        if (L1 != null)
                                        {
                                            double x1 = L1.StartPoint.X;
                                            double x2 = L1.EndPoint.X;
                                            if (Math.Round(x1, 2) == Math.Round(x2, 2))
                                            {
                                                known_x1 = x1;

                                            }


                                        }

                                    }

                                    if (Ent2 is MText)
                                    {
                                        MText M1 = Ent2 as MText;
                                        if (M1 != null)
                                        {
                                            string Continut = M1.Text.Replace("+", "");
                                            if (Functions.IsNumeric(Continut) == true)
                                            {
                                                known_sta1 = Convert.ToDouble(Continut);

                                            }


                                        }

                                    }

                                    if (Ent2 is DBText)
                                    {
                                        DBText T1 = Ent2 as DBText;
                                        if (T1 != null)
                                        {
                                            string Continut = T1.TextString.Replace("+", "");
                                            if (Functions.IsNumeric(Continut) == true)
                                            {
                                                known_sta1 = Convert.ToDouble(Continut);

                                            }


                                        }

                                    }
                                    #endregion

                                    #region ent2


                                    if (Ent11 is Polyline)
                                    {
                                        Polyline P11 = Ent11 as Polyline;
                                        if (P11 != null)
                                        {
                                            double x11 = P11.StartPoint.X;
                                            double x12 = P11.EndPoint.X;
                                            if (Math.Round(x11, 2) == Math.Round(x12, 2))
                                            {
                                                known_x2 = x11;
                                            }
                                        }
                                    }

                                    if (Ent11 is Line)
                                    {
                                        Line L11 = Ent11 as Line;
                                        if (L11 != null)
                                        {
                                            double x11 = L11.StartPoint.X;
                                            double x12 = L11.EndPoint.X;
                                            if (Math.Round(x11, 2) == Math.Round(x12, 2))
                                            {
                                                known_x2 = x11;

                                            }


                                        }

                                    }

                                    if (Ent11 is MText)
                                    {
                                        MText M11 = Ent11 as MText;
                                        if (M11 != null)
                                        {
                                            string Continut = M11.Text.Replace("+", "");
                                            if (Functions.IsNumeric(Continut) == true)
                                            {
                                                known_sta2 = Convert.ToDouble(Continut);

                                            }


                                        }

                                    }

                                    if (Ent11 is DBText)
                                    {
                                        DBText T11 = Ent11 as DBText;
                                        if (T11 != null)
                                        {
                                            string Continut = T11.TextString.Replace("+", "");
                                            if (Functions.IsNumeric(Continut) == true)
                                            {
                                                known_sta2 = Convert.ToDouble(Continut);

                                            }


                                        }

                                    }


                                    if (Ent12 is Polyline)
                                    {
                                        Polyline P12 = Ent12 as Polyline;
                                        if (P12 != null)
                                        {
                                            double x12 = P12.StartPoint.X;
                                            double x22 = P12.EndPoint.X;
                                            if (Math.Round(x12, 2) == Math.Round(x22, 2))
                                            {
                                                known_x2 = x12;

                                            }


                                        }

                                    }

                                    if (Ent12 is Line)
                                    {
                                        Line L12 = Ent12 as Line;
                                        if (L12 != null)
                                        {
                                            double x12 = L12.StartPoint.X;
                                            double x22 = L12.EndPoint.X;
                                            if (Math.Round(x12, 2) == Math.Round(x22, 2))
                                            {
                                                known_x2 = x12;

                                            }


                                        }

                                    }

                                    if (Ent12 is MText)
                                    {
                                        MText M12 = Ent12 as MText;
                                        if (M12 != null)
                                        {
                                            string Continut = M12.Text.Replace("+", "");
                                            if (Functions.IsNumeric(Continut) == true)
                                            {
                                                known_sta2 = Convert.ToDouble(Continut);

                                            }


                                        }

                                    }

                                    if (Ent12 is DBText)
                                    {
                                        DBText T12 = Ent12 as DBText;
                                        if (T12 != null)
                                        {
                                            string Continut = T12.TextString.Replace("+", "");
                                            if (Functions.IsNumeric(Continut) == true)
                                            {
                                                known_sta2 = Convert.ToDouble(Continut);

                                            }


                                        }

                                    }
                                    #endregion

                                }

                                if (((Ent3 is Polyline || Ent3 is Line) & (Ent4 is MText || Ent4 is DBText)) || ((Ent4 is Polyline || Ent4 is Line) & (Ent3 is MText || Ent3 is DBText)) ||
                                    ((Ent13 is Polyline || Ent13 is Line) && (Ent14 is MText || Ent14 is DBText)) || ((Ent14 is Polyline || Ent14 is Line) && (Ent13 is MText || Ent13 is DBText)))
                                {
                                    #region ent3

                                    if (Ent3 is Polyline)
                                    {
                                        Polyline P1 = Ent3 as Polyline;
                                        if (P1 != null)
                                        {
                                            double y1 = P1.StartPoint.Y;
                                            double y2 = P1.EndPoint.Y;
                                            if (Math.Round(y1, 2) == Math.Round(y2, 2))
                                            {
                                                known_y1 = y1;

                                            }


                                        }

                                    }

                                    if (Ent3 is Line)
                                    {
                                        Line L1 = Ent3 as Line;
                                        if (L1 != null)
                                        {
                                            double y1 = L1.StartPoint.Y;
                                            double y2 = L1.EndPoint.Y;
                                            if (Math.Round(y1, 2) == Math.Round(y2, 2))
                                            {
                                                known_y1 = y1;

                                            }


                                        }

                                    }

                                    if (Ent3 is MText)
                                    {
                                        MText M1 = Ent3 as MText;
                                        if (M1 != null)
                                        {
                                            string Continut = M1.Text.Replace("'", "");
                                            if (Functions.IsNumeric(Continut) == true)
                                            {
                                                known_el1 = Convert.ToDouble(Continut);

                                            }


                                        }

                                    }

                                    if (Ent3 is DBText)
                                    {
                                        DBText T1 = Ent3 as DBText;
                                        if (T1 != null)
                                        {
                                            string Continut = T1.TextString.Replace("'", "");
                                            if (Functions.IsNumeric(Continut) == true)
                                            {
                                                known_el1 = Convert.ToDouble(Continut);

                                            }


                                        }

                                    }


                                    if (Ent4 is Polyline)
                                    {
                                        Polyline P1 = Ent4 as Polyline;
                                        if (P1 != null)
                                        {
                                            double y1 = P1.StartPoint.Y;
                                            double y2 = P1.EndPoint.Y;
                                            if (Math.Round(y1, 2) == Math.Round(y2, 2))
                                            {
                                                known_y1 = y1;

                                            }


                                        }

                                    }

                                    if (Ent4 is Line)
                                    {
                                        Line L1 = Ent4 as Line;
                                        if (L1 != null)
                                        {
                                            double y1 = L1.StartPoint.Y;
                                            double y2 = L1.EndPoint.Y;
                                            if (Math.Round(y1, 2) == Math.Round(y2, 2))
                                            {
                                                known_y1 = y1;

                                            }


                                        }

                                    }

                                    if (Ent4 is MText)
                                    {
                                        MText M1 = Ent4 as MText;
                                        if (M1 != null)
                                        {
                                            string Continut = M1.Text.Replace("'", "");
                                            if (Functions.IsNumeric(Continut) == true)
                                            {
                                                known_el1 = Convert.ToDouble(Continut);

                                            }


                                        }

                                    }

                                    if (Ent4 is DBText)
                                    {
                                        DBText T1 = Ent4 as DBText;
                                        if (T1 != null)
                                        {
                                            string Continut = T1.TextString.Replace("'", "");
                                            if (Functions.IsNumeric(Continut) == true)
                                            {
                                                known_el1 = Convert.ToDouble(Continut);

                                            }


                                        }

                                    }
                                    #endregion

                                    #region ent4

                                    if (Ent13 is Polyline)
                                    {
                                        Polyline P13 = Ent13 as Polyline;
                                        if (P13 != null)
                                        {
                                            double y13 = P13.StartPoint.Y;
                                            double y23 = P13.EndPoint.Y;
                                            if (Math.Round(y13, 2) == Math.Round(y23, 2))
                                            {
                                                known_y2 = y13;

                                            }


                                        }

                                    }

                                    if (Ent13 is Line)
                                    {
                                        Line L13 = Ent13 as Line;
                                        if (L13 != null)
                                        {
                                            double y13 = L13.StartPoint.Y;
                                            double y23 = L13.EndPoint.Y;
                                            if (Math.Round(y13, 2) == Math.Round(y23, 2))
                                            {
                                                known_y2 = y13;

                                            }


                                        }

                                    }

                                    if (Ent13 is MText)
                                    {
                                        MText M13 = Ent13 as MText;
                                        if (M13 != null)
                                        {
                                            string Continut = M13.Text.Replace("'", "");
                                            if (Functions.IsNumeric(Continut) == true)
                                            {
                                                known_el2 = Convert.ToDouble(Continut);

                                            }


                                        }

                                    }

                                    if (Ent13 is DBText)
                                    {
                                        DBText T1 = Ent13 as DBText;
                                        if (T1 != null)
                                        {
                                            string Continut = T1.TextString.Replace("'", "");
                                            if (Functions.IsNumeric(Continut) == true)
                                            {
                                                known_el2 = Convert.ToDouble(Continut);

                                            }


                                        }

                                    }


                                    if (Ent14 is Polyline)
                                    {
                                        Polyline P1 = Ent14 as Polyline;
                                        if (P1 != null)
                                        {
                                            double y1 = P1.StartPoint.Y;
                                            double y2 = P1.EndPoint.Y;
                                            if (Math.Round(y1, 2) == Math.Round(y2, 2))
                                            {
                                                known_y2 = y1;

                                            }


                                        }

                                    }

                                    if (Ent14 is Line)
                                    {
                                        Line L1 = Ent14 as Line;
                                        if (L1 != null)
                                        {
                                            double y1 = L1.StartPoint.Y;
                                            double y2 = L1.EndPoint.Y;
                                            if (Math.Round(y1, 2) == Math.Round(y2, 2))
                                            {
                                                known_y2 = y1;

                                            }


                                        }

                                    }

                                    if (Ent14 is MText)
                                    {
                                        MText M1 = Ent14 as MText;
                                        if (M1 != null)
                                        {
                                            string Continut = M1.Text.Replace("'", "");
                                            if (Functions.IsNumeric(Continut) == true)
                                            {
                                                known_el2 = Convert.ToDouble(Continut);

                                            }


                                        }

                                    }

                                    if (Ent14 is DBText)
                                    {
                                        DBText T1 = Ent14 as DBText;
                                        if (T1 != null)
                                        {
                                            string Continut = T1.TextString.Replace("'", "");
                                            if (Functions.IsNumeric(Continut) == true)
                                            {
                                                known_el2 = Convert.ToDouble(Continut);

                                            }


                                        }

                                    }
                                    #endregion

                                }

                                if (known_x1 != -123.1234567 && known_y1 != -123.1234567 && known_sta1 != -123.1234567 && known_el1 != -123.1234567 && known_x2 != -123.1234567 && known_y2 != -123.1234567 && known_sta2 != -123.1234567 && known_el2 != -123.1234567)
                                {
                                    hexag = Math.Abs(known_x1 - known_x2) / Math.Abs(known_sta1 - known_sta2);
                                    graph_vexag = Math.Abs(known_y1 - known_y2) / Math.Abs(known_el1 - known_el2);
                                }

                                #endregion

                                string col_pt = "PointID";

                                string col_elev = "Elevation";
                                string col_north = "North";
                                string col_east = "East";
                                string col_lat = "Lat";
                                string col_long = "Long";

                                string col_depth = "Depth";
                                string col_bottom = "Bottom";
                                string col_graphic = "Graphic";
                                string col_len = "Length";
                                string col_type = "Type";
                                string col_number = "Number";

                                string col_blow1 = "Blows 1st 6in";
                                string col_blow2 = "Blows 2nd 6in";
                                string col_blow3 = "Blows 3rd 6in";
                                string col_blow4 = "Blows 4th 6in";

                                string col_recovery = "Recovery";
                                string col_rqd = "RQD";


                                double wdth = Texth * 3;

                                string cs1 = comboBox1.Text;
                                string cs2 = comboBox2.Text;

                                BlockTable1.UpgradeOpen();

                                Autodesk.AutoCAD.Colors.Color color_rect = Autodesk.AutoCAD.Colors.Color.FromRgb(0, 0, 127);
                                Autodesk.AutoCAD.Colors.Color color_sym = Autodesk.AutoCAD.Colors.Color.FromRgb(0, 0, 255);
                                Autodesk.AutoCAD.Colors.Color color_labels = Autodesk.AutoCAD.Colors.Color.FromRgb(51, 51, 51);
                                Autodesk.AutoCAD.Colors.Color color_white = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 255, 255);
                                Autodesk.AutoCAD.Colors.Color color_explanations = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 0, 255);
                                Point3d pt_legend = new Point3d();


                                for (int i = 0; i < dt_point.Rows.Count; ++i)
                                {
                                    string borehole1 = Convert.ToString(dt_point.Rows[i][col_pt]).Replace(" ", "");
                                    double north1 = -1.234;
                                    double east1 = -1.234;
                                    double lat1 = -1.234;
                                    double long1 = -1.234;

                                    if (dt_point.Rows[i][col_long] != DBNull.Value)
                                    {
                                        long1 = Convert.ToDouble(dt_point.Rows[i][col_long]);
                                    }

                                    if (dt_point.Rows[i][col_lat] != DBNull.Value)
                                    {
                                        lat1 = Convert.ToDouble(dt_point.Rows[i][col_lat]);
                                    }

                                    if (dt_point.Rows[i][col_north] != DBNull.Value)
                                    {
                                        north1 = Convert.ToDouble(dt_point.Rows[i][col_north]);
                                    }

                                    if (dt_point.Rows[i][col_east] != DBNull.Value)
                                    {
                                        east1 = Convert.ToDouble(dt_point.Rows[i][col_east]);
                                    }

                                    Point3d point_borhole_ll = new Point3d(long1, lat1, poly_cl.Elevation);
                                    Point3d point_borhole = new Point3d(east1, north1, poly_cl.Elevation);

                                    if (cs1 != "" && cs2 != "" && cs1 != cs2 && north1 == -1.234)
                                    {
                                        point_borhole = Functions.Convert_coordinate_from_CS_to_new_CS(point_borhole_ll, cs1, cs2);
                                    }

                                    Point3d point_on_poly = poly_cl.GetClosestPointTo(point_borhole, Vector3d.ZAxis, false);
                                    double sta = poly_cl.GetDistAtPoint(point_on_poly);
                                    double x_ins = known_x1 - (hexag * (known_sta1 - sta));

                                    double elev = Convert.ToDouble(dt_point.Rows[i][col_elev]);
                                    double y_ins = known_y1 - (graph_vexag * (known_el1 - elev));

                                    #region Plan view

                                    int idx2 = 1;
                                    bool exista2 = true;

                                    do
                                    {
                                        if (BlockTable1.Has("Symbol_" + borehole1 + idx2.ToString()) == false)
                                        {
                                            using (BlockTableRecord bltrec1 = new BlockTableRecord())
                                            {
                                                bltrec1.Name = "Symbol_" + borehole1 + idx2.ToString();

                                                Circle cerc_bh = new Circle(new Point3d(0, 0, 0), Vector3d.ZAxis, Texth_plan_view);

                                                cerc_bh.Layer = "0";
                                                cerc_bh.Color = color_white;
                                                cerc_bh.LineWeight = LineWeight.LineWeight100;
                                                bltrec1.AppendEntity(cerc_bh);

                                                cerc_bh = new Circle(new Point3d(0, 0, 0), Vector3d.ZAxis, Texth_plan_view);
                                                cerc_bh.Layer = "0";
                                                cerc_bh.Color = color_sym;
                                                bltrec1.AppendEntity(cerc_bh);


                                                Polyline poly_sym_background = get_poly_sym(Texth_plan_view);
                                                poly_sym_background.Layer = "0";
                                                poly_sym_background.Color = color_white;
                                                poly_sym_background.ConstantWidth = scale2 / 50;
                                                bltrec1.AppendEntity(poly_sym_background);



                                                Polyline poly_sym = get_poly_sym(Texth_plan_view);
                                                poly_sym.Layer = "0";
                                                poly_sym.Color = color_sym;
                                                bltrec1.AppendEntity(poly_sym);

                                                MText mtext_label = new MText();
                                                mtext_label.Contents = borehole1;
                                                mtext_label.TextHeight = Texth_plan_view;
                                                mtext_label.Attachment = AttachmentPoint.MiddleCenter;
                                                mtext_label.Location = new Point3d(0, 3 * Texth_plan_view, 0);
                                                mtext_label.Layer = "0";
                                                mtext_label.Color = color_white;
                                                mtext_label.LineWeight = LineWeight.LineWeight100;
                                                bltrec1.AppendEntity(mtext_label);

                                                mtext_label = new MText();
                                                mtext_label.Contents = borehole1;
                                                mtext_label.TextHeight = Texth_plan_view;
                                                mtext_label.Attachment = AttachmentPoint.MiddleCenter;
                                                mtext_label.Location = new Point3d(0, 3 * Texth_plan_view, 0);
                                                mtext_label.Layer = "0";
                                                mtext_label.Color = color_sym;
                                                bltrec1.AppendEntity(mtext_label);

                                                BlockTable1.Add(bltrec1);
                                                Trans1.AddNewlyCreatedDBObject(bltrec1, true);
                                                BlockReference b1 = Functions.InsertBlock_with_2scales(ThisDrawing.Database, BTrecord, bltrec1.Name, point_borhole, 1, 1, 0, hdd_boreholes);
                                                b1.ColorIndex = 256;




                                            }
                                            exista2 = false;
                                        }
                                        else
                                        {
                                            ++idx2;
                                        }
                                    } while (exista2 == true);





                                    double offset1 = Math.Pow(Math.Pow((point_borhole.X - point_on_poly.X), 2) + Math.Pow((point_borhole.Y - point_on_poly.Y), 2), 0.5);
                                    string left_right = Functions.Angle_left_right(poly_cl, point_borhole);
                                    #endregion


                                    #region mleader top
                                    string content1 = "EL. " + Functions.Get_String_Rounded(elev, 0) + f + "\r\n" +
                                                                Functions.Get_String_Rounded(offset1, 0) + f + " " + left_right;

                                    MLeader mleader_top = Functions.creaza_mleader_with_style(new Point3d(x_ins, y_ins, 0), content1,
                                                                                                Texth, mleaderstyle_name, textstyle_name,
                                                                                                        0.3 * scale1, 0.25 * scale1, hdd_boreholes);

                                    mleader_top.SetTextAttachmentType(TextAttachmentType.AttachmentBottomOfTopLine, LeaderDirectionType.RightLeader);
                                    mleader_top.SetTextAttachmentType(TextAttachmentType.AttachmentBottomOfTopLine, LeaderDirectionType.LeftLeader);
                                    mleader_top.ArrowSymbolId = ObjectId.Null;

                                    mleader_top.LineWeight = LineWeight.LineWeight000;
                                    mleader_top.LeaderLineWeight = LineWeight.LineWeight000;
                                    mleader_top.Color = color_labels;
                                    #endregion

                                    Point3d pt_ins = new Point3d(x_ins, y_ins, 0);

                                    #region header top
                                    int idx = 1;
                                    bool exista = true;
                                    do
                                    {
                                        if (BlockTable1.Has(borehole1 + idx.ToString()) == false)
                                        {
                                            using (BlockTableRecord bltrec1 = new BlockTableRecord())
                                            {
                                                bltrec1.Name = borehole1 + idx.ToString();
                                                MText mtext_borehole = new MText();
                                                mtext_borehole.Contents = borehole1;
                                                mtext_borehole.TextHeight = 2 * Texth;
                                                mtext_borehole.TextStyleId = HDD_textstyle.ObjectId;
                                                mtext_borehole.Attachment = AttachmentPoint.MiddleCenter;
                                                mtext_borehole.Location = new Point3d(pt_ins.X, pt_ins.Y + 2 * wdth + 0.5 * 3.25 * Texth, pt_ins.Z);
                                                mtext_borehole.Layer = "0";
                                                mtext_borehole.Color = color_labels;
                                                mtext_borehole.BackgroundFill = true;
                                                mtext_borehole.UseBackgroundColor = true;
                                                mtext_borehole.BackgroundScaleFactor = 1.2;
                                                mtext_borehole.TransformBy(Matrix3d.Displacement(pt_ins.GetVectorTo(new Point3d(0, 0, 0))));
                                                bltrec1.AppendEntity(mtext_borehole);

                                                Extents3d ext1 = mtext_borehole.GeometricExtents;
                                                double width_mtext = Math.Abs(ext1.MaxPoint.X - ext1.MinPoint.X) + wdth / 1.5;

                                                Polyline poly_top = new Polyline();
                                                poly_top.AddVertexAt(0, new Point2d(pt_ins.X, pt_ins.Y), 0, Texth / 20, Texth / 20);
                                                poly_top.AddVertexAt(1, new Point2d(pt_ins.X, pt_ins.Y + 2 * wdth), 0, Texth / 20, Texth / 20);
                                                poly_top.AddVertexAt(2, new Point2d(pt_ins.X + width_mtext / 2, pt_ins.Y + 2 * wdth), 0, Texth / 20, Texth / 20);
                                                poly_top.AddVertexAt(3, new Point2d(pt_ins.X + width_mtext / 2, pt_ins.Y + 2 * wdth + 3.25 * Texth), 0, Texth / 20, Texth / 20);
                                                poly_top.AddVertexAt(4, new Point2d(pt_ins.X - width_mtext / 2, pt_ins.Y + 2 * wdth + 3.25 * Texth), 0, Texth / 20, Texth / 20);
                                                poly_top.AddVertexAt(5, new Point2d(pt_ins.X - width_mtext / 2, pt_ins.Y + 2 * wdth), 0, Texth / 20, Texth / 20);
                                                poly_top.AddVertexAt(6, new Point2d(pt_ins.X, pt_ins.Y + 2 * wdth), 0, Texth / 20, Texth / 20);
                                                poly_top.Layer = "0";
                                                poly_top.ColorIndex = 7;
                                                //poly_top.LineWeight = LineWeight.LineWeight000;
                                                poly_top.TransformBy(Matrix3d.Displacement(pt_ins.GetVectorTo(new Point3d(0, 0, 0))));
                                                bltrec1.AppendEntity(poly_top);

                                                BlockTable1.Add(bltrec1);
                                                Trans1.AddNewlyCreatedDBObject(bltrec1, true);
                                                BlockReference b1 = Functions.InsertBlock_with_2scales(ThisDrawing.Database, BTrecord, borehole1 + idx.ToString(), pt_ins, 1, 1, 0, hdd_boreholes);
                                                b1.ColorIndex = 256;
                                            }
                                            exista = false;
                                        }
                                        else
                                        {
                                            ++idx;
                                        }
                                    } while (exista == true);
                                    #endregion

                                    System.Data.DataTable dtc = new System.Data.DataTable();
                                    dtc.Columns.Add("type", typeof(string));
                                    dtc.Columns.Add("top", typeof(double));
                                    dtc.Columns.Add("bottom", typeof(double));
                                    dtc.Columns.Add("desc", typeof(string));

                                    #region dtc
                                    if (dt_rock != null && dt_rock.Rows.Count > 0)
                                    {
                                        for (int j = 0; j < dt_rock.Rows.Count; ++j)
                                        {
                                            string borehole2 = Convert.ToString(dt_rock.Rows[j][col_pt]).Replace(" ", "");
                                            if (borehole1.ToUpper() == borehole2.ToUpper())
                                            {
                                                double top1 = Convert.ToDouble(dt_rock.Rows[j][col_depth]);
                                                double bottom1 = Convert.ToDouble(dt_rock.Rows[j][col_bottom]);
                                                string desc = Convert.ToString(dt_rock.Rows[j][col_graphic]);
                                                dtc.Rows.Add();
                                                dtc.Rows[dtc.Rows.Count - 1]["type"] = "rock";
                                                dtc.Rows[dtc.Rows.Count - 1]["top"] = top1;
                                                dtc.Rows[dtc.Rows.Count - 1]["bottom"] = bottom1;
                                                dtc.Rows[dtc.Rows.Count - 1]["desc"] = desc;

                                            }
                                        }
                                    }


                                    if (dt_soil != null && dt_soil.Rows.Count > 0)
                                    {

                                        for (int j = 0; j < dt_soil.Rows.Count; ++j)
                                        {
                                            string borehole2 = Convert.ToString(dt_soil.Rows[j][col_pt]).Replace(" ", "");
                                            if (borehole1.ToUpper() == borehole2.ToUpper())
                                            {
                                                double top1 = Convert.ToDouble(dt_soil.Rows[j][col_depth]);
                                                double bottom1 = Convert.ToDouble(dt_soil.Rows[j][col_bottom]);
                                                string desc = Convert.ToString(dt_soil.Rows[j][col_graphic]);
                                                dtc.Rows.Add();
                                                dtc.Rows[dtc.Rows.Count - 1]["type"] = "soil";
                                                dtc.Rows[dtc.Rows.Count - 1]["top"] = top1;
                                                dtc.Rows[dtc.Rows.Count - 1]["bottom"] = bottom1;
                                                dtc.Rows[dtc.Rows.Count - 1]["desc"] = desc;
                                            }
                                        }
                                    }

                                    dtc = Functions.Sort_data_table(dtc, "top");
                                    #endregion

                                    if (dtc.Rows.Count > 0)
                                    {
                                        #region polylines and hatches
                                        for (int j = 0; j < dtc.Rows.Count; ++j)
                                        {
                                            double top1 = Convert.ToDouble(dtc.Rows[j]["top"]);
                                            double bottom1 = Convert.ToDouble(dtc.Rows[j]["bottom"]);
                                            string type1 = Convert.ToString(dtc.Rows[j]["type"]);

                                            if (j < dtc.Rows.Count - 1)
                                            {
                                                for (int k = j + 1; k < dtc.Rows.Count; ++k)
                                                {
                                                    double top2 = Convert.ToDouble(dtc.Rows[k]["top"]);
                                                    double bottom2 = Convert.ToDouble(dtc.Rows[k]["bottom"]);
                                                    string type2 = Convert.ToString(dtc.Rows[k]["type"]);
                                                    if (top2 < bottom1 && type1 != type2)
                                                    {
                                                        if (type1 == "soil")
                                                        {
                                                            dtc.Rows[j]["bottom"] = top2;
                                                        }
                                                        else
                                                        {
                                                            dtc.Rows[k]["top"] = bottom1;
                                                        }
                                                    }

                                                    if (top2 < bottom1 && type1 == type2)
                                                    {
                                                        MessageBox.Show(borehole1 + " overlap on " + type1 + "\r\ntop = " + top2.ToString() + "\r\nbottom = " + bottom1.ToString());
                                                    }
                                                }
                                            }
                                        }



                                        idx = 2;
                                        exista = true;
                                        do
                                        {
                                            if (BlockTable1.Has(borehole1 + idx.ToString()) == false)
                                            {
                                                using (BlockTableRecord bltrec1 = new BlockTableRecord())
                                                {
                                                    bltrec1.Name = borehole1 + idx.ToString();

                                                    for (int j = 0; j < dtc.Rows.Count; ++j)
                                                    {
                                                        double top1 = Convert.ToDouble(dtc.Rows[j]["top"]);
                                                        double bottom1 = Convert.ToDouble(dtc.Rows[j]["bottom"]);
                                                        string desc = Convert.ToString(dtc.Rows[j]["desc"]);
                                                        if (lista_legend.Contains(desc.ToUpper()) == false) lista_legend.Add(desc.ToUpper());

                                                        Polyline poly1 = new Polyline();
                                                        poly1.AddVertexAt(0, new Point2d(pt_ins.X - wdth / 2, pt_ins.Y - top1 * graph_vexag * stick_vexag), 0, 0, 0);
                                                        poly1.AddVertexAt(1, new Point2d(pt_ins.X + wdth / 2, pt_ins.Y - top1 * graph_vexag * stick_vexag), 0, 0, 0);
                                                        poly1.AddVertexAt(2, new Point2d(pt_ins.X + wdth / 2, pt_ins.Y - bottom1 * graph_vexag * stick_vexag), 0, 0, 0);
                                                        poly1.AddVertexAt(3, new Point2d(pt_ins.X - wdth / 2, pt_ins.Y - bottom1 * graph_vexag * stick_vexag), 0, 0, 0);
                                                        poly1.Closed = true;
                                                        poly1.Layer = "0";
                                                        poly1.Color = color_rect;


                                                        if (i == dt_point.Rows.Count - 1)
                                                        {
                                                            if (j == dtc.Rows.Count - 1)
                                                            {
                                                                pt_legend = new Point3d(pt_ins.X - 15 - wdth / 2, pt_ins.Y - (bottom1 - 10) * graph_vexag * stick_vexag, 0);
                                                            }
                                                        }

                                                        poly1.LineWeight = LineWeight.LineWeight000;
                                                        poly1.TransformBy(Matrix3d.Displacement(pt_ins.GetVectorTo(new Point3d(0, 0, 0))));
                                                        bltrec1.AppendEntity(poly1);


                                                        #region hatches

                                                        if (desc.ToUpper() == "ML")
                                                        {
                                                            add_pattern_ML(bltrec1, scale1, poly1, BTrecord, Trans1);

                                                        }
                                                        if (desc.ToUpper() == "CL")
                                                        {
                                                            add_pattern_CL(bltrec1, scale1, poly1, BTrecord, Trans1);

                                                        }

                                                        if (desc.ToUpper() == "SANDSTONE")
                                                        {
                                                            add_pattern_sandstone(bltrec1, scale1, graph_vexag, stick_vexag, poly1, BTrecord, Trans1);

                                                        }

                                                        if (desc.ToUpper() == "TOPSOIL")
                                                        {
                                                            add_pattern_topsoil(bltrec1, scale1, graph_vexag, stick_vexag, poly1, BTrecord, Trans1);
                                                        }

                                                        if (desc.ToUpper() == "SM")
                                                        {
                                                            add_pattern_SM(bltrec1, scale1, graph_vexag, stick_vexag, poly1, BTrecord, Trans1);
                                                        }

                                                        if (desc.ToUpper() == "GM")
                                                        {
                                                            add_pattern_GM(bltrec1, scale1, graph_vexag, stick_vexag, poly1, BTrecord, Trans1);
                                                        }

                                                        if (desc.ToUpper() == "SHALE")
                                                        {
                                                            add_pattern_Shale(bltrec1, scale1, graph_vexag, stick_vexag, poly1, BTrecord, Trans1);
                                                        }
                                                        if (desc.ToUpper() == "SP")
                                                        {
                                                            add_pattern_SP(bltrec1, scale1, graph_vexag, stick_vexag, poly1, BTrecord, Trans1);
                                                        }
                                                        if (desc.ToUpper() == "GP")
                                                        {
                                                            add_pattern_GP(bltrec1, scale1, graph_vexag, stick_vexag, poly1, BTrecord, Trans1);
                                                        }
                                                        if (desc.ToUpper() == "GC")
                                                        {
                                                            add_pattern_GC(bltrec1, scale1, graph_vexag, stick_vexag, poly1, BTrecord, Trans1);
                                                        }
                                                        if (desc.ToUpper() == "MUDSTONE")
                                                        {
                                                            add_pattern_MUDSTONE(bltrec1, scale1, graph_vexag, stick_vexag, poly1, BTrecord, Trans1);
                                                        }
                                                        if (desc.ToUpper() == "CL-ML")
                                                        {
                                                            add_pattern_CLML(bltrec1, scale1, poly1, BTrecord, Trans1);

                                                        }
                                                        if (desc.ToUpper() == "GP-GC")
                                                        {
                                                            add_pattern_GPGC(bltrec1, scale1, graph_vexag, stick_vexag, poly1, BTrecord, Trans1);

                                                        }
                                                        if (desc.ToUpper() == "SC")
                                                        {
                                                            add_pattern_SC(bltrec1, scale1, graph_vexag, stick_vexag, poly1, BTrecord, Trans1);

                                                        }
                                                        #endregion

                                                        MText string_desc = new MText();
                                                        string_desc.Contents = desc;
                                                        string_desc.TextHeight = graph_vexag / 2;
                                                        string_desc.Attachment = AttachmentPoint.MiddleCenter;
                                                        string_desc.Location = new Point3d(pt_ins.X, pt_ins.Y - top1 * graph_vexag * stick_vexag - 0.5 * graph_vexag * stick_vexag * Math.Abs(top1 - bottom1), 0);
                                                        string_desc.Layer = no_plot;
                                                        string_desc.ColorIndex = 256;
                                                        string_desc.TransformBy(Matrix3d.Displacement(pt_ins.GetVectorTo(new Point3d(0, 0, 0))));
                                                        bltrec1.AppendEntity(string_desc);
                                                    }
                                                    BlockTable1.Add(bltrec1);
                                                    Trans1.AddNewlyCreatedDBObject(bltrec1, true);
                                                    BlockReference b1 = Functions.InsertBlock_with_2scales(ThisDrawing.Database, BTrecord, borehole1 + idx.ToString(), pt_ins, 1, 1, 0, hdd_boreholes);
                                                    b1.ColorIndex = 256;
                                                }
                                                exista = false;
                                            }
                                            else
                                            {
                                                ++idx;
                                            }
                                        } while (exista == true);
                                        #endregion

                                        double depth1 = Convert.ToDouble(dtc.Rows[dtc.Rows.Count - 1]["bottom"]);
                                        double el_bottom = elev - depth1;

                                        #region MTEXT bottom
                                        string content2 = "EL. " + Functions.Get_String_Rounded(el_bottom, 0) + f;

                                        //MLeader mleader_bottom = Functions.creaza_mleader_with_style(new Point3d(x_ins, pt_ins.Y - depth1 * vexag, 0), content2, Texth, mleaderstyle_name, textstyle_name, -100, -40, hdd_boreholes);
                                        //mleader_bottom.SetTextAttachmentType(TextAttachmentType.AttachmentBottomOfTopLine, LeaderDirectionType.RightLeader);
                                        //mleader_bottom.SetTextAttachmentType(TextAttachmentType.AttachmentBottomOfTopLine, LeaderDirectionType.LeftLeader);
                                        //mleader_bottom.ArrowSymbolId = ObjectId.Null;

                                        MText boe = new MText();
                                        boe.Contents = "B.O.E. " + Functions.Get_String_Rounded(depth1, 0) + " " + f1 + "\r\n(EL. " + Functions.Get_String_Rounded(el_bottom, 0) + f + ")";
                                        boe.TextHeight = Texth;
                                        boe.Attachment = AttachmentPoint.TopCenter;
                                        boe.TextStyleId = HDD_textstyle.ObjectId;
                                        boe.Location = new Point3d(pt_ins.X, pt_ins.Y - depth1 * graph_vexag * stick_vexag - Texth / 2, 0);
                                        boe.Layer = hdd_boreholes;
                                        boe.ColorIndex = 7;
                                        boe.BackgroundFill = true;
                                        boe.UseBackgroundColor = true;
                                        boe.BackgroundScaleFactor = 1.2;
                                        BTrecord.AppendEntity(boe);
                                        Trans1.AddNewlyCreatedDBObject(boe, true);
                                        #endregion
                                    }

                                    #region block_labels
                                    idx = 3;
                                    string suff = "A";
                                    exista = true;
                                    do
                                    {
                                        if (BlockTable1.Has(borehole1 + idx.ToString()) == false && BlockTable1.Has(borehole1 + idx.ToString() + suff) == false)
                                        {
                                            using (BlockTableRecord bltrec1 = new BlockTableRecord())
                                            {
                                                bltrec1.Name = borehole1 + idx.ToString();
                                                using (BlockTableRecord bltrec2 = new BlockTableRecord())
                                                {
                                                    bltrec2.Name = borehole1 + idx.ToString() + suff;
                                                    #region label soils
                                                    if (dt_core_s != null && dt_core_s.Rows.Count > 0)
                                                    {
                                                        for (int j = 0; j < dt_core_s.Rows.Count; ++j)
                                                        {
                                                            string borehole2 = Convert.ToString(dt_core_s.Rows[j][col_pt]).Replace(" ", "");
                                                            if (borehole1.ToUpper() == borehole2.ToUpper())
                                                            {
                                                                string type1 = Convert.ToString(dt_core_s.Rows[j][col_type]);
                                                                string number1 = Convert.ToString(dt_core_s.Rows[j][col_number]);
                                                                double depth1 = Convert.ToDouble(dt_core_s.Rows[j][col_depth]);
                                                                double deptht = Convert.ToDouble(dt_core_s.Rows[j][col_len]);
                                                                string blow1 = Convert.ToString(dt_core_s.Rows[j][col_blow1]);
                                                                string blow2 = "";
                                                                if (dt_core_s.Rows[j][col_blow2] != DBNull.Value)
                                                                {
                                                                    blow2 = Convert.ToString(dt_core_s.Rows[j][col_blow2]);
                                                                }
                                                                string blow3 = "";

                                                                if (dt_core_s.Rows[j][col_blow3] != DBNull.Value)
                                                                {
                                                                    blow3 = Convert.ToString(dt_core_s.Rows[j][col_blow3]);
                                                                }
                                                                string blow4 = "";
                                                                if (dt_core_s.Rows[j][col_blow4] != DBNull.Value)
                                                                {
                                                                    blow4 = Convert.ToString(dt_core_s.Rows[j][col_blow4]);
                                                                }
                                                                string desc = "";
                                                                if (blow4.Contains("/") == true)
                                                                {
                                                                    desc = blow4;
                                                                }
                                                                else if (blow3.Contains("/") == true)
                                                                {
                                                                    desc = blow3;
                                                                }
                                                                else if (blow2.Contains("/") == true)
                                                                {
                                                                    desc = blow2;
                                                                }
                                                                else if (blow1.Contains("/") == true)
                                                                {
                                                                    desc = blow1;
                                                                }
                                                                else
                                                                {
                                                                    int nr = 0;
                                                                    int count = 0;
                                                                    if (blow4.Contains("WOH") == true)
                                                                    {
                                                                        count = 1;
                                                                    }
                                                                    else if (Functions.IsNumeric(blow4) == true)
                                                                    {
                                                                        nr = Convert.ToInt32(blow4);
                                                                        count = 1;
                                                                    }

                                                                    if (blow3.Contains("WOH") == true)
                                                                    {
                                                                        count = ++count;
                                                                    }
                                                                    else if (Functions.IsNumeric(blow3) == true)
                                                                    {
                                                                        nr = nr + Convert.ToInt32(blow3);
                                                                        count = ++count;
                                                                    }

                                                                    if (count < 2)
                                                                    {
                                                                        if (blow2.Contains("WOH") == true)
                                                                        {
                                                                            count = ++count;
                                                                        }
                                                                        else if (Functions.IsNumeric(blow2) == true)
                                                                        {
                                                                            nr = nr + Convert.ToInt32(blow2);
                                                                            count = ++count;
                                                                        }
                                                                    }
                                                                    if (count < 2)
                                                                    {
                                                                        if (Functions.IsNumeric(blow1) == true)
                                                                        {
                                                                            nr = nr + Convert.ToInt32(blow1);
                                                                        }
                                                                    }
                                                                    if (nr == 0)
                                                                    {
                                                                        desc = "WOH";
                                                                    }
                                                                    else
                                                                    {
                                                                        desc = nr.ToString();
                                                                    }
                                                                }

                                                                MText descr_mtext = new MText();
                                                                descr_mtext.TextStyleId = HDD_textstyle.ObjectId;
                                                                descr_mtext.Contents = desc;//type1 + "-" + number1 + " (N=" + desc + ")";
                                                                descr_mtext.TextHeight = Texth;
                                                                descr_mtext.Attachment = AttachmentPoint.MiddleLeft;
                                                                descr_mtext.Location = new Point3d(pt_ins.X + 2 + wdth / 2, pt_ins.Y - (depth1 + deptht / 24) * graph_vexag * stick_vexag, pt_ins.Z);
                                                                descr_mtext.Layer = "0";
                                                                descr_mtext.Color = color_labels;
                                                                descr_mtext.BackgroundFill = true;
                                                                descr_mtext.UseBackgroundColor = true;
                                                                descr_mtext.BackgroundScaleFactor = 1.2;
                                                                descr_mtext.TransformBy(Matrix3d.Displacement(pt_ins.GetVectorTo(new Point3d(0, 0, 0))));
                                                                bltrec1.AppendEntity(descr_mtext);

                                                                MText descr_mtext2 = new MText();
                                                                descr_mtext2.TextStyleId = HDD_textstyle.ObjectId;
                                                                descr_mtext2.Contents = desc;// type1 + "-" + number1 + " (N=" + desc + ")";
                                                                descr_mtext2.TextHeight = Texth;
                                                                descr_mtext2.Attachment = AttachmentPoint.MiddleRight;
                                                                descr_mtext2.Location = new Point3d(pt_ins.X - 2 - wdth / 2, pt_ins.Y - (depth1 + deptht / 24) * graph_vexag * stick_vexag, pt_ins.Z);
                                                                descr_mtext2.Layer = "0";
                                                                descr_mtext2.Color = color_labels;
                                                                descr_mtext2.BackgroundFill = true;
                                                                descr_mtext2.UseBackgroundColor = true;
                                                                descr_mtext2.BackgroundScaleFactor = 1.2;
                                                                descr_mtext2.TransformBy(Matrix3d.Displacement(pt_ins.GetVectorTo(new Point3d(0, 0, 0))));
                                                                bltrec2.AppendEntity(descr_mtext2);
                                                            }
                                                        }
                                                    }
                                                    #endregion

                                                    #region label rock
                                                    if (dt_core_r != null && dt_core_r.Rows.Count > 0)
                                                    {
                                                        for (int j = 0; j < dt_core_r.Rows.Count; ++j)
                                                        {
                                                            string borehole2 = Convert.ToString(dt_core_r.Rows[j][col_pt]).Replace(" ", "");
                                                            if (borehole1.ToUpper() == borehole2.ToUpper())
                                                            {

                                                                string type1 = Convert.ToString(dt_core_r.Rows[j][col_type]);
                                                                string number1 = Convert.ToString(dt_core_r.Rows[j][col_number]);
                                                                double depth1 = Convert.ToDouble(dt_core_r.Rows[j][col_depth]);
                                                                double len1 = Convert.ToDouble(dt_core_r.Rows[j][col_len]);
                                                                double rec1 = Convert.ToDouble(dt_core_r.Rows[j][col_recovery]);
                                                                double rqd1 = Convert.ToDouble(dt_core_r.Rows[j][col_rqd]);
                                                                string desc = Functions.Get_String_Rounded(100 * rec1 / len1, 0) + "/" + Functions.Get_String_Rounded(100 * rqd1 / len1, 0);

                                                                if (rec1 >= rqd1 && 100 * rec1 / len1 <= 100 && 100 * rqd1 / len1 <= 100)
                                                                {
                                                                    string Mtext_content = Functions.Get_String_Rounded(100 * rqd1 / len1, 0) + "%";
                                                                    //string Mtext_content = type1 + "-" + number1 + " (" + desc + ")"
                                                                    //if(Mtext_content== "R-1 (88/57)")
                                                                    //{
                                                                    //    MessageBox.Show("investigate");
                                                                    //}

                                                                    MText descr_mtext8 = new MText();
                                                                    descr_mtext8.TextStyleId = HDD_textstyle.ObjectId;
                                                                    descr_mtext8.Attachment = AttachmentPoint.MiddleLeft;
                                                                    descr_mtext8.Contents = Mtext_content;
                                                                    descr_mtext8.TextHeight = Texth;
                                                                    descr_mtext8.Location = new Point3d(pt_ins.X + 2 + wdth / 2, pt_ins.Y - (depth1 + len1 / 24) * graph_vexag * stick_vexag, pt_ins.Z);
                                                                    descr_mtext8.Layer = "0";
                                                                    descr_mtext8.Color = color_labels;
                                                                    descr_mtext8.BackgroundFill = true;
                                                                    descr_mtext8.UseBackgroundColor = true;
                                                                    descr_mtext8.BackgroundScaleFactor = 1.2;
                                                                    descr_mtext8.TransformBy(Matrix3d.Displacement(pt_ins.GetVectorTo(new Point3d(0, 0, 0))));
                                                                    bltrec1.AppendEntity(descr_mtext8);

                                                                    MText descr_mtext2 = new MText();
                                                                    descr_mtext2.TextStyleId = HDD_textstyle.ObjectId;
                                                                    descr_mtext2.Contents = Mtext_content;
                                                                    descr_mtext2.TextHeight = Texth;
                                                                    descr_mtext2.Attachment = AttachmentPoint.MiddleRight;
                                                                    descr_mtext2.Location = new Point3d(pt_ins.X - 2 - wdth / 2, pt_ins.Y - (depth1 + len1 / 24) * graph_vexag * stick_vexag, pt_ins.Z);
                                                                    descr_mtext2.Layer = "0";
                                                                    descr_mtext2.Color = color_labels;
                                                                    descr_mtext2.BackgroundFill = true;
                                                                    descr_mtext2.UseBackgroundColor = true;
                                                                    descr_mtext2.BackgroundScaleFactor = 1.2;
                                                                    descr_mtext2.TransformBy(Matrix3d.Displacement(pt_ins.GetVectorTo(new Point3d(0, 0, 0))));
                                                                    bltrec2.AppendEntity(descr_mtext2);
                                                                }
                                                                else
                                                                {
                                                                    MText descr_mtext = new MText();
                                                                    descr_mtext.Contents = type1 + "-" + number1 + " (" + desc + ")";
                                                                    descr_mtext.TextHeight = 3 * Texth;
                                                                    descr_mtext.Attachment = AttachmentPoint.TopLeft;
                                                                    descr_mtext.Location = new Point3d(pt_ins.X + 2 + wdth / 2, pt_ins.Y - (depth1 + len1 / 24) * graph_vexag * stick_vexag, 0);
                                                                    descr_mtext.Layer = no_plot;
                                                                    descr_mtext.ColorIndex = 7;
                                                                    BTrecord.AppendEntity(descr_mtext);
                                                                    Trans1.AddNewlyCreatedDBObject(descr_mtext, true);
                                                                    MessageBox.Show("error\r\n" + borehole1 + "\r\nRECOVERY = " + rec1.ToString() + "\r\nRQD = " + rqd1.ToString() + "\r\nLENGTH = " + len1.ToString() + "\r\nSee color white label on the no plot layer");
                                                                }
                                                            }
                                                        }
                                                    }
                                                    #endregion

                                                    BlockTable1.Add(bltrec1);
                                                    Trans1.AddNewlyCreatedDBObject(bltrec1, true);

                                                    BlockTable1.Add(bltrec2);
                                                    Trans1.AddNewlyCreatedDBObject(bltrec2, true);

                                                    BlockReference b1 = Functions.InsertBlock_with_2scales(ThisDrawing.Database, BTrecord, borehole1 + idx.ToString(), pt_ins, 1, 1, 0, hdd_boreholes);
                                                    //BlockReference b2 = Functions.InsertBlock_with_2scales(ThisDrawing.Database, BTrecord, borehole1 + idx.ToString() + suff, pt_ins, 1, 1, 0, hdd_boreholes);
                                                    b1.ColorIndex = 256;
                                                    //b2.ColorIndex = 256;
                                                }
                                            }
                                            exista = false;
                                        }
                                        else
                                        {
                                            ++idx;
                                        }
                                    } while (exista == true);
                                    #endregion
                                }

                                #region LEGEND
                                int idx1 = 1;
                                string Legend = "Legend_Boreholes";
                                bool exista1 = true;
                                if (lista_legend.Count > 0)
                                {
                                    do
                                    {
                                        if (BlockTable1.Has(Legend + idx1.ToString()) == false)
                                        {
                                            using (BlockTableRecord bltrec1 = new BlockTableRecord())
                                            {
                                                bltrec1.Name = Legend + idx1.ToString();

                                                double y0 = -1;
                                                double spacing = 0.4;
                                                double width1 = 1.5;
                                                double height1 = 0.3;


                                                #region text legend stick
                                                MText descr_mtext = new MText();
                                                descr_mtext.TextStyleId = HDD_textstyle.ObjectId;
                                                descr_mtext.Contents = "27";
                                                descr_mtext.TextHeight = 0.08;
                                                descr_mtext.Attachment = AttachmentPoint.MiddleLeft;
                                                descr_mtext.Location = new Point3d(11.2562770731099, -1.08126756931543, 0);
                                                descr_mtext.Layer = "0";
                                                descr_mtext.Color = color_labels;
                                                descr_mtext.BackgroundFill = true;
                                                descr_mtext.UseBackgroundColor = true;
                                                descr_mtext.BackgroundScaleFactor = 1.2;
                                                bltrec1.AppendEntity(descr_mtext);

                                                descr_mtext = new MText();
                                                descr_mtext.TextStyleId = HDD_textstyle.ObjectId;
                                                descr_mtext.Contents = "82";
                                                descr_mtext.TextHeight = 0.08;
                                                descr_mtext.Attachment = AttachmentPoint.MiddleLeft;
                                                descr_mtext.Location = new Point3d(11.256, -1.401, 0);
                                                descr_mtext.Layer = "0";
                                                descr_mtext.Color = color_labels;
                                                descr_mtext.BackgroundFill = true;
                                                descr_mtext.UseBackgroundColor = true;
                                                descr_mtext.BackgroundScaleFactor = 1.2;
                                                bltrec1.AppendEntity(descr_mtext);


                                                descr_mtext = new MText();
                                                descr_mtext.TextStyleId = HDD_textstyle.ObjectId;
                                                descr_mtext.Contents = "70/4";
                                                descr_mtext.TextHeight = 0.08;
                                                descr_mtext.Attachment = AttachmentPoint.MiddleLeft;
                                                descr_mtext.Location = new Point3d(11.256, -1.691, 0);
                                                descr_mtext.Layer = "0";
                                                descr_mtext.Color = color_labels;
                                                descr_mtext.BackgroundFill = true;
                                                descr_mtext.UseBackgroundColor = true;
                                                descr_mtext.BackgroundScaleFactor = 1.2;
                                                bltrec1.AppendEntity(descr_mtext);

                                                descr_mtext = new MText();
                                                descr_mtext.TextStyleId = HDD_textstyle.ObjectId;
                                                descr_mtext.Contents = "63%";
                                                descr_mtext.TextHeight = 0.08;
                                                descr_mtext.Attachment = AttachmentPoint.MiddleLeft;
                                                descr_mtext.Location = new Point3d(11.256, -2.171, 0);
                                                descr_mtext.Layer = "0";
                                                descr_mtext.Color = color_labels;
                                                descr_mtext.BackgroundFill = true;
                                                descr_mtext.UseBackgroundColor = true;
                                                descr_mtext.BackgroundScaleFactor = 1.2;
                                                bltrec1.AppendEntity(descr_mtext);

                                                descr_mtext = new MText();
                                                descr_mtext.TextStyleId = HDD_textstyle.ObjectId;
                                                descr_mtext.Contents = "48%";
                                                descr_mtext.TextHeight = 0.08;
                                                descr_mtext.Attachment = AttachmentPoint.MiddleLeft;
                                                descr_mtext.Location = new Point3d(11.256, -2.4965, 0);
                                                descr_mtext.Layer = "0";
                                                descr_mtext.Color = color_labels;
                                                descr_mtext.BackgroundFill = true;
                                                descr_mtext.UseBackgroundColor = true;
                                                descr_mtext.BackgroundScaleFactor = 1.2;
                                                bltrec1.AppendEntity(descr_mtext);
                                                #endregion

                                                MText mt1 = new MText();
                                                mt1.Location = new Point3d(width1 / 2, 0, 0);
                                                mt1.Attachment = AttachmentPoint.BottomCenter;
                                                mt1.TextHeight = 0.2;
                                                mt1.Contents = "LEGEND";
                                                mt1.Layer = "0";
                                                bltrec1.AppendEntity(mt1);




                                                Circle cerc_bh = new Circle(new Point3d(width1 / 2, -0.4, 0), Vector3d.ZAxis, 0.08);
                                                cerc_bh.Layer = "0";
                                                cerc_bh.Color = color_sym;
                                                bltrec1.AppendEntity(cerc_bh);

                                                Polyline poly_sym = get_poly_sym(0.08);
                                                poly_sym.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(width1 / 2, -0.4, 0))));
                                                poly_sym.Layer = "0";
                                                poly_sym.Color = color_sym;
                                                bltrec1.AppendEntity(poly_sym);

                                                MText mt3 = new MText();
                                                mt3.Location = new Point3d(width1 + 0.2, -0.4, 0);
                                                mt3.Attachment = AttachmentPoint.MiddleLeft;
                                                mt3.TextHeight = 0.08;
                                                mt3.Contents = "BOREHOLE";
                                                mt3.Layer = "0";
                                                bltrec1.AppendEntity(mt3);


                                                #region mleader top
                                                string content1 = "EL. 1312" + f + "\r\n" + "513" + f + " LT.";

                                                MLeader mleader_top = Functions.creaza_mleader_with_style_IN_BTR(new Point3d(11.094, -0.961, 0), content1,
                                                                                                            0.06, mleaderstyle_name, textstyle_name,
                                                                                                                    0.3, 0.25, hdd_boreholes, bltrec1, Trans1, BTrecord);

                                                mleader_top.SetTextAttachmentType(TextAttachmentType.AttachmentBottomOfTopLine, LeaderDirectionType.RightLeader);
                                                mleader_top.SetTextAttachmentType(TextAttachmentType.AttachmentBottomOfTopLine, LeaderDirectionType.LeftLeader);
                                                mleader_top.ArrowSymbolId = ObjectId.Null;

                                                mleader_top.LineWeight = LineWeight.LineWeight000;
                                                mleader_top.LeaderLineWeight = LineWeight.LineWeight000;
                                                mleader_top.Color = color_labels;

                                                #endregion

                                                #region borehole stick legend label
                                                MText mtext_borehole = new MText();
                                                mtext_borehole.Contents = "B-MP-2.2-C";
                                                mtext_borehole.TextHeight = 0.12;
                                                mtext_borehole.TextStyleId = HDD_textstyle.ObjectId;
                                                mtext_borehole.Attachment = AttachmentPoint.MiddleCenter;
                                                mtext_borehole.Location = new Point3d(11.09438, -0.496267569315433, 0);
                                                mtext_borehole.Layer = "0";
                                                mtext_borehole.Color = color_labels;
                                                mtext_borehole.BackgroundFill = true;
                                                mtext_borehole.UseBackgroundColor = true;
                                                mtext_borehole.BackgroundScaleFactor = 1.2;

                                                bltrec1.AppendEntity(mtext_borehole);
                                                #endregion




                                                #region border around label legend
                                                Polyline poly_top = new Polyline();
                                                poly_top.AddVertexAt(0, new Point2d(11.09, -0.96), 0, 0.06 / 20, 0.06 / 20);
                                                poly_top.AddVertexAt(1, new Point2d(11.09, -0.59), 0, 0.06 / 20, 0.06 / 20);
                                                poly_top.AddVertexAt(2, new Point2d(11.56, -0.59), 0, 0.06 / 20, 0.06 / 20);
                                                poly_top.AddVertexAt(3, new Point2d(11.56, -0.4), 0, 0.06 / 20, 0.06 / 20);
                                                poly_top.AddVertexAt(4, new Point2d(10.63, -0.4), 0, 0.06 / 20, 0.06 / 20);
                                                poly_top.AddVertexAt(5, new Point2d(10.63, -0.59), 0, 0.06 / 20, 0.06 / 20);
                                                poly_top.AddVertexAt(6, new Point2d(11.09, -0.59), 0, 0.06 / 20, 0.06 / 20);
                                                poly_top.Layer = "0";
                                                poly_top.ColorIndex = 7;
                                                //poly_top.LineWeight = LineWeight.LineWeight000;

                                                bltrec1.AppendEntity(poly_top);
                                                #endregion



                                                #region stick legend
                                                Polyline polyL1 = new Polyline();
                                                polyL1.AddVertexAt(0, new Point2d(11, -0.96), 0, 0, 0);
                                                polyL1.AddVertexAt(1, new Point2d(11.18, -0.96), 0, 0, 0);
                                                polyL1.AddVertexAt(2, new Point2d(11.18, -1.12), 0, 0, 0);
                                                polyL1.AddVertexAt(3, new Point2d(11, -1.12), 0, 0, 0);
                                                polyL1.Closed = true;
                                                polyL1.Layer = "0";
                                                polyL1.LineWeight = LineWeight.LineWeight000;
                                                polyL1.Color = color_rect;
                                                bltrec1.AppendEntity(polyL1);
                                                add_pattern_topsoil(bltrec1, 1, 1, 1, polyL1, BTrecord, Trans1);

                                                polyL1 = new Polyline();
                                                polyL1.AddVertexAt(0, new Point2d(11, -1.12), 0, 0, 0);
                                                polyL1.AddVertexAt(1, new Point2d(11.18, -1.12), 0, 0, 0);
                                                polyL1.AddVertexAt(2, new Point2d(11.18, -1.56), 0, 0, 0);
                                                polyL1.AddVertexAt(3, new Point2d(11, -1.56), 0, 0, 0);
                                                polyL1.Closed = true;
                                                polyL1.Layer = "0";
                                                polyL1.LineWeight = LineWeight.LineWeight000;
                                                polyL1.Color = color_rect;
                                                bltrec1.AppendEntity(polyL1);
                                                add_pattern_SP(bltrec1, 1, 1, 1, polyL1, BTrecord, Trans1);


                                                polyL1 = new Polyline();
                                                polyL1.AddVertexAt(0, new Point2d(11, -1.56), 0, 0, 0);
                                                polyL1.AddVertexAt(1, new Point2d(11.18, -1.56), 0, 0, 0);
                                                polyL1.AddVertexAt(2, new Point2d(11.18, -1.96), 0, 0, 0);
                                                polyL1.AddVertexAt(3, new Point2d(11, -1.96), 0, 0, 0);
                                                polyL1.Closed = true;
                                                polyL1.Layer = "0";
                                                polyL1.LineWeight = LineWeight.LineWeight000;
                                                polyL1.Color = color_rect;
                                                bltrec1.AppendEntity(polyL1);
                                                add_pattern_SM(bltrec1, 1, 1, 1, polyL1, BTrecord, Trans1);

                                                polyL1 = new Polyline();
                                                polyL1.AddVertexAt(0, new Point2d(11, -1.96), 0, 0, 0);
                                                polyL1.AddVertexAt(1, new Point2d(11.18, -1.96), 0, 0, 0);
                                                polyL1.AddVertexAt(2, new Point2d(11.18, -2.36), 0, 0, 0);
                                                polyL1.AddVertexAt(3, new Point2d(11, -2.36), 0, 0, 0);
                                                polyL1.Closed = true;
                                                polyL1.Layer = "0";
                                                polyL1.LineWeight = LineWeight.LineWeight000;
                                                polyL1.Color = color_rect;
                                                bltrec1.AppendEntity(polyL1);
                                                add_pattern_GC_legend_2(bltrec1, 1, 1, 1, polyL1, BTrecord, Trans1);


                                                polyL1 = new Polyline();
                                                polyL1.AddVertexAt(0, new Point2d(11, -2.36), 0, 0, 0);
                                                polyL1.AddVertexAt(1, new Point2d(11.18, -2.36), 0, 0, 0);
                                                polyL1.AddVertexAt(2, new Point2d(11.18, -2.6), 0, 0, 0);
                                                polyL1.AddVertexAt(3, new Point2d(11, -2.6), 0, 0, 0);
                                                polyL1.Closed = true;
                                                polyL1.Layer = "0";
                                                polyL1.LineWeight = LineWeight.LineWeight000;
                                                polyL1.Color = color_rect;
                                                bltrec1.AppendEntity(polyL1);
                                                add_pattern_Shale(bltrec1, 1, 1, 1, polyL1, BTrecord, Trans1);
                                                #endregion


                                                #region mleader explanation2
                                                string content2 = "TEST BORING LABEL";

                                                MLeader mleader2 = Functions.creaza_mleader_with_style_IN_BTR(new Point3d(11.50338, -0.43627, 0), content2,
                                                                                                            0.06, mleaderstyle_name, textstyle_name,
                                                                                                                    0.0644, 0.3939, hdd_boreholes, bltrec1, Trans1, BTrecord);

                                                mleader2.SetTextAttachmentType(TextAttachmentType.AttachmentMiddle, LeaderDirectionType.RightLeader);
                                                mleader2.SetTextAttachmentType(TextAttachmentType.AttachmentMiddle, LeaderDirectionType.LeftLeader);

                                                mleader2.ArrowSize = 0.06;
                                                mleader2.EnableLanding = true;
                                                mleader2.LandingGap = 0.06;
                                                mleader2.DoglegLength = 0.06;

                                                mleader2.ArrowSymbolId = Arrowid_dot;


                                                mleader2.LineWeight = LineWeight.LineWeight000;
                                                mleader2.LeaderLineWeight = LineWeight.LineWeight000;
                                                mleader2.Color = color_explanations;


                                                #endregion

                                                #region mleader explanation3
                                                string content3 = "GROUND SURFACE ELEVATION";

                                                MLeader mleader3 = Functions.creaza_mleader_with_style_IN_BTR(new Point3d(11.70008, -0.65487, 0), content3,
                                                                                                             0.06, mleaderstyle_name, textstyle_name,
                                                                                                                     0.1972, 0.1417, hdd_boreholes, bltrec1, Trans1, BTrecord);

                                                mleader3.SetTextAttachmentType(TextAttachmentType.AttachmentMiddle, LeaderDirectionType.RightLeader);
                                                mleader3.SetTextAttachmentType(TextAttachmentType.AttachmentMiddle, LeaderDirectionType.LeftLeader);

                                                mleader3.ArrowSize = 0.06;
                                                mleader3.EnableLanding = true;
                                                mleader3.LandingGap = 0.06;
                                                mleader3.DoglegLength = 0.06;

                                                mleader3.ArrowSymbolId = Arrowid_dot;


                                                mleader3.LineWeight = LineWeight.LineWeight000;
                                                mleader3.LeaderLineWeight = LineWeight.LineWeight000;
                                                mleader3.Color = color_explanations;


                                                #endregion

                                                #region mleader explanation4
                                                string content4 = "SPT " + Convert.ToString(Convert.ToChar(34)) + "N" + Convert.ToString(Convert.ToChar(34)) + " VALUE";

                                                MLeader mleader4 = Functions.creaza_mleader_with_style_IN_BTR(new Point3d(11.37398, -1.04447, 0), content4,
                                                                                                             0.06, mleaderstyle_name, textstyle_name,
                                                                                                                     0.5233, 0.1719, hdd_boreholes, bltrec1, Trans1, BTrecord);

                                                mleader4.SetTextAttachmentType(TextAttachmentType.AttachmentMiddle, LeaderDirectionType.RightLeader);
                                                mleader4.SetTextAttachmentType(TextAttachmentType.AttachmentMiddle, LeaderDirectionType.LeftLeader);

                                                mleader4.ArrowSize = 0.06;
                                                mleader4.EnableLanding = true;
                                                mleader4.LandingGap = 0.06;
                                                mleader4.DoglegLength = 0.06;

                                                mleader4.ArrowSymbolId = Arrowid_dot;


                                                mleader4.LineWeight = LineWeight.LineWeight000;
                                                mleader4.LeaderLineWeight = LineWeight.LineWeight000;
                                                mleader4.Color = color_explanations;


                                                #endregion

                                                #region mleader explanation5
                                                string content5 = "OFFSET FROM PROFILE LINE";

                                                MLeader mleader5 = Functions.creaza_mleader_with_style_IN_BTR(new Point3d(11.67634, -0.76347, 0), content5,
                                                                                                             0.06, mleaderstyle_name, textstyle_name,
                                                                                                                     0.2183, 0.086, hdd_boreholes, bltrec1, Trans1, BTrecord);

                                                mleader5.SetTextAttachmentType(TextAttachmentType.AttachmentMiddle, LeaderDirectionType.RightLeader);
                                                mleader5.SetTextAttachmentType(TextAttachmentType.AttachmentMiddle, LeaderDirectionType.LeftLeader);

                                                mleader5.ArrowSize = 0.06;
                                                mleader5.EnableLanding = true;
                                                mleader5.LandingGap = 0.06;
                                                mleader5.DoglegLength = 0.06;

                                                mleader5.ArrowSymbolId = Arrowid_dot;


                                                mleader5.LineWeight = LineWeight.LineWeight000;
                                                mleader5.LeaderLineWeight = LineWeight.LineWeight000;
                                                mleader5.Color = color_explanations;


                                                #endregion

                                                #region mleader explanation6
                                                string content6 = "ROCK QUALITY DESIGNATION (RQD)";

                                                MLeader mleader6 = Functions.creaza_mleader_with_style_IN_BTR(new Point3d(11.45318, -2.12967, 0), content6,
                                                                                                             0.06, mleaderstyle_name, textstyle_name,
                                                                                                                     0.4441, 0.2648, hdd_boreholes, bltrec1, Trans1, BTrecord);

                                                mleader6.SetTextAttachmentType(TextAttachmentType.AttachmentMiddle, LeaderDirectionType.RightLeader);
                                                mleader6.SetTextAttachmentType(TextAttachmentType.AttachmentMiddle, LeaderDirectionType.LeftLeader);

                                                mleader6.ArrowSize = 0.06;
                                                mleader6.EnableLanding = true;
                                                mleader6.LandingGap = 0.06;
                                                mleader6.DoglegLength = 0.06;

                                                mleader6.ArrowSymbolId = Arrowid_dot;


                                                mleader6.LineWeight = LineWeight.LineWeight000;
                                                mleader6.LeaderLineWeight = LineWeight.LineWeight000;
                                                mleader6.Color = color_explanations;


                                                #endregion

                                                #region mleader explanation7
                                                string content7 = "TEST BORING STRATA SYMBOL";

                                                MLeader mleader7 = Functions.creaza_mleader_with_style_IN_BTR(new Point3d(11.10578, -1.32147, 0), content7,
                                                                                                             0.06, mleaderstyle_name, textstyle_name,
                                                                                                                     0.7888, 0.2861, hdd_boreholes, bltrec1, Trans1, BTrecord);

                                                mleader7.SetTextAttachmentType(TextAttachmentType.AttachmentMiddle, LeaderDirectionType.RightLeader);
                                                mleader7.SetTextAttachmentType(TextAttachmentType.AttachmentMiddle, LeaderDirectionType.LeftLeader);

                                                mleader7.ArrowSize = 0.06;
                                                mleader7.EnableLanding = true;
                                                mleader7.LandingGap = 0.06;
                                                mleader7.DoglegLength = 0.06;

                                                mleader7.ArrowSymbolId = Arrowid_dot;


                                                mleader7.LineWeight = LineWeight.LineWeight000;
                                                mleader7.LeaderLineWeight = LineWeight.LineWeight000;
                                                mleader7.Color = color_explanations;


                                                #endregion


                                                for (int i = 0; i < lista_legend.Count; i++)
                                                {
                                                    Polyline poly1 = new Polyline();
                                                    poly1.AddVertexAt(0, new Point2d(0, y0 - i * spacing), 0, 0, 0);
                                                    poly1.AddVertexAt(1, new Point2d(width1, y0 - i * spacing), 0, 0, 0);
                                                    poly1.AddVertexAt(2, new Point2d(width1, y0 - i * spacing - height1), 0, 0, 0);
                                                    poly1.AddVertexAt(3, new Point2d(0, y0 - i * spacing - height1), 0, 0, 0);
                                                    poly1.Closed = true;
                                                    poly1.Layer = "0";
                                                    poly1.LineWeight = LineWeight.LineWeight000;
                                                    poly1.Color = color_rect;
                                                    bltrec1.AppendEntity(poly1);

                                                    MText mt2 = new MText();
                                                    mt2.Location = new Point3d(width1 + 0.2, y0 - i * spacing - 0.5 * height1, 0);
                                                    mt2.Attachment = AttachmentPoint.MiddleLeft;
                                                    mt2.TextHeight = 0.08;
                                                    mt2.Contents = lista_legend[i];
                                                    mt2.Layer = "0";
                                                    bltrec1.AppendEntity(mt2);

                                                    if (lista_legend[i] == "TOPSOIL")
                                                    {
                                                        add_pattern_topsoil(bltrec1, 1, 1, 1, poly1, BTrecord, Trans1);
                                                    }

                                                    if (lista_legend[i] == "GM")
                                                    {
                                                        add_pattern_GM_legend(bltrec1, 1, 1, 1, poly1, BTrecord, Trans1);
                                                    }


                                                    if (lista_legend[i] == "ML")
                                                    {
                                                        add_pattern_ML(bltrec1, 1, poly1, BTrecord, Trans1);

                                                    }

                                                    if (lista_legend[i] == "CL")
                                                    {
                                                        add_pattern_CL(bltrec1, 1, poly1, BTrecord, Trans1);
                                                    }

                                                    if (lista_legend[i] == "SANDSTONE")
                                                    {
                                                        add_pattern_sandstone(bltrec1, 1, 1, 1, poly1, BTrecord, Trans1);

                                                    }

                                                    if (lista_legend[i] == "SM")
                                                    {
                                                        add_pattern_SM(bltrec1, 1, 1, 1, poly1, BTrecord, Trans1);
                                                    }
                                                    if (lista_legend[i] == "SHALE")
                                                    {
                                                        add_pattern_Shale_legend(bltrec1, 1, 1, 1, poly1, BTrecord, Trans1);
                                                    }
                                                    if (lista_legend[i] == "SP")
                                                    {
                                                        add_pattern_SP(bltrec1, 1, 1, 1, poly1, BTrecord, Trans1);
                                                    }
                                                    if (lista_legend[i] == "GP")
                                                    {
                                                        add_pattern_GP_legend(bltrec1, 1, 1, 1, poly1, BTrecord, Trans1);
                                                    }
                                                    if (lista_legend[i] == "GC")
                                                    {
                                                        add_pattern_GC(bltrec1, 1, 1, 1, poly1, BTrecord, Trans1);
                                                    }
                                                    if (lista_legend[i] == "MUDSTONE")
                                                    {
                                                        add_pattern_MUDSTONE_legend(bltrec1, 1, 1, 1, poly1, BTrecord, Trans1);
                                                    }
                                                    if (lista_legend[i] == "CL-ML")
                                                    {
                                                        add_pattern_CLML(bltrec1, 1, poly1, BTrecord, Trans1);
                                                    }
                                                    if (lista_legend[i] == "GP-GC")
                                                    {
                                                        add_pattern_GPGC_legend(bltrec1, 1, 1, 1, poly1, BTrecord, Trans1);
                                                    }
                                                    if (lista_legend[i] == "SC")
                                                    {
                                                        add_pattern_SC(bltrec1, 1, 1, 1, poly1, BTrecord, Trans1);
                                                    }
                                                }

                                                BlockTable1.Add(bltrec1);
                                                Trans1.AddNewlyCreatedDBObject(bltrec1, true);
                                                BlockReference b1 = Functions.InsertBlock_with_2scales(ThisDrawing.Database, BTrecord, Legend + idx1.ToString(), pt_legend, 1, 1, 0, "TEXT");
                                                b1.ColorIndex = 256;
                                            }
                                            exista1 = false;
                                        }
                                        else
                                        {
                                            ++idx1;
                                        }
                                    } while (exista1 == true);
                                    #endregion
                                }
                            }

                            Trans1.Commit();
                        }
                    }
                    Editor1.SetImpliedSelection(Empty_array);
                    Editor1.WriteMessage("\nCommand:");
                }
            }
            catch (System.Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            set_enable_true();
            this.MdiParent.WindowState = FormWindowState.Normal;
        }


        private void add_pattern_GC(BlockTableRecord bltrec1, double scale1, double graph_vexag, double stick_vexag, Polyline poly1, BlockTableRecord BTrecord, Autodesk.AutoCAD.DatabaseServices.Transaction Trans1)
        {



            Autodesk.AutoCAD.Colors.Color color_GC = Autodesk.AutoCAD.Colors.Color.FromRgb(127, 95, 93);



            double pattern_width = 0.33 * scale1;
            double pattern_height = 0.4 * scale1;

            double spc_h_edge = 0;
            double spc_v_edge = 0;
            double spc_hor = 0;
            double spc_ver = 0;

            int nr_col = 0;
            int nr_rows = 1;

            double x1 = poly1.GetPoint2dAt(3).X + spc_h_edge;
            double y1 = poly1.GetPoint2dAt(3).Y + spc_v_edge;


            double stick_width = poly1.GetPoint2dAt(1).X - poly1.GetPoint2dAt(0).X;
            double stick_height = poly1.GetPoint2dAt(1).Y - poly1.GetPoint2dAt(2).Y;



            if (stick_height <= pattern_height)
            {
                nr_rows = 1;
            }
            else
            {
                double nr1 = Math.Floor((stick_height - 2 * spc_v_edge) / (pattern_height + spc_ver));
                nr_rows = Convert.ToInt32(nr1);
            }

            if (stick_width - 2 * spc_h_edge < pattern_width + 2 * spc_hor)
            {
                nr_col = 1;
            }
            else
            {
                double nr2 = Math.Floor((stick_width - 2 * spc_h_edge) / (pattern_width + spc_hor));

                nr_col = Convert.ToInt32(nr2);
            }

            double dif_len = stick_width - (nr_col * (pattern_width + spc_hor) - spc_hor);
            double dif_hght = stick_height - (nr_rows * (pattern_height + spc_ver) - spc_ver);

            Point3d pt_ins = poly1.GetPoint3dAt(3);

            if (nr_rows > 0 && nr_col > 0)
            {
                for (int m = 0; m < nr_col; ++m)
                {
                    for (int n = 0; n < nr_rows; ++n)
                    {
                        double x2 = x1 + dif_len / 2 + m * (pattern_width + spc_hor);
                        double y2 = y1 + dif_hght / 2 + n * (pattern_height + spc_ver);

                        double x3 = x2 + pattern_width;
                        double y3 = y2;


                        Polyline polygc1 = get_poly_gc1(scale1);
                        polygc1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));
                        polygc1.Layer = "0";
                        polygc1.Color = color_GC;
                        polygc1.LineWeight = LineWeight.LineWeight000;
                        bltrec1.AppendEntity(polygc1);

                        Polyline polygc2 = get_poly_gc2(scale1);
                        polygc2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                        polygc2.Layer = "0";
                        polygc2.Color = color_GC;
                        polygc2.LineWeight = LineWeight.LineWeight000;
                        bltrec1.AppendEntity(polygc2);

                        Polyline polygc3 = get_poly_gc3(scale1);
                        polygc3.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                        polygc3.Layer = "0";
                        polygc3.Color = color_GC;
                        polygc3.LineWeight = LineWeight.LineWeight000;
                        bltrec1.AppendEntity(polygc3);

                        Polyline polygc4 = get_poly_gc4(scale1);
                        polygc4.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                        polygc4.Layer = "0";
                        polygc4.Color = color_GC;
                        polygc4.LineWeight = LineWeight.LineWeight000;
                        bltrec1.AppendEntity(polygc4);

                        Polyline polygc5 = get_poly_gc5(scale1);
                        polygc5.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                        polygc5.Layer = "0";
                        polygc5.Color = color_GC;
                        polygc5.LineWeight = LineWeight.LineWeight000;
                        bltrec1.AppendEntity(polygc5);

                        Polyline polygc6 = get_poly_gc6(scale1);
                        polygc6.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                        polygc6.Layer = "0";
                        polygc6.Color = color_GC;
                        polygc6.LineWeight = LineWeight.LineWeight000;
                        bltrec1.AppendEntity(polygc6);

                        Polyline polygc7 = get_poly_gc7(scale1);
                        polygc7.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                        polygc7.Layer = "0";
                        polygc7.Color = color_GC;
                        polygc7.LineWeight = LineWeight.LineWeight000;
                        bltrec1.AppendEntity(polygc7);

                        Polyline polygc8 = get_poly_gc8(scale1);
                        polygc8.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                        polygc8.Layer = "0";
                        polygc8.Color = color_GC;
                        polygc8.LineWeight = LineWeight.LineWeight000;
                        bltrec1.AppendEntity(polygc8);

                        Polyline polygc9 = get_poly_gc9(scale1);
                        polygc9.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                        polygc9.Layer = "0";
                        polygc9.Color = color_GC;
                        polygc9.LineWeight = LineWeight.LineWeight000;
                        bltrec1.AppendEntity(polygc9);

                        Polyline polygc10 = get_poly_gc10(scale1);
                        polygc10.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                        polygc10.Layer = "0";
                        polygc10.Color = color_GC;
                        polygc10.LineWeight = LineWeight.LineWeight000;
                        bltrec1.AppendEntity(polygc10);

                    }
                }



            }

            else
            {
            }
            string nume_hatch = "ANSI31";
            double hatch_scale = scale1 / 2;
            double hatch_angle = 0;


            Polyline poly2 = new Polyline();
            poly2 = poly1.Clone() as Polyline;
            BTrecord.AppendEntity(poly2);
            Trans1.AddNewlyCreatedDBObject(poly2, true);

            Hatch hatch1 = CreateHatch(poly2, nume_hatch, hatch_scale, hatch_angle * Math.PI / 180);
            hatch1.Layer = "0";
            hatch1.LineWeight = LineWeight.LineWeight000;
            hatch1.Color = color_GC;
            bltrec1.AppendEntity(hatch1);
            poly2.Erase();

        }


        private void add_pattern_GC_legend_2(BlockTableRecord bltrec1, double scale1, double graph_vexag, double stick_vexag, Polyline poly1, BlockTableRecord BTrecord, Autodesk.AutoCAD.DatabaseServices.Transaction Trans1)
        {



            Autodesk.AutoCAD.Colors.Color color_GC = Autodesk.AutoCAD.Colors.Color.FromRgb(127, 95, 93);



            double pattern_width = 0.18 * scale1;
            double pattern_height = 0.4 * scale1;

            double spc_h_edge = 0;
            double spc_v_edge = 0;
            double spc_hor = 0;
            double spc_ver = 0;

            int nr_col = 1;
            int nr_rows = 1;

            double x1 = poly1.GetPoint2dAt(3).X + spc_h_edge;
            double y1 = poly1.GetPoint2dAt(3).Y + spc_v_edge;


            double stick_width = poly1.GetPoint2dAt(1).X - poly1.GetPoint2dAt(0).X;
            double stick_height = poly1.GetPoint2dAt(1).Y - poly1.GetPoint2dAt(2).Y;





            double dif_len = 0;
            double dif_hght = 0;

            Point3d pt_ins = poly1.GetPoint3dAt(3);

            if (nr_rows > 0 && nr_col > 0)
            {
                for (int m = 0; m < nr_col; ++m)
                {
                    for (int n = 0; n < nr_rows; ++n)
                    {
                        double x2 = x1 + dif_len / 2 + m * (pattern_width + spc_hor);
                        double y2 = y1 + dif_hght / 2 + n * (pattern_height + spc_ver);

                        double x3 = x2 + pattern_width;
                        double y3 = y2;


                        Polyline polygc1 = get_poly_gc1l(scale1);

                        polygc1.Layer = "0";
                        polygc1.Color = color_GC;
                        polygc1.LineWeight = LineWeight.LineWeight000;
                        bltrec1.AppendEntity(polygc1);

                        Polyline polygc2 = get_poly_gc2l(scale1);

                        polygc2.Layer = "0";
                        polygc2.Color = color_GC;
                        polygc2.LineWeight = LineWeight.LineWeight000;
                        bltrec1.AppendEntity(polygc2);

                        Polyline polygc3 = get_poly_gc3l(scale1);

                        polygc3.Layer = "0";
                        polygc3.Color = color_GC;
                        polygc3.LineWeight = LineWeight.LineWeight000;
                        bltrec1.AppendEntity(polygc3);

                        Polyline polygc4 = get_poly_gc4l(scale1);

                        polygc4.Layer = "0";
                        polygc4.Color = color_GC;
                        polygc4.LineWeight = LineWeight.LineWeight000;
                        bltrec1.AppendEntity(polygc4);

                        Polyline polygc5 = get_poly_gc5l(scale1);

                        polygc5.Layer = "0";
                        polygc5.Color = color_GC;
                        polygc5.LineWeight = LineWeight.LineWeight000;
                        bltrec1.AppendEntity(polygc5);

                        Polyline polygc6 = get_poly_gc6l(scale1);

                        polygc6.Layer = "0";
                        polygc6.Color = color_GC;
                        polygc6.LineWeight = LineWeight.LineWeight000;
                        bltrec1.AppendEntity(polygc6);

                        Polyline polygc7 = get_poly_gc7l(scale1);

                        polygc7.Layer = "0";
                        polygc7.Color = color_GC;
                        polygc7.LineWeight = LineWeight.LineWeight000;
                        bltrec1.AppendEntity(polygc7);

                        Polyline polygc8 = get_poly_gc8l(scale1);

                        polygc8.Layer = "0";
                        polygc8.Color = color_GC;
                        polygc8.LineWeight = LineWeight.LineWeight000;
                        bltrec1.AppendEntity(polygc8);

                        Polyline polygc9 = get_poly_gc9l(scale1);

                        polygc9.Layer = "0";
                        polygc9.Color = color_GC;
                        polygc9.LineWeight = LineWeight.LineWeight000;
                        bltrec1.AppendEntity(polygc9);

                        Polyline polygc10 = get_poly_gc10l(scale1);

                        polygc10.Layer = "0";
                        polygc10.Color = color_GC;
                        polygc10.LineWeight = LineWeight.LineWeight000;
                        bltrec1.AppendEntity(polygc10);

                        Polyline polygc11 = get_poly_gc11l(scale1);

                        polygc11.Layer = "0";
                        polygc11.Color = color_GC;
                        polygc11.LineWeight = LineWeight.LineWeight000;
                        bltrec1.AppendEntity(polygc11);

                        Polyline polygc12 = get_poly_gc12l(scale1);

                        polygc12.Layer = "0";
                        polygc12.Color = color_GC;
                        polygc12.LineWeight = LineWeight.LineWeight000;
                        bltrec1.AppendEntity(polygc12);

                        Polyline polygc13 = get_poly_gc13l(scale1);

                        polygc13.Layer = "0";
                        polygc13.Color = color_GC;
                        polygc13.LineWeight = LineWeight.LineWeight000;
                        bltrec1.AppendEntity(polygc13);

                    }
                }



            }

            else
            {
            }
            string nume_hatch = "ANSI31";
            double hatch_scale = scale1 / 2;
            double hatch_angle = 0;


            Polyline poly2 = new Polyline();
            poly2 = poly1.Clone() as Polyline;
            BTrecord.AppendEntity(poly2);
            Trans1.AddNewlyCreatedDBObject(poly2, true);

            Hatch hatch1 = CreateHatch(poly2, nume_hatch, hatch_scale, hatch_angle * Math.PI / 180);
            hatch1.Layer = "0";
            hatch1.LineWeight = LineWeight.LineWeight000;
            hatch1.Color = color_GC;
            bltrec1.AppendEntity(hatch1);
            poly2.Erase();

        }



        private void add_pattern_MUDSTONE_legend(BlockTableRecord bltrec1, double scale1, double graph_vexag, double stick_vexag, Polyline poly1, BlockTableRecord BTrecord, Autodesk.AutoCAD.DatabaseServices.Transaction Trans1)
        {


            Autodesk.AutoCAD.Colors.Color color1 = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 255, 255);
            Autodesk.AutoCAD.Colors.Color color_MDST = Autodesk.AutoCAD.Colors.Color.FromRgb(127, 95, 93);


            double ins_x = poly1.GetPoint2dAt(3).X;
            double ins_y = poly1.GetPoint2dAt(3).Y;


            string nume_hatch = "ANSI31";
            double hatch_scale = 0.23 * scale1;
            double hatch_angle = 0;

            Polyline poly3 = new Polyline();
            poly3 = poly1.Clone() as Polyline;
            BTrecord.AppendEntity(poly3);
            Trans1.AddNewlyCreatedDBObject(poly3, true);

            Hatch hatch1 = CreateHatch(poly3, nume_hatch, hatch_scale, hatch_angle * Math.PI / 180);
            hatch1.Layer = "0";
            hatch1.LineWeight = LineWeight.LineWeight000;
            hatch1.Color = color_MDST;
            bltrec1.AppendEntity(hatch1);
            poly3.Erase();



            Polyline poly2 = new Polyline();

            #region poly1
            poly2 = new Polyline();
            poly2.AddVertexAt(0, new Point2d(scale1 * 0.160255200661309, scale1 * 0), 0, 0.04 * scale1, 0.04 * scale1);
            poly2.AddVertexAt(1, new Point2d(scale1 * 0.0733566503040491, scale1 * 0.130596833303571), 0, 0.04 * scale1, 0.04 * scale1);
            poly2.AddVertexAt(2, new Point2d(scale1 * 0.0340442184824502, scale1 * 0.219096973538399), 0, 0.04 * scale1, 0.04 * scale1);
            poly2.AddVertexAt(3, new Point2d(scale1 * 0.0584125409368421, scale1 * 0.233223754912614), 0, 0.04 * scale1, 0.04 * scale1);
            poly2.AddVertexAt(4, new Point2d(scale1 * 0.285670708864927, scale1 * 0.222697917371988), 0, 0.04 * scale1, 0.04 * scale1);
            poly2.AddVertexAt(5, new Point2d(scale1 * 0.356486938893795, scale1 * 0.205731928348541), 0, 0.04 * scale1, 0.04 * scale1);
            poly2.AddVertexAt(6, new Point2d(scale1 * 0.451402228325605, scale1 * 0.0944488886743784), 0, 0.04 * scale1, 0.04 * scale1);
            poly2.AddVertexAt(7, new Point2d(scale1 * 0.481425078585744, scale1 * 0.0218067541718483), 0, 0.04 * scale1, 0.04 * scale1);
            poly2.AddVertexAt(8, new Point2d(scale1 * 0.486610263233463, scale1 * 0), 0, 0.04 * scale1, 0.04 * scale1);

            poly2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(ins_x, ins_y, 0))));

            poly2.Layer = "0";
            poly2.Color = color1;

            poly2.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly2);
            #endregion

            #region poly2
            poly2 = new Polyline();
            poly2.AddVertexAt(0, new Point2d(scale1 * 0.430097049102187, scale1 * 0.119428057223558), 0, 0.04 * scale1, 0.04 * scale1);
            poly2.AddVertexAt(1, new Point2d(scale1 * 0.592630793573335, scale1 * 0.274080628529191), 0, 0.04 * scale1, 0.04 * scale1);
            poly2.AddVertexAt(2, new Point2d(scale1 * 0.584675989346579, scale1 * 0.307032341137528), 0, 0.04 * scale1, 0.04 * scale1);

            poly2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(ins_x, ins_y, 0))));

            poly2.Layer = "0";
            poly2.Color = color1;

            poly2.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly2);
            #endregion

            #region poly3
            poly2 = new Polyline();
            poly2.AddVertexAt(0, new Point2d(scale1 * 0.0311776630114761, scale1 * 0.307032341137528), 0, 0.04 * scale1, 0.04 * scale1);
            poly2.AddVertexAt(1, new Point2d(scale1 * 0.0340442184824502, scale1 * 0.219096973538399), 0, 0.04 * scale1, 0.04 * scale1);

            poly2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(ins_x, ins_y, 0))));

            poly2.Layer = "0";
            poly2.Color = color1;

            poly2.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly2);
            #endregion

            #region poly4
            poly2 = new Polyline();
            poly2.AddVertexAt(0, new Point2d(scale1 * 0.725778054308964, scale1 * 0), 0, 0.04 * scale1, 0.04 * scale1);
            poly2.AddVertexAt(1, new Point2d(scale1 * 0.720947858178989, scale1 * 0.0407117139548063), 0, 0.04 * scale1, 0.04 * scale1);
            poly2.AddVertexAt(2, new Point2d(scale1 * 0.592630793573335, scale1 * 0.274080628529191), 0, 0.04 * scale1, 0.04 * scale1);

            poly2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(ins_x, ins_y, 0))));

            poly2.Layer = "0";
            poly2.Color = color1;

            poly2.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly2);
            #endregion

            #region poly5
            poly2 = new Polyline();
            poly2.AddVertexAt(0, new Point2d(scale1 * 0.677539282245559, scale1 * 0.119658440351486), 0, 0.04 * scale1, 0.04 * scale1);
            poly2.AddVertexAt(1, new Point2d(scale1 * 0.833917956100779, scale1 * 0.278783425688744), 0, 0.04 * scale1, 0.04 * scale1);

            poly2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(ins_x, ins_y, 0))));

            poly2.Layer = "0";
            poly2.Color = color1;

            poly2.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly2);
            #endregion

            #region poly6
            poly2 = new Polyline();
            poly2.AddVertexAt(0, new Point2d(scale1 * 0.855120390187949, scale1 * 0.166141012683511), 0, 0.04 * scale1, 0.04 * scale1);
            poly2.AddVertexAt(1, new Point2d(scale1 * 0.995852520456539, scale1 * 0.307032341137528), 0, 0.04 * scale1, 0.04 * scale1);

            poly2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(ins_x, ins_y, 0))));

            poly2.Layer = "0";
            poly2.Color = color1;

            poly2.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly2);
            #endregion

            #region poly7
            poly2 = new Polyline();
            poly2.AddVertexAt(0, new Point2d(scale1 * 1.02327804174274, scale1 * 0.307032341137528), 0, 0.04 * scale1, 0.04 * scale1);
            poly2.AddVertexAt(1, new Point2d(scale1 * 1.0293888906017, scale1 * 0.219096973538399), 0, 0.04 * scale1, 0.04 * scale1);

            poly2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(ins_x, ins_y, 0))));

            poly2.Layer = "0";
            poly2.Color = color1;

            poly2.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly2);
            #endregion

            #region poly8
            poly2 = new Polyline();
            poly2.AddVertexAt(0, new Point2d(scale1 * 1.40388278568127, scale1 * 0.145320844298751), 0, 0.04 * scale1, 0.04 * scale1);
            poly2.AddVertexAt(1, new Point2d(scale1 * 1.5, scale1 * 0.243126062189193), 0, 0.04 * scale1, 0.04 * scale1);

            poly2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(ins_x, ins_y, 0))));

            poly2.Layer = "0";
            poly2.Color = color1;

            poly2.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly2);
            #endregion

            #region poly9
            poly2 = new Polyline();
            poly2.AddVertexAt(0, new Point2d(scale1 * 1.46775637533637, scale1 * 0), 0, 0.04 * scale1, 0.04 * scale1);
            poly2.AddVertexAt(1, new Point2d(scale1 * 1.45248363376595, scale1 * 0.089711032807827), 0, 0.04 * scale1, 0.04 * scale1);
            poly2.AddVertexAt(2, new Point2d(scale1 * 1.35595276509412, scale1 * 0.200163085013628), 0, 0.04 * scale1, 0.04 * scale1);
            poly2.AddVertexAt(3, new Point2d(scale1 * 1.28101538098417, scale1 * 0.222697917371988), 0, 0.04 * scale1, 0.04 * scale1);
            poly2.AddVertexAt(4, new Point2d(scale1 * 1.05375721305609, scale1 * 0.233223754912614), 0, 0.04 * scale1, 0.04 * scale1);
            poly2.AddVertexAt(5, new Point2d(scale1 * 1.0293888906017, scale1 * 0.219096973538399), 0, 0.04 * scale1, 0.04 * scale1);
            poly2.AddVertexAt(6, new Point2d(scale1 * 1.0398377305828, scale1 * 0.171563275158405), 0, 0.04 * scale1, 0.04 * scale1);
            poly2.AddVertexAt(7, new Point2d(scale1 * 1.09961416805163, scale1 * 0.079808434471488), 0, 0.04 * scale1, 0.04 * scale1);
            poly2.AddVertexAt(8, new Point2d(scale1 * 1.15941175713026, scale1 * 0), 0, 0.04 * scale1, 0.04 * scale1);

            poly2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(ins_x, ins_y, 0))));

            poly2.Layer = "0";
            poly2.Color = color1;

            poly2.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly2);
            #endregion

            #region poly10
            poly2 = new Polyline();
            poly2.AddVertexAt(0, new Point2d(scale1 * 0.828600724926219, scale1 * 0.307032341137528), 0, 0.04 * scale1, 0.04 * scale1);
            poly2.AddVertexAt(1, new Point2d(scale1 * 0.880299265496429, scale1 * 0.0323729142546654), 0, 0.04 * scale1, 0.04 * scale1);
            poly2.AddVertexAt(2, new Point2d(scale1 * 0.89070469686902, scale1 * 0), 0, 0.04 * scale1, 0.04 * scale1);


            poly2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(ins_x, ins_y, 0))));

            poly2.Layer = "0";
            poly2.Color = color1;

            poly2.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly2);
            #endregion


        }



        private void add_pattern_MUDSTONE(BlockTableRecord bltrec1, double scale1, double graph_vexag, double stick_vexag, Polyline poly1, BlockTableRecord BTrecord, Autodesk.AutoCAD.DatabaseServices.Transaction Trans1)
        {



            Autodesk.AutoCAD.Colors.Color color_mdst = Autodesk.AutoCAD.Colors.Color.FromRgb(127, 95, 93);
            Autodesk.AutoCAD.Colors.Color color1 = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 255, 255);



            double ins_x = poly1.GetPoint2dAt(3).X;
            double ins_y = poly1.GetPoint2dAt(3).Y;


            string nume_hatch = "ANSI31";
            double hatch_scale = 0.23 * scale1;
            double hatch_angle = 0;

            Polyline poly3 = new Polyline();
            poly3 = poly1.Clone() as Polyline;
            BTrecord.AppendEntity(poly3);
            Trans1.AddNewlyCreatedDBObject(poly3, true);

            Hatch hatch1 = CreateHatch(poly3, nume_hatch, hatch_scale, hatch_angle * Math.PI / 180);
            hatch1.Layer = "0";
            hatch1.LineWeight = LineWeight.LineWeight000;
            hatch1.Color = color_mdst;
            bltrec1.AppendEntity(hatch1);
            poly3.Erase();


            double pattern_width = 0.1 * scale1;
            double pattern_height = 0.1 * scale1;

            double spc_h_edge = 0;
            double spc_v_edge = 0;
            double spc_hor = 0;
            double spc_ver = 0;

            int nr_col = 0;
            int nr_rows = 0;

            double x1 = poly1.GetPoint2dAt(3).X + spc_h_edge;
            double y1 = poly1.GetPoint2dAt(3).Y + spc_v_edge;


            double stick_width = poly1.GetPoint2dAt(1).X - poly1.GetPoint2dAt(0).X;
            double stick_height = poly1.GetPoint2dAt(1).Y - poly1.GetPoint2dAt(2).Y;



            if (stick_height <= pattern_height)
            {
                nr_rows = 1;
            }
            else
            {
                double nr1 = Math.Floor((stick_height - 2 * spc_v_edge) / (pattern_height + spc_ver));
                nr_rows = Convert.ToInt32(nr1);
            }

            if (stick_width - 2 * spc_h_edge < pattern_width + 2 * spc_hor)
            {
                nr_col = 1;
            }
            else
            {
                double nr2 = Math.Floor((stick_width - 2 * spc_h_edge) / (pattern_width + spc_hor));

                nr_col = Convert.ToInt32(nr2);
            }

            double dif_len = stick_width - (nr_col * (pattern_width + spc_hor) - spc_hor);
            double dif_hght = stick_height - (nr_rows * (pattern_height + spc_ver) - spc_ver);

            Point3d pt_ins = poly1.GetPoint3dAt(3);

            if (nr_rows > 0 && nr_col > 0)
            {
                for (int m = 0; m < nr_col; ++m)
                {
                    for (int n = 0; n < nr_rows; ++n)
                    {
                        double x2 = x1 + dif_len / 2 + m * (pattern_width + spc_hor);
                        double y2 = y1 + dif_hght / 2 + n * (pattern_height + spc_ver);

                        double x3 = x2 + pattern_width;
                        double y3 = y2;


                        Polyline polymdst1 = get_poly_mdst1(scale1);
                        polymdst1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));
                        polymdst1.Layer = "0";
                        polymdst1.Color = color1;
                        polymdst1.LineWeight = LineWeight.LineWeight000;
                        bltrec1.AppendEntity(polymdst1);

                        Polyline polymdst2 = get_poly_mdst2(scale1);
                        polymdst2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                        polymdst2.Layer = "0";
                        polymdst2.Color = color1;
                        polymdst2.LineWeight = LineWeight.LineWeight000;
                        bltrec1.AppendEntity(polymdst2);

                        Polyline polymdst3 = get_poly_mdst3(scale1);
                        polymdst3.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                        polymdst3.Layer = "0";
                        polymdst3.Color = color1;
                        polymdst3.LineWeight = LineWeight.LineWeight000;
                        bltrec1.AppendEntity(polymdst3);


                    }
                }



            }

            else
            {
            }


        }



        private void add_pattern_GP(BlockTableRecord bltrec1, double scale1, double graph_vexag, double stick_vexag, Polyline poly1, BlockTableRecord BTrecord, Autodesk.AutoCAD.DatabaseServices.Transaction Trans1)
        {



            Autodesk.AutoCAD.Colors.Color color_GP = Autodesk.AutoCAD.Colors.Color.FromRgb(127, 95, 93);



            double pattern_width = 0.1 * scale1;
            double pattern_height = 0.06 * scale1;

            double spc_h_edge = 0;
            double spc_v_edge = 0;
            double spc_hor = 0;
            double spc_ver = 0;

            int nr_col = 0;
            int nr_rows = 1;

            double x1 = poly1.GetPoint2dAt(3).X + spc_h_edge;
            double y1 = poly1.GetPoint2dAt(3).Y + spc_v_edge;


            double stick_width = poly1.GetPoint2dAt(1).X - poly1.GetPoint2dAt(0).X;
            double rectangle_height = poly1.GetPoint2dAt(1).Y - poly1.GetPoint2dAt(2).Y;

            if (rectangle_height >= pattern_height + 2 * spc_v_edge)
            {
                double nr1 = Math.Floor((rectangle_height - 2 * spc_v_edge) / (pattern_height + spc_ver));
                nr_rows = Convert.ToInt32(nr1);

                if (stick_width - 2 * spc_h_edge < pattern_width + 2 * spc_hor)
                {
                    nr_col = 1;
                }
                else
                {
                    double nr2 = Math.Floor((stick_width - 2 * spc_h_edge) / (pattern_width + spc_hor));

                    nr_col = Convert.ToInt32(nr2);
                }

                double dif_len = stick_width - (nr_col * (pattern_width + spc_hor) - spc_hor);
                double dif_hght = rectangle_height - (nr_rows * (pattern_height + spc_ver) - spc_ver);

                Point3d pt_ins = poly1.GetPoint3dAt(3);

                if (nr_rows > 0 && nr_col > 0)
                {
                    for (int m = 0; m < nr_col; ++m)
                    {
                        for (int n = 0; n < nr_rows; ++n)
                        {
                            double x2 = x1 + dif_len / 2 + m * (pattern_width + spc_hor);
                            double y2 = y1 + dif_hght / 2 + n * (pattern_height + spc_ver);

                            double x3 = x2 + pattern_width;
                            double y3 = y2;


                            Polyline polygp1 = get_poly_gp1(scale1);
                            polygp1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));
                            polygp1.Layer = "0";
                            polygp1.Color = color_GP;
                            polygp1.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp1);

                            Polyline polygp2 = get_poly_gp2(scale1);
                            polygp2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp2.Layer = "0";
                            polygp2.Color = color_GP;
                            polygp2.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp2);

                            Polyline polygp3 = get_poly_gp3(scale1);
                            polygp3.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp3.Layer = "0";
                            polygp3.Color = color_GP;
                            polygp3.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp3);

                            Polyline polygp4 = get_poly_gp4(scale1);
                            polygp4.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp4.Layer = "0";
                            polygp4.Color = color_GP;
                            polygp4.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp4);

                            Polyline polygp5 = get_poly_gp5(scale1);
                            polygp5.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp5.Layer = "0";
                            polygp5.Color = color_GP;
                            polygp5.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp5);

                            Polyline polygp6 = get_poly_gp6(scale1);
                            polygp6.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp6.Layer = "0";
                            polygp6.Color = color_GP;
                            polygp6.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp6);

                            Polyline polygp7 = get_poly_gp7(scale1);
                            polygp7.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp7.Layer = "0";
                            polygp7.Color = color_GP;
                            polygp7.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp7);

                            Polyline polygp8 = get_poly_gp8(scale1);
                            polygp8.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp8.Layer = "0";
                            polygp8.Color = color_GP;
                            polygp8.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp8);

                            Polyline polygp9 = get_poly_gp9(scale1);
                            polygp9.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp9.Layer = "0";
                            polygp9.Color = color_GP;
                            polygp9.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp9);

                            Polyline polygp10 = get_poly_gp10(scale1);
                            polygp10.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp10.Layer = "0";
                            polygp10.Color = color_GP;
                            polygp10.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp10);

                        }
                    }



                }
            }
            else
            {
            }


        }



        private void add_pattern_GP_legend(BlockTableRecord bltrec1, double scale1, double graph_vexag, double stick_vexag, Polyline poly1, BlockTableRecord BTrecord, Autodesk.AutoCAD.DatabaseServices.Transaction Trans1)
        {



            Autodesk.AutoCAD.Colors.Color color_GP = Autodesk.AutoCAD.Colors.Color.FromRgb(127, 95, 93);



            double pattern_width = 0.1 * scale1;
            double pattern_height = 0.06 * scale1;

            double spc_h_edge = 0;
            double spc_v_edge = 0;
            double spc_hor = 0;
            double spc_ver = 0;

            int nr_col = 0;
            int nr_rows = 1;

            double x1 = poly1.GetPoint2dAt(3).X + spc_h_edge;
            double y1 = poly1.GetPoint2dAt(3).Y + spc_v_edge;


            double stick_width = poly1.GetPoint2dAt(1).X - poly1.GetPoint2dAt(0).X;
            double rectangle_height = poly1.GetPoint2dAt(1).Y - poly1.GetPoint2dAt(2).Y;

            if (rectangle_height >= pattern_height + 2 * spc_v_edge)
            {
                double nr1 = Math.Floor((rectangle_height - 2 * spc_v_edge) / (pattern_height + spc_ver));
                nr_rows = Convert.ToInt32(nr1);

                if (stick_width - 2 * spc_h_edge < pattern_width + 2 * spc_hor)
                {
                    nr_col = 1;
                }
                else
                {
                    double nr2 = Math.Floor((stick_width - 2 * spc_h_edge) / (pattern_width + spc_hor));

                    nr_col = Convert.ToInt32(nr2);
                }

                double dif_len = stick_width - (nr_col * (pattern_width + spc_hor) - spc_hor);
                double dif_hght = rectangle_height - (nr_rows * (pattern_height + spc_ver) - spc_ver);

                Point3d pt_ins = poly1.GetPoint3dAt(3);

                if (nr_rows > 0 && nr_col > 0)
                {
                    for (int m = 0; m < nr_col; ++m)
                    {
                        for (int n = 0; n < nr_rows; ++n)
                        {
                            double x2 = x1 + dif_len / 2 + m * (pattern_width + spc_hor);
                            double y2 = y1 + dif_hght / 2 + n * (pattern_height + spc_ver);

                            double x3 = x2 + pattern_width;
                            double y3 = y2;


                            Polyline polygp1 = get_poly_gp1(scale1);
                            polygp1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));
                            polygp1.Layer = "0";
                            polygp1.Color = color_GP;
                            polygp1.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp1);

                            Polyline polygp2 = get_poly_gp2(scale1);
                            polygp2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp2.Layer = "0";
                            polygp2.Color = color_GP;
                            polygp2.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp2);

                            Polyline polygp3 = get_poly_gp3(scale1);
                            polygp3.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp3.Layer = "0";
                            polygp3.Color = color_GP;
                            polygp3.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp3);

                            Polyline polygp4 = get_poly_gp4(scale1);
                            polygp4.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp4.Layer = "0";
                            polygp4.Color = color_GP;
                            polygp4.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp4);

                            Polyline polygp5 = get_poly_gp5(scale1);
                            polygp5.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp5.Layer = "0";
                            polygp5.Color = color_GP;
                            polygp5.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp5);

                            Polyline polygp6 = get_poly_gp6(scale1);
                            polygp6.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp6.Layer = "0";
                            polygp6.Color = color_GP;
                            polygp6.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp6);

                            Polyline polygp7 = get_poly_gp7(scale1);
                            polygp7.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp7.Layer = "0";
                            polygp7.Color = color_GP;
                            polygp7.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp7);

                            Polyline polygp8 = get_poly_gp8(scale1);
                            polygp8.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp8.Layer = "0";
                            polygp8.Color = color_GP;
                            polygp8.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp8);

                            Polyline polygp9 = get_poly_gp9(scale1);
                            polygp9.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp9.Layer = "0";
                            polygp9.Color = color_GP;
                            polygp9.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp9);

                            Polyline polygp10 = get_poly_gp10(scale1);
                            polygp10.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp10.Layer = "0";
                            polygp10.Color = color_GP;
                            polygp10.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp10);

                        }
                    }



                }
            }
            else
            {
            }

            #region poly extra for filling the gap
            Polyline polye1 = new Polyline();
            polye1.AddVertexAt(0, new Point2d(scale1 * 1.46028574100171, scale1 * 0.3), 0, 0, 0);
            polye1.AddVertexAt(1, new Point2d(scale1 * 1.46164016907527, scale1 * 0.297968357889651), 0, 0, 0);
            polye1.AddVertexAt(2, new Point2d(scale1 * 1.46467928719744, scale1 * 0.297687823909144), 0, 0, 0);
            polye1.AddVertexAt(3, new Point2d(scale1 * 1.46818596195379, scale1 * 0.29937102779219), 0, 0, 0);
            polye1.AddVertexAt(4, new Point2d(scale1 * 1.46818596195379, scale1 * 0.3), 0, 0, 0);
            polye1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            polye1.Layer = "0";
            polye1.Color = color_GP;
            polye1.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polye1);
            Polyline poly2 = new Polyline();
            poly2.AddVertexAt(0, new Point2d(scale1 * 1.44565464779079, scale1 * 0.3), 0, 0, 0);
            poly2.AddVertexAt(1, new Point2d(scale1 * 1.44457435192772, scale1 * 0.297407289928637), 0, 0, 0);
            poly2.AddVertexAt(2, new Point2d(scale1 * 1.43943122895175, scale1 * 0.295583819055336), 0, 0, 0);
            poly2.AddVertexAt(3, new Point2d(scale1 * 1.42937876131689, scale1 * 0.295303285074829), 0, 0, 0);
            poly2.AddVertexAt(4, new Point2d(scale1 * 1.424469416658, scale1 * 0.296144887016352), 0, 0, 0);
            poly2.AddVertexAt(5, new Point2d(scale1 * 1.42096274190166, scale1 * 0.297828090899398), 0, 0, 0);
            poly2.AddVertexAt(6, new Point2d(scale1 * 1.41875935585757, scale1 * 0.3), 0, 0, 0);
            poly2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly2.Layer = "0";
            poly2.Color = color_GP;
            poly2.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly2);
            Polyline poly3 = new Polyline();
            poly3.AddVertexAt(0, new Point2d(scale1 * 1.45836727263602, scale1 * 0.282994061761134), 0, 0, 0);
            poly3.AddVertexAt(1, new Point2d(scale1 * 1.4564970460993, scale1 * 0.282853794770882), 0, 0, 0);
            poly3.AddVertexAt(2, new Point2d(scale1 * 1.4541592629284, scale1 * 0.28355512972215), 0, 0, 0);
            poly3.AddVertexAt(3, new Point2d(scale1 * 1.45228903639168, scale1 * 0.285518867585705), 0, 0, 0);
            poly3.AddVertexAt(4, new Point2d(scale1 * 1.45392548461131, scale1 * 0.287061804478495), 0, 0, 0);
            poly3.AddVertexAt(5, new Point2d(scale1 * 1.45836727263602, scale1 * 0.286220202536974), 0, 0, 0);
            poly3.AddVertexAt(6, new Point2d(scale1 * 1.45953616422147, scale1 * 0.284817532634435), 0, 0, 0);
            poly3.Closed = true;
            poly3.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly3.Layer = "0";
            poly3.Color = color_GP;
            poly3.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly3);
            Polyline poly4 = new Polyline();
            poly4.AddVertexAt(0, new Point2d(scale1 * 1.4679521836367, scale1 * 0.28355512972215), 0, 0, 0);
            poly4.AddVertexAt(1, new Point2d(scale1 * 1.46608195709998, scale1 * 0.282012192829359), 0, 0, 0);
            poly4.AddVertexAt(2, new Point2d(scale1 * 1.46584817878289, scale1 * 0.280048454965803), 0, 0, 0);
            poly4.AddVertexAt(3, new Point2d(scale1 * 1.46841974027088, scale1 * 0.278505518073012), 0, 0, 0);
            poly4.AddVertexAt(4, new Point2d(scale1 * 1.47169263671013, scale1 * 0.278926319043773), 0, 0, 0);
            poly4.AddVertexAt(5, new Point2d(scale1 * 1.4723939716614, scale1 * 0.281030323897581), 0, 0, 0);
            poly4.AddVertexAt(6, new Point2d(scale1 * 1.47192641502722, scale1 * 0.282713527780627), 0, 0, 0);
            poly4.AddVertexAt(7, new Point2d(scale1 * 1.4679521836367, scale1 * 0.28355512972215), 0, 0, 0);
            poly4.Closed = true;
            poly4.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly4.Layer = "0";
            poly4.Color = color_GP;
            poly4.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly4);
            Polyline poly5 = new Polyline();
            poly5.AddVertexAt(0, new Point2d(scale1 * 1.41657711212705, scale1 * 0.287355090003573), 0, 0, 0);
            poly5.AddVertexAt(1, new Point2d(scale1 * 1.41423932895615, scale1 * 0.285251085149766), 0, 0, 0);
            poly5.AddVertexAt(2, new Point2d(scale1 * 1.41564199885869, scale1 * 0.283427614276465), 0, 0, 0);
            poly5.AddVertexAt(3, new Point2d(scale1 * 1.41961623024921, scale1 * 0.282726279325197), 0, 0, 0);
            poly5.AddVertexAt(4, new Point2d(scale1 * 1.42265534837138, scale1 * 0.283988682237481), 0, 0, 0);
            poly5.AddVertexAt(5, new Point2d(scale1 * 1.42265534837138, scale1 * 0.286653755052303), 0, 0, 0);
            poly5.AddVertexAt(6, new Point2d(scale1 * 1.41657711212705, scale1 * 0.287355090003573), 0, 0, 0);
            poly5.Closed = true;
            poly5.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly5.Layer = "0";
            poly5.Color = color_GP;
            poly5.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly5);
            Polyline poly6 = new Polyline();
            poly6.AddVertexAt(0, new Point2d(scale1 * 1.41108075824353, scale1 * 0.266735842436258), 0, 0, 0);
            poly6.AddVertexAt(1, new Point2d(scale1 * 1.40947696180731, scale1 * 0.263254670769046), 0, 0, 0);
            poly6.AddVertexAt(2, new Point2d(scale1 * 1.4097107401244, scale1 * 0.259607729022446), 0, 0, 0);
            poly6.AddVertexAt(3, new Point2d(scale1 * 1.41064585339276, scale1 * 0.256802389217368), 0, 0, 0);
            poly6.AddVertexAt(4, new Point2d(scale1 * 1.41485386310037, scale1 * 0.255960787275844), 0, 0, 0);
            poly6.AddVertexAt(5, new Point2d(scale1 * 1.41906187280799, scale1 * 0.256521855236861), 0, 0, 0);
            poly6.AddVertexAt(6, new Point2d(scale1 * 1.42256854756433, scale1 * 0.258625860090668), 0, 0, 0);
            poly6.AddVertexAt(7, new Point2d(scale1 * 1.41906187280799, scale1 * 0.263535204749552), 0, 0, 0);
            poly6.AddVertexAt(8, new Point2d(scale1 * 1.41108075824353, scale1 * 0.266735842436258), 0, 0, 0);
            poly6.Closed = true;
            poly6.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly6.Layer = "0";
            poly6.Color = color_GP;
            poly6.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly6);
            Polyline poly7 = new Polyline();
            poly7.AddVertexAt(0, new Point2d(scale1 * 1.5, scale1 * 0.282677952223438), 0, 0, 0);
            poly7.AddVertexAt(1, new Point2d(scale1 * 1.49989813913476, scale1 * 0.282764533958894), 0, 0, 0);
            poly7.AddVertexAt(2, new Point2d(scale1 * 1.49756035596386, scale1 * 0.286271208715241), 0, 0, 0);
            poly7.AddVertexAt(3, new Point2d(scale1 * 1.49569012942714, scale1 * 0.288796014539811), 0, 0, 0);
            poly7.AddVertexAt(4, new Point2d(scale1 * 1.49148211971953, scale1 * 0.289777883471588), 0, 0, 0);
            poly7.AddVertexAt(5, new Point2d(scale1 * 1.48797544496318, scale1 * 0.290759752403364), 0, 0, 0);
            poly7.AddVertexAt(6, new Point2d(scale1 * 1.48470254852392, scale1 * 0.288936281530065), 0, 0, 0);
            poly7.AddVertexAt(7, new Point2d(scale1 * 1.48259854367012, scale1 * 0.283746402890672), 0, 0, 0);
            poly7.AddVertexAt(8, new Point2d(scale1 * 1.48259854367012, scale1 * 0.281221597066103), 0, 0, 0);
            poly7.AddVertexAt(9, new Point2d(scale1 * 1.48189720871885, scale1 * 0.275751184446203), 0, 0, 0);
            poly7.AddVertexAt(10, new Point2d(scale1 * 1.48540388347519, scale1 * 0.273366645611888), 0, 0, 0);
            poly7.AddVertexAt(11, new Point2d(scale1 * 1.49031322813408, scale1 * 0.272525043670365), 0, 0, 0);
            poly7.AddVertexAt(12, new Point2d(scale1 * 1.5, scale1 * 0.272880884269521), 0, 0, 0);
            poly7.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly7.Layer = "0";
            poly7.Color = color_GP;
            poly7.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly7);
            Polyline poly8 = new Polyline();
            poly8.AddVertexAt(0, new Point2d(scale1 * 1.45836727263602, scale1 * 0.274437775355651), 0, 0, 0);
            poly8.AddVertexAt(1, new Point2d(scale1 * 1.46608195709998, scale1 * 0.272754571472604), 0, 0, 0);
            poly8.AddVertexAt(2, new Point2d(scale1 * 1.46935485353923, scale1 * 0.268125760794228), 0, 0, 0);
            poly8.AddVertexAt(3, new Point2d(scale1 * 1.46608195709998, scale1 * 0.263637217106105), 0, 0, 0);
            poly8.AddVertexAt(4, new Point2d(scale1 * 1.45836727263602, scale1 * 0.261954013223059), 0, 0, 0);
            poly8.AddVertexAt(5, new Point2d(scale1 * 1.4497174749037, scale1 * 0.261392945262043), 0, 0, 0);
            poly8.AddVertexAt(6, new Point2d(scale1 * 1.4450419085619, scale1 * 0.26237481419382), 0, 0, 0);
            poly8.AddVertexAt(7, new Point2d(scale1 * 1.44130145548847, scale1 * 0.26489962001839), 0, 0, 0);
            poly8.AddVertexAt(8, new Point2d(scale1 * 1.43078143121943, scale1 * 0.268967362735751), 0, 0, 0);
            poly8.AddVertexAt(9, new Point2d(scale1 * 1.42517075160927, scale1 * 0.272894838462859), 0, 0, 0);
            poly8.AddVertexAt(10, new Point2d(scale1 * 1.42657342151181, scale1 * 0.27471830933616), 0, 0, 0);
            poly8.AddVertexAt(11, new Point2d(scale1 * 1.43498944092704, scale1 * 0.278365251082757), 0, 0, 0);
            poly8.AddVertexAt(12, new Point2d(scale1 * 1.44153523380556, scale1 * 0.279066586034028), 0, 0, 0);
            poly8.AddVertexAt(13, new Point2d(scale1 * 1.44574324351317, scale1 * 0.277804183121743), 0, 0, 0);
            poly8.AddVertexAt(14, new Point2d(scale1 * 1.45135392312332, scale1 * 0.276541780209458), 0, 0, 0);
            poly8.AddVertexAt(15, new Point2d(scale1 * 1.45579571114803, scale1 * 0.275840445258188), 0, 0, 0);
            poly8.Closed = true;
            poly8.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly8.Layer = "0";
            poly8.Color = color_GP;
            poly8.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly8);
            Polyline poly9 = new Polyline();
            poly9.AddVertexAt(0, new Point2d(scale1 * 1.47889889791182, scale1 * 0.261890255500208), 0, 0, 0);
            poly9.AddVertexAt(1, new Point2d(scale1 * 1.47274102660718, scale1 * 0.262017770945895), 0, 0, 0);
            poly9.AddVertexAt(2, new Point2d(scale1 * 1.46743251686181, scale1 * 0.259722492923566), 0, 0, 0);
            poly9.AddVertexAt(3, new Point2d(scale1 * 1.46637081491274, scale1 * 0.256662122227115), 0, 0, 0);
            poly9.AddVertexAt(4, new Point2d(scale1 * 1.47167932465811, scale1 * 0.255769514107321), 0, 0, 0);
            poly9.AddVertexAt(5, new Point2d(scale1 * 1.47826187674237, scale1 * 0.2567896376728), 0, 0, 0);
            poly9.AddVertexAt(6, new Point2d(scale1 * 1.48208400375904, scale1 * 0.259467462032192), 0, 0, 0);
            poly9.AddVertexAt(7, new Point2d(scale1 * 1.47889889791182, scale1 * 0.261890255500208), 0, 0, 0);
            poly9.Closed = true;
            poly9.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly9.Layer = "0";
            poly9.Color = color_GP;
            poly9.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly9);
            Polyline poly10 = new Polyline();
            poly10.AddVertexAt(0, new Point2d(scale1 * 1.49647313832164, scale1 * 0.248628649148944), 0, 0, 0);
            poly10.AddVertexAt(1, new Point2d(scale1 * 1.49670691663873, scale1 * 0.243298503519296), 0, 0, 0);
            poly10.AddVertexAt(2, new Point2d(scale1 * 1.49507046841911, scale1 * 0.242036100607014), 0, 0, 0);
            poly10.AddVertexAt(3, new Point2d(scale1 * 1.49086245871149, scale1 * 0.240633430704475), 0, 0, 0);
            poly10.AddVertexAt(4, new Point2d(scale1 * 1.48688822732096, scale1 * 0.240773697694727), 0, 0, 0);
            poly10.AddVertexAt(5, new Point2d(scale1 * 1.4838491091988, scale1 * 0.242036100607014), 0, 0, 0);
            poly10.AddVertexAt(6, new Point2d(scale1 * 1.48057621275954, scale1 * 0.243298503519296), 0, 0, 0);
            poly10.AddVertexAt(7, new Point2d(scale1 * 1.47753709463737, scale1 * 0.244560906431583), 0, 0, 0);
            poly10.AddVertexAt(8, new Point2d(scale1 * 1.47566686810066, scale1 * 0.246945445265898), 0, 0, 0);
            poly10.AddVertexAt(9, new Point2d(scale1 * 1.47566686810066, scale1 * 0.248628649148944), 0, 0, 0);
            poly10.AddVertexAt(10, new Point2d(scale1 * 1.479173542857, scale1 * 0.252275590895543), 0, 0, 0);
            poly10.AddVertexAt(11, new Point2d(scale1 * 1.48595311405261, scale1 * 0.253818527788335), 0, 0, 0);
            poly10.AddVertexAt(12, new Point2d(scale1 * 1.49273268524821, scale1 * 0.252275590895543), 0, 0, 0);
            poly10.Closed = true;
            poly10.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly10.Layer = "0";
            poly10.Color = color_GP;
            poly10.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly10);
            Polyline poly11 = new Polyline();
            poly11.AddVertexAt(0, new Point2d(scale1 * 1.46210772570945, scale1 * 0.243438770509551), 0, 0, 0);
            poly11.AddVertexAt(1, new Point2d(scale1 * 1.45930238590438, scale1 * 0.241895833616759), 0, 0, 0);
            poly11.AddVertexAt(2, new Point2d(scale1 * 1.46023749917273, scale1 * 0.240072362743458), 0, 0, 0);
            poly11.AddVertexAt(3, new Point2d(scale1 * 1.46164016907527, scale1 * 0.237968357889651), 0, 0, 0);
            poly11.AddVertexAt(4, new Point2d(scale1 * 1.46467928719744, scale1 * 0.237687823909144), 0, 0, 0);
            poly11.AddVertexAt(5, new Point2d(scale1 * 1.46818596195379, scale1 * 0.23937102779219), 0, 0, 0);
            poly11.AddVertexAt(6, new Point2d(scale1 * 1.46818596195379, scale1 * 0.241615299636252), 0, 0, 0);
            poly11.AddVertexAt(7, new Point2d(scale1 * 1.46584817878289, scale1 * 0.244140105460821), 0, 0, 0);
            poly11.AddVertexAt(8, new Point2d(scale1 * 1.46210772570945, scale1 * 0.243438770509551), 0, 0, 0);
            poly11.Closed = true;
            poly11.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly11.Layer = "0";
            poly11.Color = color_GP;
            poly11.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly11);
            Polyline poly12 = new Polyline();
            poly12.AddVertexAt(0, new Point2d(scale1 * 1.41652095387695, scale1 * 0.248628649148944), 0, 0, 0);
            poly12.AddVertexAt(1, new Point2d(scale1 * 1.42306674675547, scale1 * 0.250452120022245), 0, 0, 0);
            poly12.AddVertexAt(2, new Point2d(scale1 * 1.4359245541954, scale1 * 0.249750785070974), 0, 0, 0);
            poly12.AddVertexAt(3, new Point2d(scale1 * 1.44247034707391, scale1 * 0.248628649148944), 0, 0, 0);
            poly12.AddVertexAt(4, new Point2d(scale1 * 1.44363923865936, scale1 * 0.246805178275643), 0, 0, 0);
            poly12.AddVertexAt(5, new Point2d(scale1 * 1.44387301697645, scale1 * 0.244981707402343), 0, 0, 0);
            poly12.AddVertexAt(6, new Point2d(scale1 * 1.44410679529354, scale1 * 0.242737435558282), 0, 0, 0);
            poly12.AddVertexAt(7, new Point2d(scale1 * 1.44597702183026, scale1 * 0.240773697694727), 0, 0, 0);
            poly12.AddVertexAt(8, new Point2d(scale1 * 1.44457435192772, scale1 * 0.237407289928637), 0, 0, 0);
            poly12.AddVertexAt(9, new Point2d(scale1 * 1.43943122895175, scale1 * 0.235583819055336), 0, 0, 0);
            poly12.AddVertexAt(10, new Point2d(scale1 * 1.42937876131689, scale1 * 0.235303285074829), 0, 0, 0);
            poly12.AddVertexAt(11, new Point2d(scale1 * 1.424469416658, scale1 * 0.236144887016352), 0, 0, 0);
            poly12.AddVertexAt(12, new Point2d(scale1 * 1.42096274190166, scale1 * 0.237828090899398), 0, 0, 0);
            poly12.AddVertexAt(13, new Point2d(scale1 * 1.4176898454624, scale1 * 0.241054231675236), 0, 0, 0);
            poly12.AddVertexAt(14, new Point2d(scale1 * 1.41581961892568, scale1 * 0.244420639441328), 0, 0, 0);
            poly12.Closed = true;
            poly12.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly12.Layer = "0";
            poly12.Color = color_GP;
            poly12.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly12);
            Polyline poly13 = new Polyline();
            poly13.AddVertexAt(0, new Point2d(scale1 * 1.36028574100171, scale1 * 0.3), 0, 0, 0);
            poly13.AddVertexAt(1, new Point2d(scale1 * 1.36164016907527, scale1 * 0.297968357889651), 0, 0, 0);
            poly13.AddVertexAt(2, new Point2d(scale1 * 1.36467928719744, scale1 * 0.297687823909144), 0, 0, 0);
            poly13.AddVertexAt(3, new Point2d(scale1 * 1.36818596195379, scale1 * 0.29937102779219), 0, 0, 0);
            poly13.AddVertexAt(4, new Point2d(scale1 * 1.36818596195379, scale1 * 0.3), 0, 0, 0);
            poly13.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly13.Layer = "0";
            poly13.Color = color_GP;
            poly13.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly13);
            Polyline poly14 = new Polyline();
            poly14.AddVertexAt(0, new Point2d(scale1 * 1.34565464779079, scale1 * 0.3), 0, 0, 0);
            poly14.AddVertexAt(1, new Point2d(scale1 * 1.34457435192772, scale1 * 0.297407289928637), 0, 0, 0);
            poly14.AddVertexAt(2, new Point2d(scale1 * 1.33943122895175, scale1 * 0.295583819055336), 0, 0, 0);
            poly14.AddVertexAt(3, new Point2d(scale1 * 1.32937876131689, scale1 * 0.295303285074829), 0, 0, 0);
            poly14.AddVertexAt(4, new Point2d(scale1 * 1.324469416658, scale1 * 0.296144887016352), 0, 0, 0);
            poly14.AddVertexAt(5, new Point2d(scale1 * 1.32096274190166, scale1 * 0.297828090899398), 0, 0, 0);
            poly14.AddVertexAt(6, new Point2d(scale1 * 1.31875935585757, scale1 * 0.3), 0, 0, 0);
            poly14.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly14.Layer = "0";
            poly14.Color = color_GP;
            poly14.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly14);
            Polyline poly15 = new Polyline();
            poly15.AddVertexAt(0, new Point2d(scale1 * 1.35836727263602, scale1 * 0.282994061761134), 0, 0, 0);
            poly15.AddVertexAt(1, new Point2d(scale1 * 1.3564970460993, scale1 * 0.282853794770882), 0, 0, 0);
            poly15.AddVertexAt(2, new Point2d(scale1 * 1.3541592629284, scale1 * 0.28355512972215), 0, 0, 0);
            poly15.AddVertexAt(3, new Point2d(scale1 * 1.35228903639168, scale1 * 0.285518867585705), 0, 0, 0);
            poly15.AddVertexAt(4, new Point2d(scale1 * 1.35392548461131, scale1 * 0.287061804478495), 0, 0, 0);
            poly15.AddVertexAt(5, new Point2d(scale1 * 1.35836727263602, scale1 * 0.286220202536974), 0, 0, 0);
            poly15.AddVertexAt(6, new Point2d(scale1 * 1.35953616422147, scale1 * 0.284817532634435), 0, 0, 0);
            poly15.Closed = true;
            poly15.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly15.Layer = "0";
            poly15.Color = color_GP;
            poly15.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly15);
            Polyline poly16 = new Polyline();
            poly16.AddVertexAt(0, new Point2d(scale1 * 1.3679521836367, scale1 * 0.28355512972215), 0, 0, 0);
            poly16.AddVertexAt(1, new Point2d(scale1 * 1.36608195709998, scale1 * 0.282012192829359), 0, 0, 0);
            poly16.AddVertexAt(2, new Point2d(scale1 * 1.36584817878289, scale1 * 0.280048454965803), 0, 0, 0);
            poly16.AddVertexAt(3, new Point2d(scale1 * 1.36841974027088, scale1 * 0.278505518073012), 0, 0, 0);
            poly16.AddVertexAt(4, new Point2d(scale1 * 1.37169263671013, scale1 * 0.278926319043773), 0, 0, 0);
            poly16.AddVertexAt(5, new Point2d(scale1 * 1.3723939716614, scale1 * 0.281030323897581), 0, 0, 0);
            poly16.AddVertexAt(6, new Point2d(scale1 * 1.37192641502722, scale1 * 0.282713527780627), 0, 0, 0);
            poly16.AddVertexAt(7, new Point2d(scale1 * 1.3679521836367, scale1 * 0.28355512972215), 0, 0, 0);
            poly16.Closed = true;
            poly16.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly16.Layer = "0";
            poly16.Color = color_GP;
            poly16.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly16);
            Polyline poly17 = new Polyline();
            poly17.AddVertexAt(0, new Point2d(scale1 * 1.31657711212705, scale1 * 0.287355090003573), 0, 0, 0);
            poly17.AddVertexAt(1, new Point2d(scale1 * 1.31423932895615, scale1 * 0.285251085149766), 0, 0, 0);
            poly17.AddVertexAt(2, new Point2d(scale1 * 1.31564199885869, scale1 * 0.283427614276465), 0, 0, 0);
            poly17.AddVertexAt(3, new Point2d(scale1 * 1.31961623024921, scale1 * 0.282726279325197), 0, 0, 0);
            poly17.AddVertexAt(4, new Point2d(scale1 * 1.32265534837138, scale1 * 0.283988682237481), 0, 0, 0);
            poly17.AddVertexAt(5, new Point2d(scale1 * 1.32265534837138, scale1 * 0.286653755052303), 0, 0, 0);
            poly17.AddVertexAt(6, new Point2d(scale1 * 1.31657711212705, scale1 * 0.287355090003573), 0, 0, 0);
            poly17.Closed = true;
            poly17.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly17.Layer = "0";
            poly17.Color = color_GP;
            poly17.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly17);
            Polyline poly18 = new Polyline();
            poly18.AddVertexAt(0, new Point2d(scale1 * 1.31108075824353, scale1 * 0.266735842436258), 0, 0, 0);
            poly18.AddVertexAt(1, new Point2d(scale1 * 1.30947696180731, scale1 * 0.263254670769046), 0, 0, 0);
            poly18.AddVertexAt(2, new Point2d(scale1 * 1.3097107401244, scale1 * 0.259607729022446), 0, 0, 0);
            poly18.AddVertexAt(3, new Point2d(scale1 * 1.31064585339276, scale1 * 0.256802389217368), 0, 0, 0);
            poly18.AddVertexAt(4, new Point2d(scale1 * 1.31485386310037, scale1 * 0.255960787275844), 0, 0, 0);
            poly18.AddVertexAt(5, new Point2d(scale1 * 1.31906187280799, scale1 * 0.256521855236861), 0, 0, 0);
            poly18.AddVertexAt(6, new Point2d(scale1 * 1.32256854756433, scale1 * 0.258625860090668), 0, 0, 0);
            poly18.AddVertexAt(7, new Point2d(scale1 * 1.31906187280799, scale1 * 0.263535204749552), 0, 0, 0);
            poly18.AddVertexAt(8, new Point2d(scale1 * 1.31108075824353, scale1 * 0.266735842436258), 0, 0, 0);
            poly18.Closed = true;
            poly18.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly18.Layer = "0";
            poly18.Color = color_GP;
            poly18.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly18);
            Polyline poly19 = new Polyline();
            poly19.AddVertexAt(0, new Point2d(scale1 * 1.38797544496318, scale1 * 0.290759752403364), 0, 0, 0);
            poly19.AddVertexAt(1, new Point2d(scale1 * 1.38470254852392, scale1 * 0.288936281530065), 0, 0, 0);
            poly19.AddVertexAt(2, new Point2d(scale1 * 1.38259854367012, scale1 * 0.283746402890672), 0, 0, 0);
            poly19.AddVertexAt(3, new Point2d(scale1 * 1.38259854367012, scale1 * 0.281221597066103), 0, 0, 0);
            poly19.AddVertexAt(4, new Point2d(scale1 * 1.38189720871885, scale1 * 0.275751184446203), 0, 0, 0);
            poly19.AddVertexAt(5, new Point2d(scale1 * 1.38540388347519, scale1 * 0.273366645611888), 0, 0, 0);
            poly19.AddVertexAt(6, new Point2d(scale1 * 1.39031322813408, scale1 * 0.272525043670365), 0, 0, 0);
            poly19.AddVertexAt(7, new Point2d(scale1 * 1.40176836567148, scale1 * 0.272945844641126), 0, 0, 0);
            poly19.AddVertexAt(8, new Point2d(scale1 * 1.40480748379364, scale1 * 0.274488781533918), 0, 0, 0);
            poly19.AddVertexAt(9, new Point2d(scale1 * 1.405742597062, scale1 * 0.278275990270772), 0, 0, 0);
            poly19.AddVertexAt(10, new Point2d(scale1 * 1.40270347893983, scale1 * 0.28037999512458), 0, 0, 0);
            poly19.AddVertexAt(11, new Point2d(scale1 * 1.39989813913476, scale1 * 0.282764533958894), 0, 0, 0);
            poly19.AddVertexAt(12, new Point2d(scale1 * 1.39756035596386, scale1 * 0.286271208715241), 0, 0, 0);
            poly19.AddVertexAt(13, new Point2d(scale1 * 1.39569012942714, scale1 * 0.288796014539811), 0, 0, 0);
            poly19.AddVertexAt(14, new Point2d(scale1 * 1.39148211971953, scale1 * 0.289777883471588), 0, 0, 0);
            poly19.AddVertexAt(15, new Point2d(scale1 * 1.38797544496318, scale1 * 0.290759752403364), 0, 0, 0);
            poly19.Closed = true;
            poly19.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly19.Layer = "0";
            poly19.Color = color_GP;
            poly19.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly19);
            Polyline poly20 = new Polyline();
            poly20.AddVertexAt(0, new Point2d(scale1 * 1.35836727263602, scale1 * 0.274437775355651), 0, 0, 0);
            poly20.AddVertexAt(1, new Point2d(scale1 * 1.36608195709998, scale1 * 0.272754571472604), 0, 0, 0);
            poly20.AddVertexAt(2, new Point2d(scale1 * 1.36935485353923, scale1 * 0.268125760794228), 0, 0, 0);
            poly20.AddVertexAt(3, new Point2d(scale1 * 1.36608195709998, scale1 * 0.263637217106105), 0, 0, 0);
            poly20.AddVertexAt(4, new Point2d(scale1 * 1.35836727263602, scale1 * 0.261954013223059), 0, 0, 0);
            poly20.AddVertexAt(5, new Point2d(scale1 * 1.3497174749037, scale1 * 0.261392945262043), 0, 0, 0);
            poly20.AddVertexAt(6, new Point2d(scale1 * 1.3450419085619, scale1 * 0.26237481419382), 0, 0, 0);
            poly20.AddVertexAt(7, new Point2d(scale1 * 1.34130145548847, scale1 * 0.26489962001839), 0, 0, 0);
            poly20.AddVertexAt(8, new Point2d(scale1 * 1.33078143121943, scale1 * 0.268967362735751), 0, 0, 0);
            poly20.AddVertexAt(9, new Point2d(scale1 * 1.32517075160927, scale1 * 0.272894838462859), 0, 0, 0);
            poly20.AddVertexAt(10, new Point2d(scale1 * 1.32657342151181, scale1 * 0.27471830933616), 0, 0, 0);
            poly20.AddVertexAt(11, new Point2d(scale1 * 1.33498944092704, scale1 * 0.278365251082757), 0, 0, 0);
            poly20.AddVertexAt(12, new Point2d(scale1 * 1.34153523380556, scale1 * 0.279066586034028), 0, 0, 0);
            poly20.AddVertexAt(13, new Point2d(scale1 * 1.34574324351317, scale1 * 0.277804183121743), 0, 0, 0);
            poly20.AddVertexAt(14, new Point2d(scale1 * 1.35135392312332, scale1 * 0.276541780209458), 0, 0, 0);
            poly20.AddVertexAt(15, new Point2d(scale1 * 1.35579571114803, scale1 * 0.275840445258188), 0, 0, 0);
            poly20.Closed = true;
            poly20.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly20.Layer = "0";
            poly20.Color = color_GP;
            poly20.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly20);
            Polyline poly21 = new Polyline();
            poly21.AddVertexAt(0, new Point2d(scale1 * 1.37889889791182, scale1 * 0.261890255500208), 0, 0, 0);
            poly21.AddVertexAt(1, new Point2d(scale1 * 1.37274102660718, scale1 * 0.262017770945895), 0, 0, 0);
            poly21.AddVertexAt(2, new Point2d(scale1 * 1.36743251686181, scale1 * 0.259722492923566), 0, 0, 0);
            poly21.AddVertexAt(3, new Point2d(scale1 * 1.36637081491274, scale1 * 0.256662122227115), 0, 0, 0);
            poly21.AddVertexAt(4, new Point2d(scale1 * 1.37167932465811, scale1 * 0.255769514107321), 0, 0, 0);
            poly21.AddVertexAt(5, new Point2d(scale1 * 1.37826187674237, scale1 * 0.2567896376728), 0, 0, 0);
            poly21.AddVertexAt(6, new Point2d(scale1 * 1.38208400375904, scale1 * 0.259467462032192), 0, 0, 0);
            poly21.AddVertexAt(7, new Point2d(scale1 * 1.37889889791182, scale1 * 0.261890255500208), 0, 0, 0);
            poly21.Closed = true;
            poly21.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly21.Layer = "0";
            poly21.Color = color_GP;
            poly21.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly21);
            Polyline poly22 = new Polyline();
            poly22.AddVertexAt(0, new Point2d(scale1 * 1.39647313832164, scale1 * 0.248628649148944), 0, 0, 0);
            poly22.AddVertexAt(1, new Point2d(scale1 * 1.39670691663873, scale1 * 0.243298503519296), 0, 0, 0);
            poly22.AddVertexAt(2, new Point2d(scale1 * 1.39507046841911, scale1 * 0.242036100607014), 0, 0, 0);
            poly22.AddVertexAt(3, new Point2d(scale1 * 1.39086245871149, scale1 * 0.240633430704475), 0, 0, 0);
            poly22.AddVertexAt(4, new Point2d(scale1 * 1.38688822732096, scale1 * 0.240773697694727), 0, 0, 0);
            poly22.AddVertexAt(5, new Point2d(scale1 * 1.3838491091988, scale1 * 0.242036100607014), 0, 0, 0);
            poly22.AddVertexAt(6, new Point2d(scale1 * 1.38057621275954, scale1 * 0.243298503519296), 0, 0, 0);
            poly22.AddVertexAt(7, new Point2d(scale1 * 1.37753709463737, scale1 * 0.244560906431583), 0, 0, 0);
            poly22.AddVertexAt(8, new Point2d(scale1 * 1.37566686810066, scale1 * 0.246945445265898), 0, 0, 0);
            poly22.AddVertexAt(9, new Point2d(scale1 * 1.37566686810066, scale1 * 0.248628649148944), 0, 0, 0);
            poly22.AddVertexAt(10, new Point2d(scale1 * 1.379173542857, scale1 * 0.252275590895543), 0, 0, 0);
            poly22.AddVertexAt(11, new Point2d(scale1 * 1.38595311405261, scale1 * 0.253818527788335), 0, 0, 0);
            poly22.AddVertexAt(12, new Point2d(scale1 * 1.39273268524821, scale1 * 0.252275590895543), 0, 0, 0);
            poly22.Closed = true;
            poly22.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly22.Layer = "0";
            poly22.Color = color_GP;
            poly22.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly22);
            Polyline poly23 = new Polyline();
            poly23.AddVertexAt(0, new Point2d(scale1 * 1.36210772570945, scale1 * 0.243438770509551), 0, 0, 0);
            poly23.AddVertexAt(1, new Point2d(scale1 * 1.35930238590438, scale1 * 0.241895833616759), 0, 0, 0);
            poly23.AddVertexAt(2, new Point2d(scale1 * 1.36023749917273, scale1 * 0.240072362743458), 0, 0, 0);
            poly23.AddVertexAt(3, new Point2d(scale1 * 1.36164016907527, scale1 * 0.237968357889651), 0, 0, 0);
            poly23.AddVertexAt(4, new Point2d(scale1 * 1.36467928719744, scale1 * 0.237687823909144), 0, 0, 0);
            poly23.AddVertexAt(5, new Point2d(scale1 * 1.36818596195379, scale1 * 0.23937102779219), 0, 0, 0);
            poly23.AddVertexAt(6, new Point2d(scale1 * 1.36818596195379, scale1 * 0.241615299636252), 0, 0, 0);
            poly23.AddVertexAt(7, new Point2d(scale1 * 1.36584817878289, scale1 * 0.244140105460821), 0, 0, 0);
            poly23.AddVertexAt(8, new Point2d(scale1 * 1.36210772570945, scale1 * 0.243438770509551), 0, 0, 0);
            poly23.Closed = true;
            poly23.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly23.Layer = "0";
            poly23.Color = color_GP;
            poly23.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly23);
            Polyline poly24 = new Polyline();
            poly24.AddVertexAt(0, new Point2d(scale1 * 1.31652095387695, scale1 * 0.248628649148944), 0, 0, 0);
            poly24.AddVertexAt(1, new Point2d(scale1 * 1.32306674675547, scale1 * 0.250452120022245), 0, 0, 0);
            poly24.AddVertexAt(2, new Point2d(scale1 * 1.3359245541954, scale1 * 0.249750785070974), 0, 0, 0);
            poly24.AddVertexAt(3, new Point2d(scale1 * 1.34247034707391, scale1 * 0.248628649148944), 0, 0, 0);
            poly24.AddVertexAt(4, new Point2d(scale1 * 1.34363923865936, scale1 * 0.246805178275643), 0, 0, 0);
            poly24.AddVertexAt(5, new Point2d(scale1 * 1.34387301697645, scale1 * 0.244981707402343), 0, 0, 0);
            poly24.AddVertexAt(6, new Point2d(scale1 * 1.34410679529354, scale1 * 0.242737435558282), 0, 0, 0);
            poly24.AddVertexAt(7, new Point2d(scale1 * 1.34597702183026, scale1 * 0.240773697694727), 0, 0, 0);
            poly24.AddVertexAt(8, new Point2d(scale1 * 1.34457435192772, scale1 * 0.237407289928637), 0, 0, 0);
            poly24.AddVertexAt(9, new Point2d(scale1 * 1.33943122895175, scale1 * 0.235583819055336), 0, 0, 0);
            poly24.AddVertexAt(10, new Point2d(scale1 * 1.32937876131689, scale1 * 0.235303285074829), 0, 0, 0);
            poly24.AddVertexAt(11, new Point2d(scale1 * 1.324469416658, scale1 * 0.236144887016352), 0, 0, 0);
            poly24.AddVertexAt(12, new Point2d(scale1 * 1.32096274190166, scale1 * 0.237828090899398), 0, 0, 0);
            poly24.AddVertexAt(13, new Point2d(scale1 * 1.3176898454624, scale1 * 0.241054231675236), 0, 0, 0);
            poly24.AddVertexAt(14, new Point2d(scale1 * 1.31581961892568, scale1 * 0.244420639441328), 0, 0, 0);
            poly24.Closed = true;
            poly24.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly24.Layer = "0";
            poly24.Color = color_GP;
            poly24.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly24);
            Polyline poly25 = new Polyline();
            poly25.AddVertexAt(0, new Point2d(scale1 * 1.26028574100171, scale1 * 0.3), 0, 0, 0);
            poly25.AddVertexAt(1, new Point2d(scale1 * 1.26164016907527, scale1 * 0.297968357889651), 0, 0, 0);
            poly25.AddVertexAt(2, new Point2d(scale1 * 1.26467928719744, scale1 * 0.297687823909144), 0, 0, 0);
            poly25.AddVertexAt(3, new Point2d(scale1 * 1.26818596195379, scale1 * 0.29937102779219), 0, 0, 0);
            poly25.AddVertexAt(4, new Point2d(scale1 * 1.26818596195379, scale1 * 0.3), 0, 0, 0);
            poly25.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly25.Layer = "0";
            poly25.Color = color_GP;
            poly25.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly25);
            Polyline poly26 = new Polyline();
            poly26.AddVertexAt(0, new Point2d(scale1 * 1.24565464779079, scale1 * 0.3), 0, 0, 0);
            poly26.AddVertexAt(1, new Point2d(scale1 * 1.24457435192772, scale1 * 0.297407289928637), 0, 0, 0);
            poly26.AddVertexAt(2, new Point2d(scale1 * 1.23943122895175, scale1 * 0.295583819055336), 0, 0, 0);
            poly26.AddVertexAt(3, new Point2d(scale1 * 1.22937876131689, scale1 * 0.295303285074829), 0, 0, 0);
            poly26.AddVertexAt(4, new Point2d(scale1 * 1.224469416658, scale1 * 0.296144887016352), 0, 0, 0);
            poly26.AddVertexAt(5, new Point2d(scale1 * 1.22096274190166, scale1 * 0.297828090899398), 0, 0, 0);
            poly26.AddVertexAt(6, new Point2d(scale1 * 1.21875935585757, scale1 * 0.3), 0, 0, 0);
            poly26.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly26.Layer = "0";
            poly26.Color = color_GP;
            poly26.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly26);
            Polyline poly27 = new Polyline();
            poly27.AddVertexAt(0, new Point2d(scale1 * 1.25836727263602, scale1 * 0.282994061761134), 0, 0, 0);
            poly27.AddVertexAt(1, new Point2d(scale1 * 1.2564970460993, scale1 * 0.282853794770882), 0, 0, 0);
            poly27.AddVertexAt(2, new Point2d(scale1 * 1.2541592629284, scale1 * 0.28355512972215), 0, 0, 0);
            poly27.AddVertexAt(3, new Point2d(scale1 * 1.25228903639168, scale1 * 0.285518867585705), 0, 0, 0);
            poly27.AddVertexAt(4, new Point2d(scale1 * 1.25392548461131, scale1 * 0.287061804478495), 0, 0, 0);
            poly27.AddVertexAt(5, new Point2d(scale1 * 1.25836727263602, scale1 * 0.286220202536974), 0, 0, 0);
            poly27.AddVertexAt(6, new Point2d(scale1 * 1.25953616422147, scale1 * 0.284817532634435), 0, 0, 0);
            poly27.Closed = true;
            poly27.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly27.Layer = "0";
            poly27.Color = color_GP;
            poly27.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly27);
            Polyline poly28 = new Polyline();
            poly28.AddVertexAt(0, new Point2d(scale1 * 1.2679521836367, scale1 * 0.28355512972215), 0, 0, 0);
            poly28.AddVertexAt(1, new Point2d(scale1 * 1.26608195709998, scale1 * 0.282012192829359), 0, 0, 0);
            poly28.AddVertexAt(2, new Point2d(scale1 * 1.26584817878289, scale1 * 0.280048454965803), 0, 0, 0);
            poly28.AddVertexAt(3, new Point2d(scale1 * 1.26841974027088, scale1 * 0.278505518073012), 0, 0, 0);
            poly28.AddVertexAt(4, new Point2d(scale1 * 1.27169263671013, scale1 * 0.278926319043773), 0, 0, 0);
            poly28.AddVertexAt(5, new Point2d(scale1 * 1.2723939716614, scale1 * 0.281030323897581), 0, 0, 0);
            poly28.AddVertexAt(6, new Point2d(scale1 * 1.27192641502722, scale1 * 0.282713527780627), 0, 0, 0);
            poly28.AddVertexAt(7, new Point2d(scale1 * 1.2679521836367, scale1 * 0.28355512972215), 0, 0, 0);
            poly28.Closed = true;
            poly28.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly28.Layer = "0";
            poly28.Color = color_GP;
            poly28.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly28);
            Polyline poly29 = new Polyline();
            poly29.AddVertexAt(0, new Point2d(scale1 * 1.21657711212705, scale1 * 0.287355090003573), 0, 0, 0);
            poly29.AddVertexAt(1, new Point2d(scale1 * 1.21423932895615, scale1 * 0.285251085149766), 0, 0, 0);
            poly29.AddVertexAt(2, new Point2d(scale1 * 1.21564199885869, scale1 * 0.283427614276465), 0, 0, 0);
            poly29.AddVertexAt(3, new Point2d(scale1 * 1.21961623024921, scale1 * 0.282726279325197), 0, 0, 0);
            poly29.AddVertexAt(4, new Point2d(scale1 * 1.22265534837138, scale1 * 0.283988682237481), 0, 0, 0);
            poly29.AddVertexAt(5, new Point2d(scale1 * 1.22265534837138, scale1 * 0.286653755052303), 0, 0, 0);
            poly29.AddVertexAt(6, new Point2d(scale1 * 1.21657711212705, scale1 * 0.287355090003573), 0, 0, 0);
            poly29.Closed = true;
            poly29.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly29.Layer = "0";
            poly29.Color = color_GP;
            poly29.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly29);
            Polyline poly30 = new Polyline();
            poly30.AddVertexAt(0, new Point2d(scale1 * 1.21108075824353, scale1 * 0.266735842436258), 0, 0, 0);
            poly30.AddVertexAt(1, new Point2d(scale1 * 1.20947696180731, scale1 * 0.263254670769046), 0, 0, 0);
            poly30.AddVertexAt(2, new Point2d(scale1 * 1.2097107401244, scale1 * 0.259607729022446), 0, 0, 0);
            poly30.AddVertexAt(3, new Point2d(scale1 * 1.21064585339276, scale1 * 0.256802389217368), 0, 0, 0);
            poly30.AddVertexAt(4, new Point2d(scale1 * 1.21485386310037, scale1 * 0.255960787275844), 0, 0, 0);
            poly30.AddVertexAt(5, new Point2d(scale1 * 1.21906187280799, scale1 * 0.256521855236861), 0, 0, 0);
            poly30.AddVertexAt(6, new Point2d(scale1 * 1.22256854756433, scale1 * 0.258625860090668), 0, 0, 0);
            poly30.AddVertexAt(7, new Point2d(scale1 * 1.21906187280799, scale1 * 0.263535204749552), 0, 0, 0);
            poly30.AddVertexAt(8, new Point2d(scale1 * 1.21108075824353, scale1 * 0.266735842436258), 0, 0, 0);
            poly30.Closed = true;
            poly30.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly30.Layer = "0";
            poly30.Color = color_GP;
            poly30.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly30);
            Polyline poly31 = new Polyline();
            poly31.AddVertexAt(0, new Point2d(scale1 * 1.28797544496318, scale1 * 0.290759752403364), 0, 0, 0);
            poly31.AddVertexAt(1, new Point2d(scale1 * 1.28470254852392, scale1 * 0.288936281530065), 0, 0, 0);
            poly31.AddVertexAt(2, new Point2d(scale1 * 1.28259854367012, scale1 * 0.283746402890672), 0, 0, 0);
            poly31.AddVertexAt(3, new Point2d(scale1 * 1.28259854367012, scale1 * 0.281221597066103), 0, 0, 0);
            poly31.AddVertexAt(4, new Point2d(scale1 * 1.28189720871885, scale1 * 0.275751184446203), 0, 0, 0);
            poly31.AddVertexAt(5, new Point2d(scale1 * 1.28540388347519, scale1 * 0.273366645611888), 0, 0, 0);
            poly31.AddVertexAt(6, new Point2d(scale1 * 1.29031322813408, scale1 * 0.272525043670365), 0, 0, 0);
            poly31.AddVertexAt(7, new Point2d(scale1 * 1.30176836567148, scale1 * 0.272945844641126), 0, 0, 0);
            poly31.AddVertexAt(8, new Point2d(scale1 * 1.30480748379364, scale1 * 0.274488781533918), 0, 0, 0);
            poly31.AddVertexAt(9, new Point2d(scale1 * 1.305742597062, scale1 * 0.278275990270772), 0, 0, 0);
            poly31.AddVertexAt(10, new Point2d(scale1 * 1.30270347893983, scale1 * 0.28037999512458), 0, 0, 0);
            poly31.AddVertexAt(11, new Point2d(scale1 * 1.29989813913476, scale1 * 0.282764533958894), 0, 0, 0);
            poly31.AddVertexAt(12, new Point2d(scale1 * 1.29756035596386, scale1 * 0.286271208715241), 0, 0, 0);
            poly31.AddVertexAt(13, new Point2d(scale1 * 1.29569012942714, scale1 * 0.288796014539811), 0, 0, 0);
            poly31.AddVertexAt(14, new Point2d(scale1 * 1.29148211971953, scale1 * 0.289777883471588), 0, 0, 0);
            poly31.AddVertexAt(15, new Point2d(scale1 * 1.28797544496318, scale1 * 0.290759752403364), 0, 0, 0);
            poly31.Closed = true;
            poly31.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly31.Layer = "0";
            poly31.Color = color_GP;
            poly31.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly31);
            Polyline poly32 = new Polyline();
            poly32.AddVertexAt(0, new Point2d(scale1 * 1.25836727263602, scale1 * 0.274437775355651), 0, 0, 0);
            poly32.AddVertexAt(1, new Point2d(scale1 * 1.26608195709998, scale1 * 0.272754571472604), 0, 0, 0);
            poly32.AddVertexAt(2, new Point2d(scale1 * 1.26935485353923, scale1 * 0.268125760794228), 0, 0, 0);
            poly32.AddVertexAt(3, new Point2d(scale1 * 1.26608195709998, scale1 * 0.263637217106105), 0, 0, 0);
            poly32.AddVertexAt(4, new Point2d(scale1 * 1.25836727263602, scale1 * 0.261954013223059), 0, 0, 0);
            poly32.AddVertexAt(5, new Point2d(scale1 * 1.2497174749037, scale1 * 0.261392945262043), 0, 0, 0);
            poly32.AddVertexAt(6, new Point2d(scale1 * 1.2450419085619, scale1 * 0.26237481419382), 0, 0, 0);
            poly32.AddVertexAt(7, new Point2d(scale1 * 1.24130145548847, scale1 * 0.26489962001839), 0, 0, 0);
            poly32.AddVertexAt(8, new Point2d(scale1 * 1.23078143121943, scale1 * 0.268967362735751), 0, 0, 0);
            poly32.AddVertexAt(9, new Point2d(scale1 * 1.22517075160927, scale1 * 0.272894838462859), 0, 0, 0);
            poly32.AddVertexAt(10, new Point2d(scale1 * 1.22657342151181, scale1 * 0.27471830933616), 0, 0, 0);
            poly32.AddVertexAt(11, new Point2d(scale1 * 1.23498944092704, scale1 * 0.278365251082757), 0, 0, 0);
            poly32.AddVertexAt(12, new Point2d(scale1 * 1.24153523380556, scale1 * 0.279066586034028), 0, 0, 0);
            poly32.AddVertexAt(13, new Point2d(scale1 * 1.24574324351317, scale1 * 0.277804183121743), 0, 0, 0);
            poly32.AddVertexAt(14, new Point2d(scale1 * 1.25135392312332, scale1 * 0.276541780209458), 0, 0, 0);
            poly32.AddVertexAt(15, new Point2d(scale1 * 1.25579571114803, scale1 * 0.275840445258188), 0, 0, 0);
            poly32.Closed = true;
            poly32.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly32.Layer = "0";
            poly32.Color = color_GP;
            poly32.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly32);
            Polyline poly33 = new Polyline();
            poly33.AddVertexAt(0, new Point2d(scale1 * 1.27889889791182, scale1 * 0.261890255500208), 0, 0, 0);
            poly33.AddVertexAt(1, new Point2d(scale1 * 1.27274102660718, scale1 * 0.262017770945895), 0, 0, 0);
            poly33.AddVertexAt(2, new Point2d(scale1 * 1.26743251686181, scale1 * 0.259722492923566), 0, 0, 0);
            poly33.AddVertexAt(3, new Point2d(scale1 * 1.26637081491274, scale1 * 0.256662122227115), 0, 0, 0);
            poly33.AddVertexAt(4, new Point2d(scale1 * 1.27167932465811, scale1 * 0.255769514107321), 0, 0, 0);
            poly33.AddVertexAt(5, new Point2d(scale1 * 1.27826187674237, scale1 * 0.2567896376728), 0, 0, 0);
            poly33.AddVertexAt(6, new Point2d(scale1 * 1.28208400375904, scale1 * 0.259467462032192), 0, 0, 0);
            poly33.AddVertexAt(7, new Point2d(scale1 * 1.27889889791182, scale1 * 0.261890255500208), 0, 0, 0);
            poly33.Closed = true;
            poly33.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly33.Layer = "0";
            poly33.Color = color_GP;
            poly33.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly33);
            Polyline poly34 = new Polyline();
            poly34.AddVertexAt(0, new Point2d(scale1 * 1.29647313832164, scale1 * 0.248628649148944), 0, 0, 0);
            poly34.AddVertexAt(1, new Point2d(scale1 * 1.29670691663873, scale1 * 0.243298503519296), 0, 0, 0);
            poly34.AddVertexAt(2, new Point2d(scale1 * 1.29507046841911, scale1 * 0.242036100607014), 0, 0, 0);
            poly34.AddVertexAt(3, new Point2d(scale1 * 1.29086245871149, scale1 * 0.240633430704475), 0, 0, 0);
            poly34.AddVertexAt(4, new Point2d(scale1 * 1.28688822732096, scale1 * 0.240773697694727), 0, 0, 0);
            poly34.AddVertexAt(5, new Point2d(scale1 * 1.2838491091988, scale1 * 0.242036100607014), 0, 0, 0);
            poly34.AddVertexAt(6, new Point2d(scale1 * 1.28057621275954, scale1 * 0.243298503519296), 0, 0, 0);
            poly34.AddVertexAt(7, new Point2d(scale1 * 1.27753709463737, scale1 * 0.244560906431583), 0, 0, 0);
            poly34.AddVertexAt(8, new Point2d(scale1 * 1.27566686810066, scale1 * 0.246945445265898), 0, 0, 0);
            poly34.AddVertexAt(9, new Point2d(scale1 * 1.27566686810066, scale1 * 0.248628649148944), 0, 0, 0);
            poly34.AddVertexAt(10, new Point2d(scale1 * 1.279173542857, scale1 * 0.252275590895543), 0, 0, 0);
            poly34.AddVertexAt(11, new Point2d(scale1 * 1.28595311405261, scale1 * 0.253818527788335), 0, 0, 0);
            poly34.AddVertexAt(12, new Point2d(scale1 * 1.29273268524821, scale1 * 0.252275590895543), 0, 0, 0);
            poly34.Closed = true;
            poly34.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly34.Layer = "0";
            poly34.Color = color_GP;
            poly34.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly34);
            Polyline poly35 = new Polyline();
            poly35.AddVertexAt(0, new Point2d(scale1 * 1.26210772570945, scale1 * 0.243438770509551), 0, 0, 0);
            poly35.AddVertexAt(1, new Point2d(scale1 * 1.25930238590438, scale1 * 0.241895833616759), 0, 0, 0);
            poly35.AddVertexAt(2, new Point2d(scale1 * 1.26023749917273, scale1 * 0.240072362743458), 0, 0, 0);
            poly35.AddVertexAt(3, new Point2d(scale1 * 1.26164016907527, scale1 * 0.237968357889651), 0, 0, 0);
            poly35.AddVertexAt(4, new Point2d(scale1 * 1.26467928719744, scale1 * 0.237687823909144), 0, 0, 0);
            poly35.AddVertexAt(5, new Point2d(scale1 * 1.26818596195379, scale1 * 0.23937102779219), 0, 0, 0);
            poly35.AddVertexAt(6, new Point2d(scale1 * 1.26818596195379, scale1 * 0.241615299636252), 0, 0, 0);
            poly35.AddVertexAt(7, new Point2d(scale1 * 1.26584817878289, scale1 * 0.244140105460821), 0, 0, 0);
            poly35.AddVertexAt(8, new Point2d(scale1 * 1.26210772570945, scale1 * 0.243438770509551), 0, 0, 0);
            poly35.Closed = true;
            poly35.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly35.Layer = "0";
            poly35.Color = color_GP;
            poly35.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly35);
            Polyline poly36 = new Polyline();
            poly36.AddVertexAt(0, new Point2d(scale1 * 1.21652095387695, scale1 * 0.248628649148944), 0, 0, 0);
            poly36.AddVertexAt(1, new Point2d(scale1 * 1.22306674675547, scale1 * 0.250452120022245), 0, 0, 0);
            poly36.AddVertexAt(2, new Point2d(scale1 * 1.2359245541954, scale1 * 0.249750785070974), 0, 0, 0);
            poly36.AddVertexAt(3, new Point2d(scale1 * 1.24247034707391, scale1 * 0.248628649148944), 0, 0, 0);
            poly36.AddVertexAt(4, new Point2d(scale1 * 1.24363923865936, scale1 * 0.246805178275643), 0, 0, 0);
            poly36.AddVertexAt(5, new Point2d(scale1 * 1.24387301697645, scale1 * 0.244981707402343), 0, 0, 0);
            poly36.AddVertexAt(6, new Point2d(scale1 * 1.24410679529354, scale1 * 0.242737435558282), 0, 0, 0);
            poly36.AddVertexAt(7, new Point2d(scale1 * 1.24597702183026, scale1 * 0.240773697694727), 0, 0, 0);
            poly36.AddVertexAt(8, new Point2d(scale1 * 1.24457435192772, scale1 * 0.237407289928637), 0, 0, 0);
            poly36.AddVertexAt(9, new Point2d(scale1 * 1.23943122895175, scale1 * 0.235583819055336), 0, 0, 0);
            poly36.AddVertexAt(10, new Point2d(scale1 * 1.22937876131689, scale1 * 0.235303285074829), 0, 0, 0);
            poly36.AddVertexAt(11, new Point2d(scale1 * 1.224469416658, scale1 * 0.236144887016352), 0, 0, 0);
            poly36.AddVertexAt(12, new Point2d(scale1 * 1.22096274190166, scale1 * 0.237828090899398), 0, 0, 0);
            poly36.AddVertexAt(13, new Point2d(scale1 * 1.2176898454624, scale1 * 0.241054231675236), 0, 0, 0);
            poly36.AddVertexAt(14, new Point2d(scale1 * 1.21581961892568, scale1 * 0.244420639441328), 0, 0, 0);
            poly36.Closed = true;
            poly36.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly36.Layer = "0";
            poly36.Color = color_GP;
            poly36.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly36);
            Polyline poly37 = new Polyline();
            poly37.AddVertexAt(0, new Point2d(scale1 * 1.16028574100171, scale1 * 0.3), 0, 0, 0);
            poly37.AddVertexAt(1, new Point2d(scale1 * 1.16164016907527, scale1 * 0.297968357889651), 0, 0, 0);
            poly37.AddVertexAt(2, new Point2d(scale1 * 1.16467928719744, scale1 * 0.297687823909144), 0, 0, 0);
            poly37.AddVertexAt(3, new Point2d(scale1 * 1.16818596195379, scale1 * 0.29937102779219), 0, 0, 0);
            poly37.AddVertexAt(4, new Point2d(scale1 * 1.16818596195379, scale1 * 0.3), 0, 0, 0);
            poly37.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly37.Layer = "0";
            poly37.Color = color_GP;
            poly37.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly37);
            Polyline poly38 = new Polyline();
            poly38.AddVertexAt(0, new Point2d(scale1 * 1.14565464779079, scale1 * 0.3), 0, 0, 0);
            poly38.AddVertexAt(1, new Point2d(scale1 * 1.14457435192772, scale1 * 0.297407289928637), 0, 0, 0);
            poly38.AddVertexAt(2, new Point2d(scale1 * 1.13943122895175, scale1 * 0.295583819055336), 0, 0, 0);
            poly38.AddVertexAt(3, new Point2d(scale1 * 1.12937876131689, scale1 * 0.295303285074829), 0, 0, 0);
            poly38.AddVertexAt(4, new Point2d(scale1 * 1.124469416658, scale1 * 0.296144887016352), 0, 0, 0);
            poly38.AddVertexAt(5, new Point2d(scale1 * 1.12096274190166, scale1 * 0.297828090899398), 0, 0, 0);
            poly38.AddVertexAt(6, new Point2d(scale1 * 1.11875935585757, scale1 * 0.3), 0, 0, 0);
            poly38.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly38.Layer = "0";
            poly38.Color = color_GP;
            poly38.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly38);
            Polyline poly39 = new Polyline();
            poly39.AddVertexAt(0, new Point2d(scale1 * 1.15836727263602, scale1 * 0.282994061761134), 0, 0, 0);
            poly39.AddVertexAt(1, new Point2d(scale1 * 1.1564970460993, scale1 * 0.282853794770882), 0, 0, 0);
            poly39.AddVertexAt(2, new Point2d(scale1 * 1.1541592629284, scale1 * 0.28355512972215), 0, 0, 0);
            poly39.AddVertexAt(3, new Point2d(scale1 * 1.15228903639168, scale1 * 0.285518867585705), 0, 0, 0);
            poly39.AddVertexAt(4, new Point2d(scale1 * 1.15392548461131, scale1 * 0.287061804478495), 0, 0, 0);
            poly39.AddVertexAt(5, new Point2d(scale1 * 1.15836727263602, scale1 * 0.286220202536974), 0, 0, 0);
            poly39.AddVertexAt(6, new Point2d(scale1 * 1.15953616422147, scale1 * 0.284817532634435), 0, 0, 0);
            poly39.Closed = true;
            poly39.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly39.Layer = "0";
            poly39.Color = color_GP;
            poly39.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly39);
            Polyline poly40 = new Polyline();
            poly40.AddVertexAt(0, new Point2d(scale1 * 1.1679521836367, scale1 * 0.28355512972215), 0, 0, 0);
            poly40.AddVertexAt(1, new Point2d(scale1 * 1.16608195709998, scale1 * 0.282012192829359), 0, 0, 0);
            poly40.AddVertexAt(2, new Point2d(scale1 * 1.16584817878289, scale1 * 0.280048454965803), 0, 0, 0);
            poly40.AddVertexAt(3, new Point2d(scale1 * 1.16841974027088, scale1 * 0.278505518073012), 0, 0, 0);
            poly40.AddVertexAt(4, new Point2d(scale1 * 1.17169263671013, scale1 * 0.278926319043773), 0, 0, 0);
            poly40.AddVertexAt(5, new Point2d(scale1 * 1.1723939716614, scale1 * 0.281030323897581), 0, 0, 0);
            poly40.AddVertexAt(6, new Point2d(scale1 * 1.17192641502722, scale1 * 0.282713527780627), 0, 0, 0);
            poly40.AddVertexAt(7, new Point2d(scale1 * 1.1679521836367, scale1 * 0.28355512972215), 0, 0, 0);
            poly40.Closed = true;
            poly40.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly40.Layer = "0";
            poly40.Color = color_GP;
            poly40.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly40);
            Polyline poly41 = new Polyline();
            poly41.AddVertexAt(0, new Point2d(scale1 * 1.11657711212705, scale1 * 0.287355090003573), 0, 0, 0);
            poly41.AddVertexAt(1, new Point2d(scale1 * 1.11423932895615, scale1 * 0.285251085149766), 0, 0, 0);
            poly41.AddVertexAt(2, new Point2d(scale1 * 1.11564199885869, scale1 * 0.283427614276465), 0, 0, 0);
            poly41.AddVertexAt(3, new Point2d(scale1 * 1.11961623024921, scale1 * 0.282726279325197), 0, 0, 0);
            poly41.AddVertexAt(4, new Point2d(scale1 * 1.12265534837138, scale1 * 0.283988682237481), 0, 0, 0);
            poly41.AddVertexAt(5, new Point2d(scale1 * 1.12265534837138, scale1 * 0.286653755052303), 0, 0, 0);
            poly41.AddVertexAt(6, new Point2d(scale1 * 1.11657711212705, scale1 * 0.287355090003573), 0, 0, 0);
            poly41.Closed = true;
            poly41.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly41.Layer = "0";
            poly41.Color = color_GP;
            poly41.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly41);
            Polyline poly42 = new Polyline();
            poly42.AddVertexAt(0, new Point2d(scale1 * 1.11108075824353, scale1 * 0.266735842436258), 0, 0, 0);
            poly42.AddVertexAt(1, new Point2d(scale1 * 1.10947696180731, scale1 * 0.263254670769046), 0, 0, 0);
            poly42.AddVertexAt(2, new Point2d(scale1 * 1.1097107401244, scale1 * 0.259607729022446), 0, 0, 0);
            poly42.AddVertexAt(3, new Point2d(scale1 * 1.11064585339276, scale1 * 0.256802389217368), 0, 0, 0);
            poly42.AddVertexAt(4, new Point2d(scale1 * 1.11485386310037, scale1 * 0.255960787275844), 0, 0, 0);
            poly42.AddVertexAt(5, new Point2d(scale1 * 1.11906187280799, scale1 * 0.256521855236861), 0, 0, 0);
            poly42.AddVertexAt(6, new Point2d(scale1 * 1.12256854756433, scale1 * 0.258625860090668), 0, 0, 0);
            poly42.AddVertexAt(7, new Point2d(scale1 * 1.11906187280799, scale1 * 0.263535204749552), 0, 0, 0);
            poly42.AddVertexAt(8, new Point2d(scale1 * 1.11108075824353, scale1 * 0.266735842436258), 0, 0, 0);
            poly42.Closed = true;
            poly42.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly42.Layer = "0";
            poly42.Color = color_GP;
            poly42.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly42);
            Polyline poly43 = new Polyline();
            poly43.AddVertexAt(0, new Point2d(scale1 * 1.18797544496318, scale1 * 0.290759752403364), 0, 0, 0);
            poly43.AddVertexAt(1, new Point2d(scale1 * 1.18470254852392, scale1 * 0.288936281530065), 0, 0, 0);
            poly43.AddVertexAt(2, new Point2d(scale1 * 1.18259854367012, scale1 * 0.283746402890672), 0, 0, 0);
            poly43.AddVertexAt(3, new Point2d(scale1 * 1.18259854367012, scale1 * 0.281221597066103), 0, 0, 0);
            poly43.AddVertexAt(4, new Point2d(scale1 * 1.18189720871885, scale1 * 0.275751184446203), 0, 0, 0);
            poly43.AddVertexAt(5, new Point2d(scale1 * 1.18540388347519, scale1 * 0.273366645611888), 0, 0, 0);
            poly43.AddVertexAt(6, new Point2d(scale1 * 1.19031322813408, scale1 * 0.272525043670365), 0, 0, 0);
            poly43.AddVertexAt(7, new Point2d(scale1 * 1.20176836567148, scale1 * 0.272945844641126), 0, 0, 0);
            poly43.AddVertexAt(8, new Point2d(scale1 * 1.20480748379364, scale1 * 0.274488781533918), 0, 0, 0);
            poly43.AddVertexAt(9, new Point2d(scale1 * 1.205742597062, scale1 * 0.278275990270772), 0, 0, 0);
            poly43.AddVertexAt(10, new Point2d(scale1 * 1.20270347893983, scale1 * 0.28037999512458), 0, 0, 0);
            poly43.AddVertexAt(11, new Point2d(scale1 * 1.19989813913476, scale1 * 0.282764533958894), 0, 0, 0);
            poly43.AddVertexAt(12, new Point2d(scale1 * 1.19756035596386, scale1 * 0.286271208715241), 0, 0, 0);
            poly43.AddVertexAt(13, new Point2d(scale1 * 1.19569012942714, scale1 * 0.288796014539811), 0, 0, 0);
            poly43.AddVertexAt(14, new Point2d(scale1 * 1.19148211971953, scale1 * 0.289777883471588), 0, 0, 0);
            poly43.AddVertexAt(15, new Point2d(scale1 * 1.18797544496318, scale1 * 0.290759752403364), 0, 0, 0);
            poly43.Closed = true;
            poly43.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly43.Layer = "0";
            poly43.Color = color_GP;
            poly43.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly43);
            Polyline poly44 = new Polyline();
            poly44.AddVertexAt(0, new Point2d(scale1 * 1.15836727263602, scale1 * 0.274437775355651), 0, 0, 0);
            poly44.AddVertexAt(1, new Point2d(scale1 * 1.16608195709998, scale1 * 0.272754571472604), 0, 0, 0);
            poly44.AddVertexAt(2, new Point2d(scale1 * 1.16935485353923, scale1 * 0.268125760794228), 0, 0, 0);
            poly44.AddVertexAt(3, new Point2d(scale1 * 1.16608195709998, scale1 * 0.263637217106105), 0, 0, 0);
            poly44.AddVertexAt(4, new Point2d(scale1 * 1.15836727263602, scale1 * 0.261954013223059), 0, 0, 0);
            poly44.AddVertexAt(5, new Point2d(scale1 * 1.1497174749037, scale1 * 0.261392945262043), 0, 0, 0);
            poly44.AddVertexAt(6, new Point2d(scale1 * 1.1450419085619, scale1 * 0.26237481419382), 0, 0, 0);
            poly44.AddVertexAt(7, new Point2d(scale1 * 1.14130145548847, scale1 * 0.26489962001839), 0, 0, 0);
            poly44.AddVertexAt(8, new Point2d(scale1 * 1.13078143121943, scale1 * 0.268967362735751), 0, 0, 0);
            poly44.AddVertexAt(9, new Point2d(scale1 * 1.12517075160927, scale1 * 0.272894838462859), 0, 0, 0);
            poly44.AddVertexAt(10, new Point2d(scale1 * 1.12657342151181, scale1 * 0.27471830933616), 0, 0, 0);
            poly44.AddVertexAt(11, new Point2d(scale1 * 1.13498944092704, scale1 * 0.278365251082757), 0, 0, 0);
            poly44.AddVertexAt(12, new Point2d(scale1 * 1.14153523380556, scale1 * 0.279066586034028), 0, 0, 0);
            poly44.AddVertexAt(13, new Point2d(scale1 * 1.14574324351317, scale1 * 0.277804183121743), 0, 0, 0);
            poly44.AddVertexAt(14, new Point2d(scale1 * 1.15135392312332, scale1 * 0.276541780209458), 0, 0, 0);
            poly44.AddVertexAt(15, new Point2d(scale1 * 1.15579571114803, scale1 * 0.275840445258188), 0, 0, 0);
            poly44.Closed = true;
            poly44.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly44.Layer = "0";
            poly44.Color = color_GP;
            poly44.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly44);
            Polyline poly45 = new Polyline();
            poly45.AddVertexAt(0, new Point2d(scale1 * 1.17889889791182, scale1 * 0.261890255500208), 0, 0, 0);
            poly45.AddVertexAt(1, new Point2d(scale1 * 1.17274102660718, scale1 * 0.262017770945895), 0, 0, 0);
            poly45.AddVertexAt(2, new Point2d(scale1 * 1.16743251686181, scale1 * 0.259722492923566), 0, 0, 0);
            poly45.AddVertexAt(3, new Point2d(scale1 * 1.16637081491274, scale1 * 0.256662122227115), 0, 0, 0);
            poly45.AddVertexAt(4, new Point2d(scale1 * 1.17167932465811, scale1 * 0.255769514107321), 0, 0, 0);
            poly45.AddVertexAt(5, new Point2d(scale1 * 1.17826187674237, scale1 * 0.2567896376728), 0, 0, 0);
            poly45.AddVertexAt(6, new Point2d(scale1 * 1.18208400375904, scale1 * 0.259467462032192), 0, 0, 0);
            poly45.AddVertexAt(7, new Point2d(scale1 * 1.17889889791182, scale1 * 0.261890255500208), 0, 0, 0);
            poly45.Closed = true;
            poly45.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly45.Layer = "0";
            poly45.Color = color_GP;
            poly45.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly45);
            Polyline poly46 = new Polyline();
            poly46.AddVertexAt(0, new Point2d(scale1 * 1.19647313832164, scale1 * 0.248628649148944), 0, 0, 0);
            poly46.AddVertexAt(1, new Point2d(scale1 * 1.19670691663873, scale1 * 0.243298503519296), 0, 0, 0);
            poly46.AddVertexAt(2, new Point2d(scale1 * 1.19507046841911, scale1 * 0.242036100607014), 0, 0, 0);
            poly46.AddVertexAt(3, new Point2d(scale1 * 1.19086245871149, scale1 * 0.240633430704475), 0, 0, 0);
            poly46.AddVertexAt(4, new Point2d(scale1 * 1.18688822732096, scale1 * 0.240773697694727), 0, 0, 0);
            poly46.AddVertexAt(5, new Point2d(scale1 * 1.1838491091988, scale1 * 0.242036100607014), 0, 0, 0);
            poly46.AddVertexAt(6, new Point2d(scale1 * 1.18057621275954, scale1 * 0.243298503519296), 0, 0, 0);
            poly46.AddVertexAt(7, new Point2d(scale1 * 1.17753709463737, scale1 * 0.244560906431583), 0, 0, 0);
            poly46.AddVertexAt(8, new Point2d(scale1 * 1.17566686810066, scale1 * 0.246945445265898), 0, 0, 0);
            poly46.AddVertexAt(9, new Point2d(scale1 * 1.17566686810066, scale1 * 0.248628649148944), 0, 0, 0);
            poly46.AddVertexAt(10, new Point2d(scale1 * 1.179173542857, scale1 * 0.252275590895543), 0, 0, 0);
            poly46.AddVertexAt(11, new Point2d(scale1 * 1.18595311405261, scale1 * 0.253818527788335), 0, 0, 0);
            poly46.AddVertexAt(12, new Point2d(scale1 * 1.19273268524821, scale1 * 0.252275590895543), 0, 0, 0);
            poly46.Closed = true;
            poly46.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly46.Layer = "0";
            poly46.Color = color_GP;
            poly46.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly46);
            Polyline poly47 = new Polyline();
            poly47.AddVertexAt(0, new Point2d(scale1 * 1.16210772570945, scale1 * 0.243438770509551), 0, 0, 0);
            poly47.AddVertexAt(1, new Point2d(scale1 * 1.15930238590438, scale1 * 0.241895833616759), 0, 0, 0);
            poly47.AddVertexAt(2, new Point2d(scale1 * 1.16023749917273, scale1 * 0.240072362743458), 0, 0, 0);
            poly47.AddVertexAt(3, new Point2d(scale1 * 1.16164016907527, scale1 * 0.237968357889651), 0, 0, 0);
            poly47.AddVertexAt(4, new Point2d(scale1 * 1.16467928719744, scale1 * 0.237687823909144), 0, 0, 0);
            poly47.AddVertexAt(5, new Point2d(scale1 * 1.16818596195379, scale1 * 0.23937102779219), 0, 0, 0);
            poly47.AddVertexAt(6, new Point2d(scale1 * 1.16818596195379, scale1 * 0.241615299636252), 0, 0, 0);
            poly47.AddVertexAt(7, new Point2d(scale1 * 1.16584817878289, scale1 * 0.244140105460821), 0, 0, 0);
            poly47.AddVertexAt(8, new Point2d(scale1 * 1.16210772570945, scale1 * 0.243438770509551), 0, 0, 0);
            poly47.Closed = true;
            poly47.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly47.Layer = "0";
            poly47.Color = color_GP;
            poly47.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly47);
            Polyline poly48 = new Polyline();
            poly48.AddVertexAt(0, new Point2d(scale1 * 1.11652095387695, scale1 * 0.248628649148944), 0, 0, 0);
            poly48.AddVertexAt(1, new Point2d(scale1 * 1.12306674675547, scale1 * 0.250452120022245), 0, 0, 0);
            poly48.AddVertexAt(2, new Point2d(scale1 * 1.1359245541954, scale1 * 0.249750785070974), 0, 0, 0);
            poly48.AddVertexAt(3, new Point2d(scale1 * 1.14247034707391, scale1 * 0.248628649148944), 0, 0, 0);
            poly48.AddVertexAt(4, new Point2d(scale1 * 1.14363923865936, scale1 * 0.246805178275643), 0, 0, 0);
            poly48.AddVertexAt(5, new Point2d(scale1 * 1.14387301697645, scale1 * 0.244981707402343), 0, 0, 0);
            poly48.AddVertexAt(6, new Point2d(scale1 * 1.14410679529354, scale1 * 0.242737435558282), 0, 0, 0);
            poly48.AddVertexAt(7, new Point2d(scale1 * 1.14597702183026, scale1 * 0.240773697694727), 0, 0, 0);
            poly48.AddVertexAt(8, new Point2d(scale1 * 1.14457435192772, scale1 * 0.237407289928637), 0, 0, 0);
            poly48.AddVertexAt(9, new Point2d(scale1 * 1.13943122895175, scale1 * 0.235583819055336), 0, 0, 0);
            poly48.AddVertexAt(10, new Point2d(scale1 * 1.12937876131689, scale1 * 0.235303285074829), 0, 0, 0);
            poly48.AddVertexAt(11, new Point2d(scale1 * 1.124469416658, scale1 * 0.236144887016352), 0, 0, 0);
            poly48.AddVertexAt(12, new Point2d(scale1 * 1.12096274190166, scale1 * 0.237828090899398), 0, 0, 0);
            poly48.AddVertexAt(13, new Point2d(scale1 * 1.1176898454624, scale1 * 0.241054231675236), 0, 0, 0);
            poly48.AddVertexAt(14, new Point2d(scale1 * 1.11581961892568, scale1 * 0.244420639441328), 0, 0, 0);
            poly48.Closed = true;
            poly48.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly48.Layer = "0";
            poly48.Color = color_GP;
            poly48.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly48);
            Polyline poly49 = new Polyline();
            poly49.AddVertexAt(0, new Point2d(scale1 * 1.06028574100171, scale1 * 0.3), 0, 0, 0);
            poly49.AddVertexAt(1, new Point2d(scale1 * 1.06164016907527, scale1 * 0.297968357889651), 0, 0, 0);
            poly49.AddVertexAt(2, new Point2d(scale1 * 1.06467928719744, scale1 * 0.297687823909144), 0, 0, 0);
            poly49.AddVertexAt(3, new Point2d(scale1 * 1.06818596195379, scale1 * 0.29937102779219), 0, 0, 0);
            poly49.AddVertexAt(4, new Point2d(scale1 * 1.06818596195379, scale1 * 0.3), 0, 0, 0);
            poly49.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly49.Layer = "0";
            poly49.Color = color_GP;
            poly49.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly49);
            Polyline poly50 = new Polyline();
            poly50.AddVertexAt(0, new Point2d(scale1 * 1.04565464779079, scale1 * 0.3), 0, 0, 0);
            poly50.AddVertexAt(1, new Point2d(scale1 * 1.04457435192772, scale1 * 0.297407289928637), 0, 0, 0);
            poly50.AddVertexAt(2, new Point2d(scale1 * 1.03943122895175, scale1 * 0.295583819055336), 0, 0, 0);
            poly50.AddVertexAt(3, new Point2d(scale1 * 1.02937876131689, scale1 * 0.295303285074829), 0, 0, 0);
            poly50.AddVertexAt(4, new Point2d(scale1 * 1.024469416658, scale1 * 0.296144887016352), 0, 0, 0);
            poly50.AddVertexAt(5, new Point2d(scale1 * 1.02096274190166, scale1 * 0.297828090899398), 0, 0, 0);
            poly50.AddVertexAt(6, new Point2d(scale1 * 1.01875935585757, scale1 * 0.3), 0, 0, 0);
            poly50.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly50.Layer = "0";
            poly50.Color = color_GP;
            poly50.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly50);
            Polyline poly51 = new Polyline();
            poly51.AddVertexAt(0, new Point2d(scale1 * 1.05836727263602, scale1 * 0.282994061761134), 0, 0, 0);
            poly51.AddVertexAt(1, new Point2d(scale1 * 1.0564970460993, scale1 * 0.282853794770882), 0, 0, 0);
            poly51.AddVertexAt(2, new Point2d(scale1 * 1.0541592629284, scale1 * 0.28355512972215), 0, 0, 0);
            poly51.AddVertexAt(3, new Point2d(scale1 * 1.05228903639168, scale1 * 0.285518867585705), 0, 0, 0);
            poly51.AddVertexAt(4, new Point2d(scale1 * 1.05392548461131, scale1 * 0.287061804478495), 0, 0, 0);
            poly51.AddVertexAt(5, new Point2d(scale1 * 1.05836727263602, scale1 * 0.286220202536974), 0, 0, 0);
            poly51.AddVertexAt(6, new Point2d(scale1 * 1.05953616422147, scale1 * 0.284817532634435), 0, 0, 0);
            poly51.Closed = true;
            poly51.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly51.Layer = "0";
            poly51.Color = color_GP;
            poly51.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly51);
            Polyline poly52 = new Polyline();
            poly52.AddVertexAt(0, new Point2d(scale1 * 1.0679521836367, scale1 * 0.28355512972215), 0, 0, 0);
            poly52.AddVertexAt(1, new Point2d(scale1 * 1.06608195709998, scale1 * 0.282012192829359), 0, 0, 0);
            poly52.AddVertexAt(2, new Point2d(scale1 * 1.06584817878289, scale1 * 0.280048454965803), 0, 0, 0);
            poly52.AddVertexAt(3, new Point2d(scale1 * 1.06841974027088, scale1 * 0.278505518073012), 0, 0, 0);
            poly52.AddVertexAt(4, new Point2d(scale1 * 1.07169263671013, scale1 * 0.278926319043773), 0, 0, 0);
            poly52.AddVertexAt(5, new Point2d(scale1 * 1.0723939716614, scale1 * 0.281030323897581), 0, 0, 0);
            poly52.AddVertexAt(6, new Point2d(scale1 * 1.07192641502722, scale1 * 0.282713527780627), 0, 0, 0);
            poly52.AddVertexAt(7, new Point2d(scale1 * 1.0679521836367, scale1 * 0.28355512972215), 0, 0, 0);
            poly52.Closed = true;
            poly52.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly52.Layer = "0";
            poly52.Color = color_GP;
            poly52.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly52);
            Polyline poly53 = new Polyline();
            poly53.AddVertexAt(0, new Point2d(scale1 * 1.01657711212705, scale1 * 0.287355090003573), 0, 0, 0);
            poly53.AddVertexAt(1, new Point2d(scale1 * 1.01423932895615, scale1 * 0.285251085149766), 0, 0, 0);
            poly53.AddVertexAt(2, new Point2d(scale1 * 1.01564199885869, scale1 * 0.283427614276465), 0, 0, 0);
            poly53.AddVertexAt(3, new Point2d(scale1 * 1.01961623024921, scale1 * 0.282726279325197), 0, 0, 0);
            poly53.AddVertexAt(4, new Point2d(scale1 * 1.02265534837138, scale1 * 0.283988682237481), 0, 0, 0);
            poly53.AddVertexAt(5, new Point2d(scale1 * 1.02265534837138, scale1 * 0.286653755052303), 0, 0, 0);
            poly53.AddVertexAt(6, new Point2d(scale1 * 1.01657711212705, scale1 * 0.287355090003573), 0, 0, 0);
            poly53.Closed = true;
            poly53.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly53.Layer = "0";
            poly53.Color = color_GP;
            poly53.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly53);
            Polyline poly54 = new Polyline();
            poly54.AddVertexAt(0, new Point2d(scale1 * 1.01108075824353, scale1 * 0.266735842436258), 0, 0, 0);
            poly54.AddVertexAt(1, new Point2d(scale1 * 1.00947696180731, scale1 * 0.263254670769046), 0, 0, 0);
            poly54.AddVertexAt(2, new Point2d(scale1 * 1.0097107401244, scale1 * 0.259607729022446), 0, 0, 0);
            poly54.AddVertexAt(3, new Point2d(scale1 * 1.01064585339276, scale1 * 0.256802389217368), 0, 0, 0);
            poly54.AddVertexAt(4, new Point2d(scale1 * 1.01485386310037, scale1 * 0.255960787275844), 0, 0, 0);
            poly54.AddVertexAt(5, new Point2d(scale1 * 1.01906187280799, scale1 * 0.256521855236861), 0, 0, 0);
            poly54.AddVertexAt(6, new Point2d(scale1 * 1.02256854756433, scale1 * 0.258625860090668), 0, 0, 0);
            poly54.AddVertexAt(7, new Point2d(scale1 * 1.01906187280799, scale1 * 0.263535204749552), 0, 0, 0);
            poly54.AddVertexAt(8, new Point2d(scale1 * 1.01108075824353, scale1 * 0.266735842436258), 0, 0, 0);
            poly54.Closed = true;
            poly54.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly54.Layer = "0";
            poly54.Color = color_GP;
            poly54.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly54);
            Polyline poly55 = new Polyline();
            poly55.AddVertexAt(0, new Point2d(scale1 * 1.08797544496318, scale1 * 0.290759752403364), 0, 0, 0);
            poly55.AddVertexAt(1, new Point2d(scale1 * 1.08470254852392, scale1 * 0.288936281530065), 0, 0, 0);
            poly55.AddVertexAt(2, new Point2d(scale1 * 1.08259854367012, scale1 * 0.283746402890672), 0, 0, 0);
            poly55.AddVertexAt(3, new Point2d(scale1 * 1.08259854367012, scale1 * 0.281221597066103), 0, 0, 0);
            poly55.AddVertexAt(4, new Point2d(scale1 * 1.08189720871885, scale1 * 0.275751184446203), 0, 0, 0);
            poly55.AddVertexAt(5, new Point2d(scale1 * 1.08540388347519, scale1 * 0.273366645611888), 0, 0, 0);
            poly55.AddVertexAt(6, new Point2d(scale1 * 1.09031322813408, scale1 * 0.272525043670365), 0, 0, 0);
            poly55.AddVertexAt(7, new Point2d(scale1 * 1.10176836567147, scale1 * 0.272945844641126), 0, 0, 0);
            poly55.AddVertexAt(8, new Point2d(scale1 * 1.10480748379364, scale1 * 0.274488781533918), 0, 0, 0);
            poly55.AddVertexAt(9, new Point2d(scale1 * 1.105742597062, scale1 * 0.278275990270772), 0, 0, 0);
            poly55.AddVertexAt(10, new Point2d(scale1 * 1.10270347893983, scale1 * 0.28037999512458), 0, 0, 0);
            poly55.AddVertexAt(11, new Point2d(scale1 * 1.09989813913476, scale1 * 0.282764533958894), 0, 0, 0);
            poly55.AddVertexAt(12, new Point2d(scale1 * 1.09756035596386, scale1 * 0.286271208715241), 0, 0, 0);
            poly55.AddVertexAt(13, new Point2d(scale1 * 1.09569012942714, scale1 * 0.288796014539811), 0, 0, 0);
            poly55.AddVertexAt(14, new Point2d(scale1 * 1.09148211971953, scale1 * 0.289777883471588), 0, 0, 0);
            poly55.AddVertexAt(15, new Point2d(scale1 * 1.08797544496318, scale1 * 0.290759752403364), 0, 0, 0);
            poly55.Closed = true;
            poly55.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly55.Layer = "0";
            poly55.Color = color_GP;
            poly55.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly55);
            Polyline poly56 = new Polyline();
            poly56.AddVertexAt(0, new Point2d(scale1 * 1.05836727263602, scale1 * 0.274437775355651), 0, 0, 0);
            poly56.AddVertexAt(1, new Point2d(scale1 * 1.06608195709998, scale1 * 0.272754571472604), 0, 0, 0);
            poly56.AddVertexAt(2, new Point2d(scale1 * 1.06935485353923, scale1 * 0.268125760794228), 0, 0, 0);
            poly56.AddVertexAt(3, new Point2d(scale1 * 1.06608195709998, scale1 * 0.263637217106105), 0, 0, 0);
            poly56.AddVertexAt(4, new Point2d(scale1 * 1.05836727263602, scale1 * 0.261954013223059), 0, 0, 0);
            poly56.AddVertexAt(5, new Point2d(scale1 * 1.0497174749037, scale1 * 0.261392945262043), 0, 0, 0);
            poly56.AddVertexAt(6, new Point2d(scale1 * 1.0450419085619, scale1 * 0.26237481419382), 0, 0, 0);
            poly56.AddVertexAt(7, new Point2d(scale1 * 1.04130145548847, scale1 * 0.26489962001839), 0, 0, 0);
            poly56.AddVertexAt(8, new Point2d(scale1 * 1.03078143121943, scale1 * 0.268967362735751), 0, 0, 0);
            poly56.AddVertexAt(9, new Point2d(scale1 * 1.02517075160927, scale1 * 0.272894838462859), 0, 0, 0);
            poly56.AddVertexAt(10, new Point2d(scale1 * 1.02657342151181, scale1 * 0.27471830933616), 0, 0, 0);
            poly56.AddVertexAt(11, new Point2d(scale1 * 1.03498944092704, scale1 * 0.278365251082757), 0, 0, 0);
            poly56.AddVertexAt(12, new Point2d(scale1 * 1.04153523380556, scale1 * 0.279066586034028), 0, 0, 0);
            poly56.AddVertexAt(13, new Point2d(scale1 * 1.04574324351317, scale1 * 0.277804183121743), 0, 0, 0);
            poly56.AddVertexAt(14, new Point2d(scale1 * 1.05135392312332, scale1 * 0.276541780209458), 0, 0, 0);
            poly56.AddVertexAt(15, new Point2d(scale1 * 1.05579571114803, scale1 * 0.275840445258188), 0, 0, 0);
            poly56.Closed = true;
            poly56.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly56.Layer = "0";
            poly56.Color = color_GP;
            poly56.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly56);
            Polyline poly57 = new Polyline();
            poly57.AddVertexAt(0, new Point2d(scale1 * 1.07889889791182, scale1 * 0.261890255500208), 0, 0, 0);
            poly57.AddVertexAt(1, new Point2d(scale1 * 1.07274102660718, scale1 * 0.262017770945895), 0, 0, 0);
            poly57.AddVertexAt(2, new Point2d(scale1 * 1.06743251686181, scale1 * 0.259722492923566), 0, 0, 0);
            poly57.AddVertexAt(3, new Point2d(scale1 * 1.06637081491274, scale1 * 0.256662122227115), 0, 0, 0);
            poly57.AddVertexAt(4, new Point2d(scale1 * 1.07167932465811, scale1 * 0.255769514107321), 0, 0, 0);
            poly57.AddVertexAt(5, new Point2d(scale1 * 1.07826187674237, scale1 * 0.2567896376728), 0, 0, 0);
            poly57.AddVertexAt(6, new Point2d(scale1 * 1.08208400375904, scale1 * 0.259467462032192), 0, 0, 0);
            poly57.AddVertexAt(7, new Point2d(scale1 * 1.07889889791182, scale1 * 0.261890255500208), 0, 0, 0);
            poly57.Closed = true;
            poly57.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly57.Layer = "0";
            poly57.Color = color_GP;
            poly57.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly57);
            Polyline poly58 = new Polyline();
            poly58.AddVertexAt(0, new Point2d(scale1 * 1.09647313832164, scale1 * 0.248628649148944), 0, 0, 0);
            poly58.AddVertexAt(1, new Point2d(scale1 * 1.09670691663873, scale1 * 0.243298503519296), 0, 0, 0);
            poly58.AddVertexAt(2, new Point2d(scale1 * 1.09507046841911, scale1 * 0.242036100607014), 0, 0, 0);
            poly58.AddVertexAt(3, new Point2d(scale1 * 1.09086245871149, scale1 * 0.240633430704475), 0, 0, 0);
            poly58.AddVertexAt(4, new Point2d(scale1 * 1.08688822732096, scale1 * 0.240773697694727), 0, 0, 0);
            poly58.AddVertexAt(5, new Point2d(scale1 * 1.0838491091988, scale1 * 0.242036100607014), 0, 0, 0);
            poly58.AddVertexAt(6, new Point2d(scale1 * 1.08057621275954, scale1 * 0.243298503519296), 0, 0, 0);
            poly58.AddVertexAt(7, new Point2d(scale1 * 1.07753709463737, scale1 * 0.244560906431583), 0, 0, 0);
            poly58.AddVertexAt(8, new Point2d(scale1 * 1.07566686810066, scale1 * 0.246945445265898), 0, 0, 0);
            poly58.AddVertexAt(9, new Point2d(scale1 * 1.07566686810066, scale1 * 0.248628649148944), 0, 0, 0);
            poly58.AddVertexAt(10, new Point2d(scale1 * 1.079173542857, scale1 * 0.252275590895543), 0, 0, 0);
            poly58.AddVertexAt(11, new Point2d(scale1 * 1.08595311405261, scale1 * 0.253818527788335), 0, 0, 0);
            poly58.AddVertexAt(12, new Point2d(scale1 * 1.09273268524821, scale1 * 0.252275590895543), 0, 0, 0);
            poly58.Closed = true;
            poly58.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly58.Layer = "0";
            poly58.Color = color_GP;
            poly58.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly58);
            Polyline poly59 = new Polyline();
            poly59.AddVertexAt(0, new Point2d(scale1 * 1.06210772570945, scale1 * 0.243438770509551), 0, 0, 0);
            poly59.AddVertexAt(1, new Point2d(scale1 * 1.05930238590438, scale1 * 0.241895833616759), 0, 0, 0);
            poly59.AddVertexAt(2, new Point2d(scale1 * 1.06023749917273, scale1 * 0.240072362743458), 0, 0, 0);
            poly59.AddVertexAt(3, new Point2d(scale1 * 1.06164016907527, scale1 * 0.237968357889651), 0, 0, 0);
            poly59.AddVertexAt(4, new Point2d(scale1 * 1.06467928719744, scale1 * 0.237687823909144), 0, 0, 0);
            poly59.AddVertexAt(5, new Point2d(scale1 * 1.06818596195379, scale1 * 0.23937102779219), 0, 0, 0);
            poly59.AddVertexAt(6, new Point2d(scale1 * 1.06818596195379, scale1 * 0.241615299636252), 0, 0, 0);
            poly59.AddVertexAt(7, new Point2d(scale1 * 1.06584817878289, scale1 * 0.244140105460821), 0, 0, 0);
            poly59.AddVertexAt(8, new Point2d(scale1 * 1.06210772570945, scale1 * 0.243438770509551), 0, 0, 0);
            poly59.Closed = true;
            poly59.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly59.Layer = "0";
            poly59.Color = color_GP;
            poly59.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly59);
            Polyline poly60 = new Polyline();
            poly60.AddVertexAt(0, new Point2d(scale1 * 1.01652095387695, scale1 * 0.248628649148944), 0, 0, 0);
            poly60.AddVertexAt(1, new Point2d(scale1 * 1.02306674675547, scale1 * 0.250452120022245), 0, 0, 0);
            poly60.AddVertexAt(2, new Point2d(scale1 * 1.0359245541954, scale1 * 0.249750785070974), 0, 0, 0);
            poly60.AddVertexAt(3, new Point2d(scale1 * 1.04247034707391, scale1 * 0.248628649148944), 0, 0, 0);
            poly60.AddVertexAt(4, new Point2d(scale1 * 1.04363923865936, scale1 * 0.246805178275643), 0, 0, 0);
            poly60.AddVertexAt(5, new Point2d(scale1 * 1.04387301697645, scale1 * 0.244981707402343), 0, 0, 0);
            poly60.AddVertexAt(6, new Point2d(scale1 * 1.04410679529354, scale1 * 0.242737435558282), 0, 0, 0);
            poly60.AddVertexAt(7, new Point2d(scale1 * 1.04597702183026, scale1 * 0.240773697694727), 0, 0, 0);
            poly60.AddVertexAt(8, new Point2d(scale1 * 1.04457435192772, scale1 * 0.237407289928637), 0, 0, 0);
            poly60.AddVertexAt(9, new Point2d(scale1 * 1.03943122895175, scale1 * 0.235583819055336), 0, 0, 0);
            poly60.AddVertexAt(10, new Point2d(scale1 * 1.02937876131689, scale1 * 0.235303285074829), 0, 0, 0);
            poly60.AddVertexAt(11, new Point2d(scale1 * 1.024469416658, scale1 * 0.236144887016352), 0, 0, 0);
            poly60.AddVertexAt(12, new Point2d(scale1 * 1.02096274190166, scale1 * 0.237828090899398), 0, 0, 0);
            poly60.AddVertexAt(13, new Point2d(scale1 * 1.0176898454624, scale1 * 0.241054231675236), 0, 0, 0);
            poly60.AddVertexAt(14, new Point2d(scale1 * 1.01581961892568, scale1 * 0.244420639441328), 0, 0, 0);
            poly60.Closed = true;
            poly60.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly60.Layer = "0";
            poly60.Color = color_GP;
            poly60.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly60);
            Polyline poly61 = new Polyline();
            poly61.AddVertexAt(0, new Point2d(scale1 * 0.960285741001707, scale1 * 0.3), 0, 0, 0);
            poly61.AddVertexAt(1, new Point2d(scale1 * 0.961640169075273, scale1 * 0.297968357889651), 0, 0, 0);
            poly61.AddVertexAt(2, new Point2d(scale1 * 0.964679287197439, scale1 * 0.297687823909144), 0, 0, 0);
            poly61.AddVertexAt(3, new Point2d(scale1 * 0.968185961953785, scale1 * 0.29937102779219), 0, 0, 0);
            poly61.AddVertexAt(4, new Point2d(scale1 * 0.968185961953785, scale1 * 0.3), 0, 0, 0);
            poly61.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly61.Layer = "0";
            poly61.Color = color_GP;
            poly61.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly61);
            Polyline poly62 = new Polyline();
            poly62.AddVertexAt(0, new Point2d(scale1 * 0.94565464779079, scale1 * 0.3), 0, 0, 0);
            poly62.AddVertexAt(1, new Point2d(scale1 * 0.944574351927722, scale1 * 0.297407289928637), 0, 0, 0);
            poly62.AddVertexAt(2, new Point2d(scale1 * 0.939431228951747, scale1 * 0.295583819055336), 0, 0, 0);
            poly62.AddVertexAt(3, new Point2d(scale1 * 0.929378761316889, scale1 * 0.295303285074829), 0, 0, 0);
            poly62.AddVertexAt(4, new Point2d(scale1 * 0.924469416658004, scale1 * 0.296144887016352), 0, 0, 0);
            poly62.AddVertexAt(5, new Point2d(scale1 * 0.920962741901658, scale1 * 0.297828090899398), 0, 0, 0);
            poly62.AddVertexAt(6, new Point2d(scale1 * 0.918759355857569, scale1 * 0.3), 0, 0, 0);
            poly62.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly62.Layer = "0";
            poly62.Color = color_GP;
            poly62.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly62);
            Polyline poly63 = new Polyline();
            poly63.AddVertexAt(0, new Point2d(scale1 * 0.958367272636016, scale1 * 0.282994061761134), 0, 0, 0);
            poly63.AddVertexAt(1, new Point2d(scale1 * 0.956497046099298, scale1 * 0.282853794770882), 0, 0, 0);
            poly63.AddVertexAt(2, new Point2d(scale1 * 0.954159262928401, scale1 * 0.28355512972215), 0, 0, 0);
            poly63.AddVertexAt(3, new Point2d(scale1 * 0.952289036391683, scale1 * 0.285518867585705), 0, 0, 0);
            poly63.AddVertexAt(4, new Point2d(scale1 * 0.953925484611311, scale1 * 0.287061804478495), 0, 0, 0);
            poly63.AddVertexAt(5, new Point2d(scale1 * 0.958367272636016, scale1 * 0.286220202536974), 0, 0, 0);
            poly63.AddVertexAt(6, new Point2d(scale1 * 0.959536164221465, scale1 * 0.284817532634435), 0, 0, 0);
            poly63.Closed = true;
            poly63.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly63.Layer = "0";
            poly63.Color = color_GP;
            poly63.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly63);
            Polyline poly64 = new Polyline();
            poly64.AddVertexAt(0, new Point2d(scale1 * 0.967952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly64.AddVertexAt(1, new Point2d(scale1 * 0.966081957099978, scale1 * 0.282012192829359), 0, 0, 0);
            poly64.AddVertexAt(2, new Point2d(scale1 * 0.965848178782888, scale1 * 0.280048454965803), 0, 0, 0);
            poly64.AddVertexAt(3, new Point2d(scale1 * 0.968419740270875, scale1 * 0.278505518073012), 0, 0, 0);
            poly64.AddVertexAt(4, new Point2d(scale1 * 0.971692636710131, scale1 * 0.278926319043773), 0, 0, 0);
            poly64.AddVertexAt(5, new Point2d(scale1 * 0.972393971661401, scale1 * 0.281030323897581), 0, 0, 0);
            poly64.AddVertexAt(6, new Point2d(scale1 * 0.971926415027221, scale1 * 0.282713527780627), 0, 0, 0);
            poly64.AddVertexAt(7, new Point2d(scale1 * 0.967952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly64.Closed = true;
            poly64.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly64.Layer = "0";
            poly64.Color = color_GP;
            poly64.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly64);
            Polyline poly65 = new Polyline();
            poly65.AddVertexAt(0, new Point2d(scale1 * 0.916577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly65.AddVertexAt(1, new Point2d(scale1 * 0.914239328956148, scale1 * 0.285251085149766), 0, 0, 0);
            poly65.AddVertexAt(2, new Point2d(scale1 * 0.915641998858686, scale1 * 0.283427614276465), 0, 0, 0);
            poly65.AddVertexAt(3, new Point2d(scale1 * 0.919616230249212, scale1 * 0.282726279325197), 0, 0, 0);
            poly65.AddVertexAt(4, new Point2d(scale1 * 0.922655348371378, scale1 * 0.283988682237481), 0, 0, 0);
            poly65.AddVertexAt(5, new Point2d(scale1 * 0.922655348371378, scale1 * 0.286653755052303), 0, 0, 0);
            poly65.AddVertexAt(6, new Point2d(scale1 * 0.916577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly65.Closed = true;
            poly65.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly65.Layer = "0";
            poly65.Color = color_GP;
            poly65.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly65);
            Polyline poly66 = new Polyline();
            poly66.AddVertexAt(0, new Point2d(scale1 * 0.911080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly66.AddVertexAt(1, new Point2d(scale1 * 0.909476961807309, scale1 * 0.263254670769046), 0, 0, 0);
            poly66.AddVertexAt(2, new Point2d(scale1 * 0.909710740124399, scale1 * 0.259607729022446), 0, 0, 0);
            poly66.AddVertexAt(3, new Point2d(scale1 * 0.910645853392758, scale1 * 0.256802389217368), 0, 0, 0);
            poly66.AddVertexAt(4, new Point2d(scale1 * 0.914853863100373, scale1 * 0.255960787275844), 0, 0, 0);
            poly66.AddVertexAt(5, new Point2d(scale1 * 0.919061872807989, scale1 * 0.256521855236861), 0, 0, 0);
            poly66.AddVertexAt(6, new Point2d(scale1 * 0.922568547564335, scale1 * 0.258625860090668), 0, 0, 0);
            poly66.AddVertexAt(7, new Point2d(scale1 * 0.919061872807989, scale1 * 0.263535204749552), 0, 0, 0);
            poly66.AddVertexAt(8, new Point2d(scale1 * 0.911080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly66.Closed = true;
            poly66.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly66.Layer = "0";
            poly66.Color = color_GP;
            poly66.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly66);
            Polyline poly67 = new Polyline();
            poly67.AddVertexAt(0, new Point2d(scale1 * 0.987975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly67.AddVertexAt(1, new Point2d(scale1 * 0.984702548523924, scale1 * 0.288936281530065), 0, 0, 0);
            poly67.AddVertexAt(2, new Point2d(scale1 * 0.982598543670117, scale1 * 0.283746402890672), 0, 0, 0);
            poly67.AddVertexAt(3, new Point2d(scale1 * 0.982598543670117, scale1 * 0.281221597066103), 0, 0, 0);
            poly67.AddVertexAt(4, new Point2d(scale1 * 0.981897208718848, scale1 * 0.275751184446203), 0, 0, 0);
            poly67.AddVertexAt(5, new Point2d(scale1 * 0.985403883475194, scale1 * 0.273366645611888), 0, 0, 0);
            poly67.AddVertexAt(6, new Point2d(scale1 * 0.990313228134079, scale1 * 0.272525043670365), 0, 0, 0);
            poly67.AddVertexAt(7, new Point2d(scale1 * 1.00176836567148, scale1 * 0.272945844641126), 0, 0, 0);
            poly67.AddVertexAt(8, new Point2d(scale1 * 1.00480748379364, scale1 * 0.274488781533918), 0, 0, 0);
            poly67.AddVertexAt(9, new Point2d(scale1 * 1.005742597062, scale1 * 0.278275990270772), 0, 0, 0);
            poly67.AddVertexAt(10, new Point2d(scale1 * 1.00270347893983, scale1 * 0.28037999512458), 0, 0, 0);
            poly67.AddVertexAt(11, new Point2d(scale1 * 0.999898139134758, scale1 * 0.282764533958894), 0, 0, 0);
            poly67.AddVertexAt(12, new Point2d(scale1 * 0.99756035596386, scale1 * 0.286271208715241), 0, 0, 0);
            poly67.AddVertexAt(13, new Point2d(scale1 * 0.995690129427142, scale1 * 0.288796014539811), 0, 0, 0);
            poly67.AddVertexAt(14, new Point2d(scale1 * 0.991482119719527, scale1 * 0.289777883471588), 0, 0, 0);
            poly67.AddVertexAt(15, new Point2d(scale1 * 0.987975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly67.Closed = true;
            poly67.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly67.Layer = "0";
            poly67.Color = color_GP;
            poly67.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly67);
            Polyline poly68 = new Polyline();
            poly68.AddVertexAt(0, new Point2d(scale1 * 0.958367272636016, scale1 * 0.274437775355651), 0, 0, 0);
            poly68.AddVertexAt(1, new Point2d(scale1 * 0.966081957099978, scale1 * 0.272754571472604), 0, 0, 0);
            poly68.AddVertexAt(2, new Point2d(scale1 * 0.969354853539234, scale1 * 0.268125760794228), 0, 0, 0);
            poly68.AddVertexAt(3, new Point2d(scale1 * 0.966081957099978, scale1 * 0.263637217106105), 0, 0, 0);
            poly68.AddVertexAt(4, new Point2d(scale1 * 0.958367272636016, scale1 * 0.261954013223059), 0, 0, 0);
            poly68.AddVertexAt(5, new Point2d(scale1 * 0.949717474903696, scale1 * 0.261392945262043), 0, 0, 0);
            poly68.AddVertexAt(6, new Point2d(scale1 * 0.945041908561901, scale1 * 0.26237481419382), 0, 0, 0);
            poly68.AddVertexAt(7, new Point2d(scale1 * 0.941301455488465, scale1 * 0.26489962001839), 0, 0, 0);
            poly68.AddVertexAt(8, new Point2d(scale1 * 0.930781431219427, scale1 * 0.268967362735751), 0, 0, 0);
            poly68.AddVertexAt(9, new Point2d(scale1 * 0.925170751609273, scale1 * 0.272894838462859), 0, 0, 0);
            poly68.AddVertexAt(10, new Point2d(scale1 * 0.926573421511812, scale1 * 0.27471830933616), 0, 0, 0);
            poly68.AddVertexAt(11, new Point2d(scale1 * 0.934989440927042, scale1 * 0.278365251082757), 0, 0, 0);
            poly68.AddVertexAt(12, new Point2d(scale1 * 0.941535233805555, scale1 * 0.279066586034028), 0, 0, 0);
            poly68.AddVertexAt(13, new Point2d(scale1 * 0.94574324351317, scale1 * 0.277804183121743), 0, 0, 0);
            poly68.AddVertexAt(14, new Point2d(scale1 * 0.951353923123324, scale1 * 0.276541780209458), 0, 0, 0);
            poly68.AddVertexAt(15, new Point2d(scale1 * 0.955795711148029, scale1 * 0.275840445258188), 0, 0, 0);
            poly68.Closed = true;
            poly68.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly68.Layer = "0";
            poly68.Color = color_GP;
            poly68.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly68);
            Polyline poly69 = new Polyline();
            poly69.AddVertexAt(0, new Point2d(scale1 * 0.978898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly69.AddVertexAt(1, new Point2d(scale1 * 0.972741026607183, scale1 * 0.262017770945895), 0, 0, 0);
            poly69.AddVertexAt(2, new Point2d(scale1 * 0.96743251686181, scale1 * 0.259722492923566), 0, 0, 0);
            poly69.AddVertexAt(3, new Point2d(scale1 * 0.966370814912736, scale1 * 0.256662122227115), 0, 0, 0);
            poly69.AddVertexAt(4, new Point2d(scale1 * 0.97167932465811, scale1 * 0.255769514107321), 0, 0, 0);
            poly69.AddVertexAt(5, new Point2d(scale1 * 0.978261876742373, scale1 * 0.2567896376728), 0, 0, 0);
            poly69.AddVertexAt(6, new Point2d(scale1 * 0.982084003759042, scale1 * 0.259467462032192), 0, 0, 0);
            poly69.AddVertexAt(7, new Point2d(scale1 * 0.978898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly69.Closed = true;
            poly69.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly69.Layer = "0";
            poly69.Color = color_GP;
            poly69.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly69);
            Polyline poly70 = new Polyline();
            poly70.AddVertexAt(0, new Point2d(scale1 * 0.996473138321644, scale1 * 0.248628649148944), 0, 0, 0);
            poly70.AddVertexAt(1, new Point2d(scale1 * 0.996706916638733, scale1 * 0.243298503519296), 0, 0, 0);
            poly70.AddVertexAt(2, new Point2d(scale1 * 0.995070468419106, scale1 * 0.242036100607014), 0, 0, 0);
            poly70.AddVertexAt(3, new Point2d(scale1 * 0.99086245871149, scale1 * 0.240633430704475), 0, 0, 0);
            poly70.AddVertexAt(4, new Point2d(scale1 * 0.986888227320964, scale1 * 0.240773697694727), 0, 0, 0);
            poly70.AddVertexAt(5, new Point2d(scale1 * 0.983849109198798, scale1 * 0.242036100607014), 0, 0, 0);
            poly70.AddVertexAt(6, new Point2d(scale1 * 0.980576212759542, scale1 * 0.243298503519296), 0, 0, 0);
            poly70.AddVertexAt(7, new Point2d(scale1 * 0.977537094637375, scale1 * 0.244560906431583), 0, 0, 0);
            poly70.AddVertexAt(8, new Point2d(scale1 * 0.975666868100657, scale1 * 0.246945445265898), 0, 0, 0);
            poly70.AddVertexAt(9, new Point2d(scale1 * 0.975666868100657, scale1 * 0.248628649148944), 0, 0, 0);
            poly70.AddVertexAt(10, new Point2d(scale1 * 0.979173542857003, scale1 * 0.252275590895543), 0, 0, 0);
            poly70.AddVertexAt(11, new Point2d(scale1 * 0.985953114052606, scale1 * 0.253818527788335), 0, 0, 0);
            poly70.AddVertexAt(12, new Point2d(scale1 * 0.992732685248208, scale1 * 0.252275590895543), 0, 0, 0);
            poly70.Closed = true;
            poly70.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly70.Layer = "0";
            poly70.Color = color_GP;
            poly70.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly70);
            Polyline poly71 = new Polyline();
            poly71.AddVertexAt(0, new Point2d(scale1 * 0.962107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly71.AddVertexAt(1, new Point2d(scale1 * 0.959302385904375, scale1 * 0.241895833616759), 0, 0, 0);
            poly71.AddVertexAt(2, new Point2d(scale1 * 0.960237499172734, scale1 * 0.240072362743458), 0, 0, 0);
            poly71.AddVertexAt(3, new Point2d(scale1 * 0.961640169075273, scale1 * 0.237968357889651), 0, 0, 0);
            poly71.AddVertexAt(4, new Point2d(scale1 * 0.964679287197439, scale1 * 0.237687823909144), 0, 0, 0);
            poly71.AddVertexAt(5, new Point2d(scale1 * 0.968185961953785, scale1 * 0.23937102779219), 0, 0, 0);
            poly71.AddVertexAt(6, new Point2d(scale1 * 0.968185961953785, scale1 * 0.241615299636252), 0, 0, 0);
            poly71.AddVertexAt(7, new Point2d(scale1 * 0.965848178782888, scale1 * 0.244140105460821), 0, 0, 0);
            poly71.AddVertexAt(8, new Point2d(scale1 * 0.962107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly71.Closed = true;
            poly71.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly71.Layer = "0";
            poly71.Color = color_GP;
            poly71.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly71);
            Polyline poly72 = new Polyline();
            poly72.AddVertexAt(0, new Point2d(scale1 * 0.916520953876953, scale1 * 0.248628649148944), 0, 0, 0);
            poly72.AddVertexAt(1, new Point2d(scale1 * 0.923066746755466, scale1 * 0.250452120022245), 0, 0, 0);
            poly72.AddVertexAt(2, new Point2d(scale1 * 0.935924554195401, scale1 * 0.249750785070974), 0, 0, 0);
            poly72.AddVertexAt(3, new Point2d(scale1 * 0.942470347073914, scale1 * 0.248628649148944), 0, 0, 0);
            poly72.AddVertexAt(4, new Point2d(scale1 * 0.943639238659363, scale1 * 0.246805178275643), 0, 0, 0);
            poly72.AddVertexAt(5, new Point2d(scale1 * 0.943873016976452, scale1 * 0.244981707402343), 0, 0, 0);
            poly72.AddVertexAt(6, new Point2d(scale1 * 0.944106795293542, scale1 * 0.242737435558282), 0, 0, 0);
            poly72.AddVertexAt(7, new Point2d(scale1 * 0.94597702183026, scale1 * 0.240773697694727), 0, 0, 0);
            poly72.AddVertexAt(8, new Point2d(scale1 * 0.944574351927722, scale1 * 0.237407289928637), 0, 0, 0);
            poly72.AddVertexAt(9, new Point2d(scale1 * 0.939431228951747, scale1 * 0.235583819055336), 0, 0, 0);
            poly72.AddVertexAt(10, new Point2d(scale1 * 0.929378761316889, scale1 * 0.235303285074829), 0, 0, 0);
            poly72.AddVertexAt(11, new Point2d(scale1 * 0.924469416658004, scale1 * 0.236144887016352), 0, 0, 0);
            poly72.AddVertexAt(12, new Point2d(scale1 * 0.920962741901658, scale1 * 0.237828090899398), 0, 0, 0);
            poly72.AddVertexAt(13, new Point2d(scale1 * 0.917689845462402, scale1 * 0.241054231675236), 0, 0, 0);
            poly72.AddVertexAt(14, new Point2d(scale1 * 0.915819618925684, scale1 * 0.244420639441328), 0, 0, 0);
            poly72.Closed = true;
            poly72.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly72.Layer = "0";
            poly72.Color = color_GP;
            poly72.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly72);
            Polyline poly73 = new Polyline();
            poly73.AddVertexAt(0, new Point2d(scale1 * 0.860285741001707, scale1 * 0.3), 0, 0, 0);
            poly73.AddVertexAt(1, new Point2d(scale1 * 0.861640169075273, scale1 * 0.297968357889651), 0, 0, 0);
            poly73.AddVertexAt(2, new Point2d(scale1 * 0.864679287197439, scale1 * 0.297687823909144), 0, 0, 0);
            poly73.AddVertexAt(3, new Point2d(scale1 * 0.868185961953785, scale1 * 0.29937102779219), 0, 0, 0);
            poly73.AddVertexAt(4, new Point2d(scale1 * 0.868185961953785, scale1 * 0.3), 0, 0, 0);
            poly73.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly73.Layer = "0";
            poly73.Color = color_GP;
            poly73.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly73);
            Polyline poly74 = new Polyline();
            poly74.AddVertexAt(0, new Point2d(scale1 * 0.84565464779079, scale1 * 0.3), 0, 0, 0);
            poly74.AddVertexAt(1, new Point2d(scale1 * 0.844574351927722, scale1 * 0.297407289928637), 0, 0, 0);
            poly74.AddVertexAt(2, new Point2d(scale1 * 0.839431228951747, scale1 * 0.295583819055336), 0, 0, 0);
            poly74.AddVertexAt(3, new Point2d(scale1 * 0.829378761316889, scale1 * 0.295303285074829), 0, 0, 0);
            poly74.AddVertexAt(4, new Point2d(scale1 * 0.824469416658004, scale1 * 0.296144887016352), 0, 0, 0);
            poly74.AddVertexAt(5, new Point2d(scale1 * 0.820962741901658, scale1 * 0.297828090899398), 0, 0, 0);
            poly74.AddVertexAt(6, new Point2d(scale1 * 0.818759355857569, scale1 * 0.3), 0, 0, 0);
            poly74.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly74.Layer = "0";
            poly74.Color = color_GP;
            poly74.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly74);
            Polyline poly75 = new Polyline();
            poly75.AddVertexAt(0, new Point2d(scale1 * 0.858367272636016, scale1 * 0.282994061761134), 0, 0, 0);
            poly75.AddVertexAt(1, new Point2d(scale1 * 0.856497046099298, scale1 * 0.282853794770882), 0, 0, 0);
            poly75.AddVertexAt(2, new Point2d(scale1 * 0.854159262928401, scale1 * 0.28355512972215), 0, 0, 0);
            poly75.AddVertexAt(3, new Point2d(scale1 * 0.852289036391683, scale1 * 0.285518867585705), 0, 0, 0);
            poly75.AddVertexAt(4, new Point2d(scale1 * 0.853925484611311, scale1 * 0.287061804478495), 0, 0, 0);
            poly75.AddVertexAt(5, new Point2d(scale1 * 0.858367272636016, scale1 * 0.286220202536974), 0, 0, 0);
            poly75.AddVertexAt(6, new Point2d(scale1 * 0.859536164221465, scale1 * 0.284817532634435), 0, 0, 0);
            poly75.Closed = true;
            poly75.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly75.Layer = "0";
            poly75.Color = color_GP;
            poly75.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly75);
            Polyline poly76 = new Polyline();
            poly76.AddVertexAt(0, new Point2d(scale1 * 0.867952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly76.AddVertexAt(1, new Point2d(scale1 * 0.866081957099978, scale1 * 0.282012192829359), 0, 0, 0);
            poly76.AddVertexAt(2, new Point2d(scale1 * 0.865848178782888, scale1 * 0.280048454965803), 0, 0, 0);
            poly76.AddVertexAt(3, new Point2d(scale1 * 0.868419740270875, scale1 * 0.278505518073012), 0, 0, 0);
            poly76.AddVertexAt(4, new Point2d(scale1 * 0.871692636710131, scale1 * 0.278926319043773), 0, 0, 0);
            poly76.AddVertexAt(5, new Point2d(scale1 * 0.872393971661401, scale1 * 0.281030323897581), 0, 0, 0);
            poly76.AddVertexAt(6, new Point2d(scale1 * 0.871926415027221, scale1 * 0.282713527780627), 0, 0, 0);
            poly76.AddVertexAt(7, new Point2d(scale1 * 0.867952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly76.Closed = true;
            poly76.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly76.Layer = "0";
            poly76.Color = color_GP;
            poly76.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly76);
            Polyline poly77 = new Polyline();
            poly77.AddVertexAt(0, new Point2d(scale1 * 0.816577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly77.AddVertexAt(1, new Point2d(scale1 * 0.814239328956148, scale1 * 0.285251085149766), 0, 0, 0);
            poly77.AddVertexAt(2, new Point2d(scale1 * 0.815641998858686, scale1 * 0.283427614276465), 0, 0, 0);
            poly77.AddVertexAt(3, new Point2d(scale1 * 0.819616230249212, scale1 * 0.282726279325197), 0, 0, 0);
            poly77.AddVertexAt(4, new Point2d(scale1 * 0.822655348371378, scale1 * 0.283988682237481), 0, 0, 0);
            poly77.AddVertexAt(5, new Point2d(scale1 * 0.822655348371378, scale1 * 0.286653755052303), 0, 0, 0);
            poly77.AddVertexAt(6, new Point2d(scale1 * 0.816577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly77.Closed = true;
            poly77.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly77.Layer = "0";
            poly77.Color = color_GP;
            poly77.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly77);
            Polyline poly78 = new Polyline();
            poly78.AddVertexAt(0, new Point2d(scale1 * 0.811080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly78.AddVertexAt(1, new Point2d(scale1 * 0.809476961807309, scale1 * 0.263254670769046), 0, 0, 0);
            poly78.AddVertexAt(2, new Point2d(scale1 * 0.809710740124399, scale1 * 0.259607729022446), 0, 0, 0);
            poly78.AddVertexAt(3, new Point2d(scale1 * 0.810645853392758, scale1 * 0.256802389217368), 0, 0, 0);
            poly78.AddVertexAt(4, new Point2d(scale1 * 0.814853863100373, scale1 * 0.255960787275844), 0, 0, 0);
            poly78.AddVertexAt(5, new Point2d(scale1 * 0.819061872807989, scale1 * 0.256521855236861), 0, 0, 0);
            poly78.AddVertexAt(6, new Point2d(scale1 * 0.822568547564335, scale1 * 0.258625860090668), 0, 0, 0);
            poly78.AddVertexAt(7, new Point2d(scale1 * 0.819061872807989, scale1 * 0.263535204749552), 0, 0, 0);
            poly78.AddVertexAt(8, new Point2d(scale1 * 0.811080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly78.Closed = true;
            poly78.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly78.Layer = "0";
            poly78.Color = color_GP;
            poly78.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly78);
            Polyline poly79 = new Polyline();
            poly79.AddVertexAt(0, new Point2d(scale1 * 0.887975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly79.AddVertexAt(1, new Point2d(scale1 * 0.884702548523924, scale1 * 0.288936281530065), 0, 0, 0);
            poly79.AddVertexAt(2, new Point2d(scale1 * 0.882598543670117, scale1 * 0.283746402890672), 0, 0, 0);
            poly79.AddVertexAt(3, new Point2d(scale1 * 0.882598543670117, scale1 * 0.281221597066103), 0, 0, 0);
            poly79.AddVertexAt(4, new Point2d(scale1 * 0.881897208718848, scale1 * 0.275751184446203), 0, 0, 0);
            poly79.AddVertexAt(5, new Point2d(scale1 * 0.885403883475194, scale1 * 0.273366645611888), 0, 0, 0);
            poly79.AddVertexAt(6, new Point2d(scale1 * 0.890313228134079, scale1 * 0.272525043670365), 0, 0, 0);
            poly79.AddVertexAt(7, new Point2d(scale1 * 0.901768365671475, scale1 * 0.272945844641126), 0, 0, 0);
            poly79.AddVertexAt(8, new Point2d(scale1 * 0.904807483793642, scale1 * 0.274488781533918), 0, 0, 0);
            poly79.AddVertexAt(9, new Point2d(scale1 * 0.905742597062001, scale1 * 0.278275990270772), 0, 0, 0);
            poly79.AddVertexAt(10, new Point2d(scale1 * 0.902703478939834, scale1 * 0.28037999512458), 0, 0, 0);
            poly79.AddVertexAt(11, new Point2d(scale1 * 0.899898139134758, scale1 * 0.282764533958894), 0, 0, 0);
            poly79.AddVertexAt(12, new Point2d(scale1 * 0.89756035596386, scale1 * 0.286271208715241), 0, 0, 0);
            poly79.AddVertexAt(13, new Point2d(scale1 * 0.895690129427142, scale1 * 0.288796014539811), 0, 0, 0);
            poly79.AddVertexAt(14, new Point2d(scale1 * 0.891482119719527, scale1 * 0.289777883471588), 0, 0, 0);
            poly79.AddVertexAt(15, new Point2d(scale1 * 0.887975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly79.Closed = true;
            poly79.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly79.Layer = "0";
            poly79.Color = color_GP;
            poly79.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly79);
            Polyline poly80 = new Polyline();
            poly80.AddVertexAt(0, new Point2d(scale1 * 0.858367272636016, scale1 * 0.274437775355651), 0, 0, 0);
            poly80.AddVertexAt(1, new Point2d(scale1 * 0.866081957099978, scale1 * 0.272754571472604), 0, 0, 0);
            poly80.AddVertexAt(2, new Point2d(scale1 * 0.869354853539234, scale1 * 0.268125760794228), 0, 0, 0);
            poly80.AddVertexAt(3, new Point2d(scale1 * 0.866081957099978, scale1 * 0.263637217106105), 0, 0, 0);
            poly80.AddVertexAt(4, new Point2d(scale1 * 0.858367272636016, scale1 * 0.261954013223059), 0, 0, 0);
            poly80.AddVertexAt(5, new Point2d(scale1 * 0.849717474903696, scale1 * 0.261392945262043), 0, 0, 0);
            poly80.AddVertexAt(6, new Point2d(scale1 * 0.845041908561901, scale1 * 0.26237481419382), 0, 0, 0);
            poly80.AddVertexAt(7, new Point2d(scale1 * 0.841301455488465, scale1 * 0.26489962001839), 0, 0, 0);
            poly80.AddVertexAt(8, new Point2d(scale1 * 0.830781431219427, scale1 * 0.268967362735751), 0, 0, 0);
            poly80.AddVertexAt(9, new Point2d(scale1 * 0.825170751609273, scale1 * 0.272894838462859), 0, 0, 0);
            poly80.AddVertexAt(10, new Point2d(scale1 * 0.826573421511812, scale1 * 0.27471830933616), 0, 0, 0);
            poly80.AddVertexAt(11, new Point2d(scale1 * 0.834989440927042, scale1 * 0.278365251082757), 0, 0, 0);
            poly80.AddVertexAt(12, new Point2d(scale1 * 0.841535233805555, scale1 * 0.279066586034028), 0, 0, 0);
            poly80.AddVertexAt(13, new Point2d(scale1 * 0.84574324351317, scale1 * 0.277804183121743), 0, 0, 0);
            poly80.AddVertexAt(14, new Point2d(scale1 * 0.851353923123324, scale1 * 0.276541780209458), 0, 0, 0);
            poly80.AddVertexAt(15, new Point2d(scale1 * 0.855795711148029, scale1 * 0.275840445258188), 0, 0, 0);
            poly80.Closed = true;
            poly80.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly80.Layer = "0";
            poly80.Color = color_GP;
            poly80.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly80);
            Polyline poly81 = new Polyline();
            poly81.AddVertexAt(0, new Point2d(scale1 * 0.878898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly81.AddVertexAt(1, new Point2d(scale1 * 0.872741026607183, scale1 * 0.262017770945895), 0, 0, 0);
            poly81.AddVertexAt(2, new Point2d(scale1 * 0.86743251686181, scale1 * 0.259722492923566), 0, 0, 0);
            poly81.AddVertexAt(3, new Point2d(scale1 * 0.866370814912736, scale1 * 0.256662122227115), 0, 0, 0);
            poly81.AddVertexAt(4, new Point2d(scale1 * 0.87167932465811, scale1 * 0.255769514107321), 0, 0, 0);
            poly81.AddVertexAt(5, new Point2d(scale1 * 0.878261876742373, scale1 * 0.2567896376728), 0, 0, 0);
            poly81.AddVertexAt(6, new Point2d(scale1 * 0.882084003759042, scale1 * 0.259467462032192), 0, 0, 0);
            poly81.AddVertexAt(7, new Point2d(scale1 * 0.878898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly81.Closed = true;
            poly81.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly81.Layer = "0";
            poly81.Color = color_GP;
            poly81.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly81);
            Polyline poly82 = new Polyline();
            poly82.AddVertexAt(0, new Point2d(scale1 * 0.896473138321644, scale1 * 0.248628649148944), 0, 0, 0);
            poly82.AddVertexAt(1, new Point2d(scale1 * 0.896706916638733, scale1 * 0.243298503519296), 0, 0, 0);
            poly82.AddVertexAt(2, new Point2d(scale1 * 0.895070468419106, scale1 * 0.242036100607014), 0, 0, 0);
            poly82.AddVertexAt(3, new Point2d(scale1 * 0.89086245871149, scale1 * 0.240633430704475), 0, 0, 0);
            poly82.AddVertexAt(4, new Point2d(scale1 * 0.886888227320964, scale1 * 0.240773697694727), 0, 0, 0);
            poly82.AddVertexAt(5, new Point2d(scale1 * 0.883849109198798, scale1 * 0.242036100607014), 0, 0, 0);
            poly82.AddVertexAt(6, new Point2d(scale1 * 0.880576212759542, scale1 * 0.243298503519296), 0, 0, 0);
            poly82.AddVertexAt(7, new Point2d(scale1 * 0.877537094637375, scale1 * 0.244560906431583), 0, 0, 0);
            poly82.AddVertexAt(8, new Point2d(scale1 * 0.875666868100657, scale1 * 0.246945445265898), 0, 0, 0);
            poly82.AddVertexAt(9, new Point2d(scale1 * 0.875666868100657, scale1 * 0.248628649148944), 0, 0, 0);
            poly82.AddVertexAt(10, new Point2d(scale1 * 0.879173542857003, scale1 * 0.252275590895543), 0, 0, 0);
            poly82.AddVertexAt(11, new Point2d(scale1 * 0.885953114052606, scale1 * 0.253818527788335), 0, 0, 0);
            poly82.AddVertexAt(12, new Point2d(scale1 * 0.892732685248208, scale1 * 0.252275590895543), 0, 0, 0);
            poly82.Closed = true;
            poly82.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly82.Layer = "0";
            poly82.Color = color_GP;
            poly82.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly82);
            Polyline poly83 = new Polyline();
            poly83.AddVertexAt(0, new Point2d(scale1 * 0.862107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly83.AddVertexAt(1, new Point2d(scale1 * 0.859302385904375, scale1 * 0.241895833616759), 0, 0, 0);
            poly83.AddVertexAt(2, new Point2d(scale1 * 0.860237499172734, scale1 * 0.240072362743458), 0, 0, 0);
            poly83.AddVertexAt(3, new Point2d(scale1 * 0.861640169075273, scale1 * 0.237968357889651), 0, 0, 0);
            poly83.AddVertexAt(4, new Point2d(scale1 * 0.864679287197439, scale1 * 0.237687823909144), 0, 0, 0);
            poly83.AddVertexAt(5, new Point2d(scale1 * 0.868185961953785, scale1 * 0.23937102779219), 0, 0, 0);
            poly83.AddVertexAt(6, new Point2d(scale1 * 0.868185961953785, scale1 * 0.241615299636252), 0, 0, 0);
            poly83.AddVertexAt(7, new Point2d(scale1 * 0.865848178782888, scale1 * 0.244140105460821), 0, 0, 0);
            poly83.AddVertexAt(8, new Point2d(scale1 * 0.862107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly83.Closed = true;
            poly83.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly83.Layer = "0";
            poly83.Color = color_GP;
            poly83.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly83);
            Polyline poly84 = new Polyline();
            poly84.AddVertexAt(0, new Point2d(scale1 * 0.816520953876953, scale1 * 0.248628649148944), 0, 0, 0);
            poly84.AddVertexAt(1, new Point2d(scale1 * 0.823066746755466, scale1 * 0.250452120022245), 0, 0, 0);
            poly84.AddVertexAt(2, new Point2d(scale1 * 0.835924554195401, scale1 * 0.249750785070974), 0, 0, 0);
            poly84.AddVertexAt(3, new Point2d(scale1 * 0.842470347073914, scale1 * 0.248628649148944), 0, 0, 0);
            poly84.AddVertexAt(4, new Point2d(scale1 * 0.843639238659363, scale1 * 0.246805178275643), 0, 0, 0);
            poly84.AddVertexAt(5, new Point2d(scale1 * 0.843873016976453, scale1 * 0.244981707402343), 0, 0, 0);
            poly84.AddVertexAt(6, new Point2d(scale1 * 0.844106795293542, scale1 * 0.242737435558282), 0, 0, 0);
            poly84.AddVertexAt(7, new Point2d(scale1 * 0.84597702183026, scale1 * 0.240773697694727), 0, 0, 0);
            poly84.AddVertexAt(8, new Point2d(scale1 * 0.844574351927722, scale1 * 0.237407289928637), 0, 0, 0);
            poly84.AddVertexAt(9, new Point2d(scale1 * 0.839431228951747, scale1 * 0.235583819055336), 0, 0, 0);
            poly84.AddVertexAt(10, new Point2d(scale1 * 0.829378761316889, scale1 * 0.235303285074829), 0, 0, 0);
            poly84.AddVertexAt(11, new Point2d(scale1 * 0.824469416658004, scale1 * 0.236144887016352), 0, 0, 0);
            poly84.AddVertexAt(12, new Point2d(scale1 * 0.820962741901658, scale1 * 0.237828090899398), 0, 0, 0);
            poly84.AddVertexAt(13, new Point2d(scale1 * 0.817689845462402, scale1 * 0.241054231675236), 0, 0, 0);
            poly84.AddVertexAt(14, new Point2d(scale1 * 0.815819618925684, scale1 * 0.244420639441328), 0, 0, 0);
            poly84.Closed = true;
            poly84.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly84.Layer = "0";
            poly84.Color = color_GP;
            poly84.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly84);
            Polyline poly85 = new Polyline();
            poly85.AddVertexAt(0, new Point2d(scale1 * 0.760285741001707, scale1 * 0.3), 0, 0, 0);
            poly85.AddVertexAt(1, new Point2d(scale1 * 0.761640169075273, scale1 * 0.297968357889651), 0, 0, 0);
            poly85.AddVertexAt(2, new Point2d(scale1 * 0.764679287197439, scale1 * 0.297687823909144), 0, 0, 0);
            poly85.AddVertexAt(3, new Point2d(scale1 * 0.768185961953785, scale1 * 0.29937102779219), 0, 0, 0);
            poly85.AddVertexAt(4, new Point2d(scale1 * 0.768185961953785, scale1 * 0.3), 0, 0, 0);
            poly85.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly85.Layer = "0";
            poly85.Color = color_GP;
            poly85.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly85);
            Polyline poly86 = new Polyline();
            poly86.AddVertexAt(0, new Point2d(scale1 * 0.74565464779079, scale1 * 0.3), 0, 0, 0);
            poly86.AddVertexAt(1, new Point2d(scale1 * 0.744574351927722, scale1 * 0.297407289928637), 0, 0, 0);
            poly86.AddVertexAt(2, new Point2d(scale1 * 0.739431228951748, scale1 * 0.295583819055336), 0, 0, 0);
            poly86.AddVertexAt(3, new Point2d(scale1 * 0.729378761316889, scale1 * 0.295303285074829), 0, 0, 0);
            poly86.AddVertexAt(4, new Point2d(scale1 * 0.724469416658004, scale1 * 0.296144887016352), 0, 0, 0);
            poly86.AddVertexAt(5, new Point2d(scale1 * 0.720962741901658, scale1 * 0.297828090899398), 0, 0, 0);
            poly86.AddVertexAt(6, new Point2d(scale1 * 0.718759355857569, scale1 * 0.3), 0, 0, 0);
            poly86.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly86.Layer = "0";
            poly86.Color = color_GP;
            poly86.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly86);
            Polyline poly87 = new Polyline();
            poly87.AddVertexAt(0, new Point2d(scale1 * 0.758367272636016, scale1 * 0.282994061761134), 0, 0, 0);
            poly87.AddVertexAt(1, new Point2d(scale1 * 0.756497046099298, scale1 * 0.282853794770882), 0, 0, 0);
            poly87.AddVertexAt(2, new Point2d(scale1 * 0.754159262928401, scale1 * 0.28355512972215), 0, 0, 0);
            poly87.AddVertexAt(3, new Point2d(scale1 * 0.752289036391683, scale1 * 0.285518867585705), 0, 0, 0);
            poly87.AddVertexAt(4, new Point2d(scale1 * 0.753925484611311, scale1 * 0.287061804478495), 0, 0, 0);
            poly87.AddVertexAt(5, new Point2d(scale1 * 0.758367272636016, scale1 * 0.286220202536974), 0, 0, 0);
            poly87.AddVertexAt(6, new Point2d(scale1 * 0.759536164221465, scale1 * 0.284817532634435), 0, 0, 0);
            poly87.Closed = true;
            poly87.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly87.Layer = "0";
            poly87.Color = color_GP;
            poly87.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly87);
            Polyline poly88 = new Polyline();
            poly88.AddVertexAt(0, new Point2d(scale1 * 0.767952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly88.AddVertexAt(1, new Point2d(scale1 * 0.766081957099978, scale1 * 0.282012192829359), 0, 0, 0);
            poly88.AddVertexAt(2, new Point2d(scale1 * 0.765848178782888, scale1 * 0.280048454965803), 0, 0, 0);
            poly88.AddVertexAt(3, new Point2d(scale1 * 0.768419740270875, scale1 * 0.278505518073012), 0, 0, 0);
            poly88.AddVertexAt(4, new Point2d(scale1 * 0.771692636710131, scale1 * 0.278926319043773), 0, 0, 0);
            poly88.AddVertexAt(5, new Point2d(scale1 * 0.772393971661401, scale1 * 0.281030323897581), 0, 0, 0);
            poly88.AddVertexAt(6, new Point2d(scale1 * 0.771926415027221, scale1 * 0.282713527780627), 0, 0, 0);
            poly88.AddVertexAt(7, new Point2d(scale1 * 0.767952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly88.Closed = true;
            poly88.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly88.Layer = "0";
            poly88.Color = color_GP;
            poly88.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly88);
            Polyline poly89 = new Polyline();
            poly89.AddVertexAt(0, new Point2d(scale1 * 0.716577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly89.AddVertexAt(1, new Point2d(scale1 * 0.714239328956148, scale1 * 0.285251085149766), 0, 0, 0);
            poly89.AddVertexAt(2, new Point2d(scale1 * 0.715641998858686, scale1 * 0.283427614276465), 0, 0, 0);
            poly89.AddVertexAt(3, new Point2d(scale1 * 0.719616230249212, scale1 * 0.282726279325197), 0, 0, 0);
            poly89.AddVertexAt(4, new Point2d(scale1 * 0.722655348371378, scale1 * 0.283988682237481), 0, 0, 0);
            poly89.AddVertexAt(5, new Point2d(scale1 * 0.722655348371378, scale1 * 0.286653755052303), 0, 0, 0);
            poly89.AddVertexAt(6, new Point2d(scale1 * 0.716577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly89.Closed = true;
            poly89.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly89.Layer = "0";
            poly89.Color = color_GP;
            poly89.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly89);
            Polyline poly90 = new Polyline();
            poly90.AddVertexAt(0, new Point2d(scale1 * 0.711080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly90.AddVertexAt(1, new Point2d(scale1 * 0.70947696180731, scale1 * 0.263254670769046), 0, 0, 0);
            poly90.AddVertexAt(2, new Point2d(scale1 * 0.709710740124399, scale1 * 0.259607729022446), 0, 0, 0);
            poly90.AddVertexAt(3, new Point2d(scale1 * 0.710645853392758, scale1 * 0.256802389217368), 0, 0, 0);
            poly90.AddVertexAt(4, new Point2d(scale1 * 0.714853863100373, scale1 * 0.255960787275844), 0, 0, 0);
            poly90.AddVertexAt(5, new Point2d(scale1 * 0.719061872807989, scale1 * 0.256521855236861), 0, 0, 0);
            poly90.AddVertexAt(6, new Point2d(scale1 * 0.722568547564335, scale1 * 0.258625860090668), 0, 0, 0);
            poly90.AddVertexAt(7, new Point2d(scale1 * 0.719061872807989, scale1 * 0.263535204749552), 0, 0, 0);
            poly90.AddVertexAt(8, new Point2d(scale1 * 0.711080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly90.Closed = true;
            poly90.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly90.Layer = "0";
            poly90.Color = color_GP;
            poly90.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly90);
            Polyline poly91 = new Polyline();
            poly91.AddVertexAt(0, new Point2d(scale1 * 0.787975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly91.AddVertexAt(1, new Point2d(scale1 * 0.784702548523924, scale1 * 0.288936281530065), 0, 0, 0);
            poly91.AddVertexAt(2, new Point2d(scale1 * 0.782598543670117, scale1 * 0.283746402890672), 0, 0, 0);
            poly91.AddVertexAt(3, new Point2d(scale1 * 0.782598543670117, scale1 * 0.281221597066103), 0, 0, 0);
            poly91.AddVertexAt(4, new Point2d(scale1 * 0.781897208718848, scale1 * 0.275751184446203), 0, 0, 0);
            poly91.AddVertexAt(5, new Point2d(scale1 * 0.785403883475194, scale1 * 0.273366645611888), 0, 0, 0);
            poly91.AddVertexAt(6, new Point2d(scale1 * 0.790313228134079, scale1 * 0.272525043670365), 0, 0, 0);
            poly91.AddVertexAt(7, new Point2d(scale1 * 0.801768365671475, scale1 * 0.272945844641126), 0, 0, 0);
            poly91.AddVertexAt(8, new Point2d(scale1 * 0.804807483793642, scale1 * 0.274488781533918), 0, 0, 0);
            poly91.AddVertexAt(9, new Point2d(scale1 * 0.805742597062001, scale1 * 0.278275990270772), 0, 0, 0);
            poly91.AddVertexAt(10, new Point2d(scale1 * 0.802703478939834, scale1 * 0.28037999512458), 0, 0, 0);
            poly91.AddVertexAt(11, new Point2d(scale1 * 0.799898139134758, scale1 * 0.282764533958894), 0, 0, 0);
            poly91.AddVertexAt(12, new Point2d(scale1 * 0.79756035596386, scale1 * 0.286271208715241), 0, 0, 0);
            poly91.AddVertexAt(13, new Point2d(scale1 * 0.795690129427142, scale1 * 0.288796014539811), 0, 0, 0);
            poly91.AddVertexAt(14, new Point2d(scale1 * 0.791482119719527, scale1 * 0.289777883471588), 0, 0, 0);
            poly91.AddVertexAt(15, new Point2d(scale1 * 0.787975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly91.Closed = true;
            poly91.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly91.Layer = "0";
            poly91.Color = color_GP;
            poly91.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly91);
            Polyline poly92 = new Polyline();
            poly92.AddVertexAt(0, new Point2d(scale1 * 0.758367272636016, scale1 * 0.274437775355651), 0, 0, 0);
            poly92.AddVertexAt(1, new Point2d(scale1 * 0.766081957099978, scale1 * 0.272754571472604), 0, 0, 0);
            poly92.AddVertexAt(2, new Point2d(scale1 * 0.769354853539234, scale1 * 0.268125760794228), 0, 0, 0);
            poly92.AddVertexAt(3, new Point2d(scale1 * 0.766081957099978, scale1 * 0.263637217106105), 0, 0, 0);
            poly92.AddVertexAt(4, new Point2d(scale1 * 0.758367272636016, scale1 * 0.261954013223059), 0, 0, 0);
            poly92.AddVertexAt(5, new Point2d(scale1 * 0.749717474903696, scale1 * 0.261392945262043), 0, 0, 0);
            poly92.AddVertexAt(6, new Point2d(scale1 * 0.745041908561901, scale1 * 0.26237481419382), 0, 0, 0);
            poly92.AddVertexAt(7, new Point2d(scale1 * 0.741301455488465, scale1 * 0.26489962001839), 0, 0, 0);
            poly92.AddVertexAt(8, new Point2d(scale1 * 0.730781431219427, scale1 * 0.268967362735751), 0, 0, 0);
            poly92.AddVertexAt(9, new Point2d(scale1 * 0.725170751609273, scale1 * 0.272894838462859), 0, 0, 0);
            poly92.AddVertexAt(10, new Point2d(scale1 * 0.726573421511812, scale1 * 0.27471830933616), 0, 0, 0);
            poly92.AddVertexAt(11, new Point2d(scale1 * 0.734989440927042, scale1 * 0.278365251082757), 0, 0, 0);
            poly92.AddVertexAt(12, new Point2d(scale1 * 0.741535233805555, scale1 * 0.279066586034028), 0, 0, 0);
            poly92.AddVertexAt(13, new Point2d(scale1 * 0.74574324351317, scale1 * 0.277804183121743), 0, 0, 0);
            poly92.AddVertexAt(14, new Point2d(scale1 * 0.751353923123324, scale1 * 0.276541780209458), 0, 0, 0);
            poly92.AddVertexAt(15, new Point2d(scale1 * 0.755795711148029, scale1 * 0.275840445258188), 0, 0, 0);
            poly92.Closed = true;
            poly92.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly92.Layer = "0";
            poly92.Color = color_GP;
            poly92.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly92);
            Polyline poly93 = new Polyline();
            poly93.AddVertexAt(0, new Point2d(scale1 * 0.778898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly93.AddVertexAt(1, new Point2d(scale1 * 0.772741026607183, scale1 * 0.262017770945895), 0, 0, 0);
            poly93.AddVertexAt(2, new Point2d(scale1 * 0.767432516861809, scale1 * 0.259722492923566), 0, 0, 0);
            poly93.AddVertexAt(3, new Point2d(scale1 * 0.766370814912736, scale1 * 0.256662122227115), 0, 0, 0);
            poly93.AddVertexAt(4, new Point2d(scale1 * 0.77167932465811, scale1 * 0.255769514107321), 0, 0, 0);
            poly93.AddVertexAt(5, new Point2d(scale1 * 0.778261876742373, scale1 * 0.2567896376728), 0, 0, 0);
            poly93.AddVertexAt(6, new Point2d(scale1 * 0.782084003759042, scale1 * 0.259467462032192), 0, 0, 0);
            poly93.AddVertexAt(7, new Point2d(scale1 * 0.778898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly93.Closed = true;
            poly93.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly93.Layer = "0";
            poly93.Color = color_GP;
            poly93.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly93);
            Polyline poly94 = new Polyline();
            poly94.AddVertexAt(0, new Point2d(scale1 * 0.796473138321644, scale1 * 0.248628649148944), 0, 0, 0);
            poly94.AddVertexAt(1, new Point2d(scale1 * 0.796706916638733, scale1 * 0.243298503519296), 0, 0, 0);
            poly94.AddVertexAt(2, new Point2d(scale1 * 0.795070468419106, scale1 * 0.242036100607014), 0, 0, 0);
            poly94.AddVertexAt(3, new Point2d(scale1 * 0.79086245871149, scale1 * 0.240633430704475), 0, 0, 0);
            poly94.AddVertexAt(4, new Point2d(scale1 * 0.786888227320964, scale1 * 0.240773697694727), 0, 0, 0);
            poly94.AddVertexAt(5, new Point2d(scale1 * 0.783849109198798, scale1 * 0.242036100607014), 0, 0, 0);
            poly94.AddVertexAt(6, new Point2d(scale1 * 0.780576212759542, scale1 * 0.243298503519296), 0, 0, 0);
            poly94.AddVertexAt(7, new Point2d(scale1 * 0.777537094637375, scale1 * 0.244560906431583), 0, 0, 0);
            poly94.AddVertexAt(8, new Point2d(scale1 * 0.775666868100657, scale1 * 0.246945445265898), 0, 0, 0);
            poly94.AddVertexAt(9, new Point2d(scale1 * 0.775666868100657, scale1 * 0.248628649148944), 0, 0, 0);
            poly94.AddVertexAt(10, new Point2d(scale1 * 0.779173542857003, scale1 * 0.252275590895543), 0, 0, 0);
            poly94.AddVertexAt(11, new Point2d(scale1 * 0.785953114052606, scale1 * 0.253818527788335), 0, 0, 0);
            poly94.AddVertexAt(12, new Point2d(scale1 * 0.792732685248208, scale1 * 0.252275590895543), 0, 0, 0);
            poly94.Closed = true;
            poly94.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly94.Layer = "0";
            poly94.Color = color_GP;
            poly94.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly94);
            Polyline poly95 = new Polyline();
            poly95.AddVertexAt(0, new Point2d(scale1 * 0.762107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly95.AddVertexAt(1, new Point2d(scale1 * 0.759302385904375, scale1 * 0.241895833616759), 0, 0, 0);
            poly95.AddVertexAt(2, new Point2d(scale1 * 0.760237499172734, scale1 * 0.240072362743458), 0, 0, 0);
            poly95.AddVertexAt(3, new Point2d(scale1 * 0.761640169075273, scale1 * 0.237968357889651), 0, 0, 0);
            poly95.AddVertexAt(4, new Point2d(scale1 * 0.764679287197439, scale1 * 0.237687823909144), 0, 0, 0);
            poly95.AddVertexAt(5, new Point2d(scale1 * 0.768185961953785, scale1 * 0.23937102779219), 0, 0, 0);
            poly95.AddVertexAt(6, new Point2d(scale1 * 0.768185961953785, scale1 * 0.241615299636252), 0, 0, 0);
            poly95.AddVertexAt(7, new Point2d(scale1 * 0.765848178782888, scale1 * 0.244140105460821), 0, 0, 0);
            poly95.AddVertexAt(8, new Point2d(scale1 * 0.762107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly95.Closed = true;
            poly95.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly95.Layer = "0";
            poly95.Color = color_GP;
            poly95.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly95);
            Polyline poly96 = new Polyline();
            poly96.AddVertexAt(0, new Point2d(scale1 * 0.716520953876953, scale1 * 0.248628649148944), 0, 0, 0);
            poly96.AddVertexAt(1, new Point2d(scale1 * 0.723066746755466, scale1 * 0.250452120022245), 0, 0, 0);
            poly96.AddVertexAt(2, new Point2d(scale1 * 0.735924554195401, scale1 * 0.249750785070974), 0, 0, 0);
            poly96.AddVertexAt(3, new Point2d(scale1 * 0.742470347073914, scale1 * 0.248628649148944), 0, 0, 0);
            poly96.AddVertexAt(4, new Point2d(scale1 * 0.743639238659363, scale1 * 0.246805178275643), 0, 0, 0);
            poly96.AddVertexAt(5, new Point2d(scale1 * 0.743873016976452, scale1 * 0.244981707402343), 0, 0, 0);
            poly96.AddVertexAt(6, new Point2d(scale1 * 0.744106795293542, scale1 * 0.242737435558282), 0, 0, 0);
            poly96.AddVertexAt(7, new Point2d(scale1 * 0.74597702183026, scale1 * 0.240773697694727), 0, 0, 0);
            poly96.AddVertexAt(8, new Point2d(scale1 * 0.744574351927722, scale1 * 0.237407289928637), 0, 0, 0);
            poly96.AddVertexAt(9, new Point2d(scale1 * 0.739431228951748, scale1 * 0.235583819055336), 0, 0, 0);
            poly96.AddVertexAt(10, new Point2d(scale1 * 0.729378761316889, scale1 * 0.235303285074829), 0, 0, 0);
            poly96.AddVertexAt(11, new Point2d(scale1 * 0.724469416658004, scale1 * 0.236144887016352), 0, 0, 0);
            poly96.AddVertexAt(12, new Point2d(scale1 * 0.720962741901658, scale1 * 0.237828090899398), 0, 0, 0);
            poly96.AddVertexAt(13, new Point2d(scale1 * 0.717689845462402, scale1 * 0.241054231675236), 0, 0, 0);
            poly96.AddVertexAt(14, new Point2d(scale1 * 0.715819618925684, scale1 * 0.244420639441328), 0, 0, 0);
            poly96.Closed = true;
            poly96.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly96.Layer = "0";
            poly96.Color = color_GP;
            poly96.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly96);
            Polyline poly97 = new Polyline();
            poly97.AddVertexAt(0, new Point2d(scale1 * 0.660285741001707, scale1 * 0.3), 0, 0, 0);
            poly97.AddVertexAt(1, new Point2d(scale1 * 0.661640169075273, scale1 * 0.297968357889651), 0, 0, 0);
            poly97.AddVertexAt(2, new Point2d(scale1 * 0.664679287197439, scale1 * 0.297687823909144), 0, 0, 0);
            poly97.AddVertexAt(3, new Point2d(scale1 * 0.668185961953786, scale1 * 0.29937102779219), 0, 0, 0);
            poly97.AddVertexAt(4, new Point2d(scale1 * 0.668185961953786, scale1 * 0.3), 0, 0, 0);
            poly97.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly97.Layer = "0";
            poly97.Color = color_GP;
            poly97.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly97);
            Polyline poly98 = new Polyline();
            poly98.AddVertexAt(0, new Point2d(scale1 * 0.645654647790791, scale1 * 0.3), 0, 0, 0);
            poly98.AddVertexAt(1, new Point2d(scale1 * 0.644574351927722, scale1 * 0.297407289928637), 0, 0, 0);
            poly98.AddVertexAt(2, new Point2d(scale1 * 0.639431228951748, scale1 * 0.295583819055336), 0, 0, 0);
            poly98.AddVertexAt(3, new Point2d(scale1 * 0.629378761316889, scale1 * 0.295303285074829), 0, 0, 0);
            poly98.AddVertexAt(4, new Point2d(scale1 * 0.624469416658004, scale1 * 0.296144887016352), 0, 0, 0);
            poly98.AddVertexAt(5, new Point2d(scale1 * 0.620962741901658, scale1 * 0.297828090899398), 0, 0, 0);
            poly98.AddVertexAt(6, new Point2d(scale1 * 0.618759355857569, scale1 * 0.3), 0, 0, 0);
            poly98.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly98.Layer = "0";
            poly98.Color = color_GP;
            poly98.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly98);
            Polyline poly99 = new Polyline();
            poly99.AddVertexAt(0, new Point2d(scale1 * 0.658367272636016, scale1 * 0.282994061761134), 0, 0, 0);
            poly99.AddVertexAt(1, new Point2d(scale1 * 0.656497046099298, scale1 * 0.282853794770882), 0, 0, 0);
            poly99.AddVertexAt(2, new Point2d(scale1 * 0.654159262928401, scale1 * 0.28355512972215), 0, 0, 0);
            poly99.AddVertexAt(3, new Point2d(scale1 * 0.652289036391683, scale1 * 0.285518867585705), 0, 0, 0);
            poly99.AddVertexAt(4, new Point2d(scale1 * 0.653925484611311, scale1 * 0.287061804478495), 0, 0, 0);
            poly99.AddVertexAt(5, new Point2d(scale1 * 0.658367272636016, scale1 * 0.286220202536974), 0, 0, 0);
            poly99.AddVertexAt(6, new Point2d(scale1 * 0.659536164221465, scale1 * 0.284817532634435), 0, 0, 0);
            poly99.Closed = true;
            poly99.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly99.Layer = "0";
            poly99.Color = color_GP;
            poly99.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly99);
            Polyline poly100 = new Polyline();
            poly100.AddVertexAt(0, new Point2d(scale1 * 0.667952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly100.AddVertexAt(1, new Point2d(scale1 * 0.666081957099978, scale1 * 0.282012192829359), 0, 0, 0);
            poly100.AddVertexAt(2, new Point2d(scale1 * 0.665848178782888, scale1 * 0.280048454965803), 0, 0, 0);
            poly100.AddVertexAt(3, new Point2d(scale1 * 0.668419740270875, scale1 * 0.278505518073012), 0, 0, 0);
            poly100.AddVertexAt(4, new Point2d(scale1 * 0.671692636710131, scale1 * 0.278926319043773), 0, 0, 0);
            poly100.AddVertexAt(5, new Point2d(scale1 * 0.672393971661401, scale1 * 0.281030323897581), 0, 0, 0);
            poly100.AddVertexAt(6, new Point2d(scale1 * 0.671926415027221, scale1 * 0.282713527780627), 0, 0, 0);
            poly100.AddVertexAt(7, new Point2d(scale1 * 0.667952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly100.Closed = true;
            poly100.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly100.Layer = "0";
            poly100.Color = color_GP;
            poly100.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly100);
            Polyline poly101 = new Polyline();
            poly101.AddVertexAt(0, new Point2d(scale1 * 0.616577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly101.AddVertexAt(1, new Point2d(scale1 * 0.614239328956148, scale1 * 0.285251085149766), 0, 0, 0);
            poly101.AddVertexAt(2, new Point2d(scale1 * 0.615641998858686, scale1 * 0.283427614276465), 0, 0, 0);
            poly101.AddVertexAt(3, new Point2d(scale1 * 0.619616230249212, scale1 * 0.282726279325197), 0, 0, 0);
            poly101.AddVertexAt(4, new Point2d(scale1 * 0.622655348371378, scale1 * 0.283988682237481), 0, 0, 0);
            poly101.AddVertexAt(5, new Point2d(scale1 * 0.622655348371378, scale1 * 0.286653755052303), 0, 0, 0);
            poly101.AddVertexAt(6, new Point2d(scale1 * 0.616577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly101.Closed = true;
            poly101.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly101.Layer = "0";
            poly101.Color = color_GP;
            poly101.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly101);
            Polyline poly102 = new Polyline();
            poly102.AddVertexAt(0, new Point2d(scale1 * 0.611080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly102.AddVertexAt(1, new Point2d(scale1 * 0.609476961807309, scale1 * 0.263254670769046), 0, 0, 0);
            poly102.AddVertexAt(2, new Point2d(scale1 * 0.609710740124399, scale1 * 0.259607729022446), 0, 0, 0);
            poly102.AddVertexAt(3, new Point2d(scale1 * 0.610645853392758, scale1 * 0.256802389217368), 0, 0, 0);
            poly102.AddVertexAt(4, new Point2d(scale1 * 0.614853863100373, scale1 * 0.255960787275844), 0, 0, 0);
            poly102.AddVertexAt(5, new Point2d(scale1 * 0.619061872807989, scale1 * 0.256521855236861), 0, 0, 0);
            poly102.AddVertexAt(6, new Point2d(scale1 * 0.622568547564335, scale1 * 0.258625860090668), 0, 0, 0);
            poly102.AddVertexAt(7, new Point2d(scale1 * 0.619061872807989, scale1 * 0.263535204749552), 0, 0, 0);
            poly102.AddVertexAt(8, new Point2d(scale1 * 0.611080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly102.Closed = true;
            poly102.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly102.Layer = "0";
            poly102.Color = color_GP;
            poly102.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly102);
            Polyline poly103 = new Polyline();
            poly103.AddVertexAt(0, new Point2d(scale1 * 0.687975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly103.AddVertexAt(1, new Point2d(scale1 * 0.684702548523924, scale1 * 0.288936281530065), 0, 0, 0);
            poly103.AddVertexAt(2, new Point2d(scale1 * 0.682598543670117, scale1 * 0.283746402890672), 0, 0, 0);
            poly103.AddVertexAt(3, new Point2d(scale1 * 0.682598543670117, scale1 * 0.281221597066103), 0, 0, 0);
            poly103.AddVertexAt(4, new Point2d(scale1 * 0.681897208718848, scale1 * 0.275751184446203), 0, 0, 0);
            poly103.AddVertexAt(5, new Point2d(scale1 * 0.685403883475194, scale1 * 0.273366645611888), 0, 0, 0);
            poly103.AddVertexAt(6, new Point2d(scale1 * 0.690313228134079, scale1 * 0.272525043670365), 0, 0, 0);
            poly103.AddVertexAt(7, new Point2d(scale1 * 0.701768365671475, scale1 * 0.272945844641126), 0, 0, 0);
            poly103.AddVertexAt(8, new Point2d(scale1 * 0.704807483793642, scale1 * 0.274488781533918), 0, 0, 0);
            poly103.AddVertexAt(9, new Point2d(scale1 * 0.705742597062001, scale1 * 0.278275990270772), 0, 0, 0);
            poly103.AddVertexAt(10, new Point2d(scale1 * 0.702703478939834, scale1 * 0.28037999512458), 0, 0, 0);
            poly103.AddVertexAt(11, new Point2d(scale1 * 0.699898139134758, scale1 * 0.282764533958894), 0, 0, 0);
            poly103.AddVertexAt(12, new Point2d(scale1 * 0.69756035596386, scale1 * 0.286271208715241), 0, 0, 0);
            poly103.AddVertexAt(13, new Point2d(scale1 * 0.695690129427142, scale1 * 0.288796014539811), 0, 0, 0);
            poly103.AddVertexAt(14, new Point2d(scale1 * 0.691482119719527, scale1 * 0.289777883471588), 0, 0, 0);
            poly103.AddVertexAt(15, new Point2d(scale1 * 0.687975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly103.Closed = true;
            poly103.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly103.Layer = "0";
            poly103.Color = color_GP;
            poly103.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly103);
            Polyline poly104 = new Polyline();
            poly104.AddVertexAt(0, new Point2d(scale1 * 0.658367272636016, scale1 * 0.274437775355651), 0, 0, 0);
            poly104.AddVertexAt(1, new Point2d(scale1 * 0.666081957099978, scale1 * 0.272754571472604), 0, 0, 0);
            poly104.AddVertexAt(2, new Point2d(scale1 * 0.669354853539234, scale1 * 0.268125760794228), 0, 0, 0);
            poly104.AddVertexAt(3, new Point2d(scale1 * 0.666081957099978, scale1 * 0.263637217106105), 0, 0, 0);
            poly104.AddVertexAt(4, new Point2d(scale1 * 0.658367272636016, scale1 * 0.261954013223059), 0, 0, 0);
            poly104.AddVertexAt(5, new Point2d(scale1 * 0.649717474903696, scale1 * 0.261392945262043), 0, 0, 0);
            poly104.AddVertexAt(6, new Point2d(scale1 * 0.645041908561901, scale1 * 0.26237481419382), 0, 0, 0);
            poly104.AddVertexAt(7, new Point2d(scale1 * 0.641301455488465, scale1 * 0.26489962001839), 0, 0, 0);
            poly104.AddVertexAt(8, new Point2d(scale1 * 0.630781431219427, scale1 * 0.268967362735751), 0, 0, 0);
            poly104.AddVertexAt(9, new Point2d(scale1 * 0.625170751609273, scale1 * 0.272894838462859), 0, 0, 0);
            poly104.AddVertexAt(10, new Point2d(scale1 * 0.626573421511812, scale1 * 0.27471830933616), 0, 0, 0);
            poly104.AddVertexAt(11, new Point2d(scale1 * 0.634989440927042, scale1 * 0.278365251082757), 0, 0, 0);
            poly104.AddVertexAt(12, new Point2d(scale1 * 0.641535233805555, scale1 * 0.279066586034028), 0, 0, 0);
            poly104.AddVertexAt(13, new Point2d(scale1 * 0.64574324351317, scale1 * 0.277804183121743), 0, 0, 0);
            poly104.AddVertexAt(14, new Point2d(scale1 * 0.651353923123324, scale1 * 0.276541780209458), 0, 0, 0);
            poly104.AddVertexAt(15, new Point2d(scale1 * 0.655795711148029, scale1 * 0.275840445258188), 0, 0, 0);
            poly104.Closed = true;
            poly104.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly104.Layer = "0";
            poly104.Color = color_GP;
            poly104.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly104);
            Polyline poly105 = new Polyline();
            poly105.AddVertexAt(0, new Point2d(scale1 * 0.678898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly105.AddVertexAt(1, new Point2d(scale1 * 0.672741026607184, scale1 * 0.262017770945895), 0, 0, 0);
            poly105.AddVertexAt(2, new Point2d(scale1 * 0.66743251686181, scale1 * 0.259722492923566), 0, 0, 0);
            poly105.AddVertexAt(3, new Point2d(scale1 * 0.666370814912736, scale1 * 0.256662122227115), 0, 0, 0);
            poly105.AddVertexAt(4, new Point2d(scale1 * 0.67167932465811, scale1 * 0.255769514107321), 0, 0, 0);
            poly105.AddVertexAt(5, new Point2d(scale1 * 0.678261876742373, scale1 * 0.2567896376728), 0, 0, 0);
            poly105.AddVertexAt(6, new Point2d(scale1 * 0.682084003759042, scale1 * 0.259467462032192), 0, 0, 0);
            poly105.AddVertexAt(7, new Point2d(scale1 * 0.678898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly105.Closed = true;
            poly105.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly105.Layer = "0";
            poly105.Color = color_GP;
            poly105.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly105);
            Polyline poly106 = new Polyline();
            poly106.AddVertexAt(0, new Point2d(scale1 * 0.696473138321644, scale1 * 0.248628649148944), 0, 0, 0);
            poly106.AddVertexAt(1, new Point2d(scale1 * 0.696706916638733, scale1 * 0.243298503519296), 0, 0, 0);
            poly106.AddVertexAt(2, new Point2d(scale1 * 0.695070468419106, scale1 * 0.242036100607014), 0, 0, 0);
            poly106.AddVertexAt(3, new Point2d(scale1 * 0.69086245871149, scale1 * 0.240633430704475), 0, 0, 0);
            poly106.AddVertexAt(4, new Point2d(scale1 * 0.686888227320964, scale1 * 0.240773697694727), 0, 0, 0);
            poly106.AddVertexAt(5, new Point2d(scale1 * 0.683849109198798, scale1 * 0.242036100607014), 0, 0, 0);
            poly106.AddVertexAt(6, new Point2d(scale1 * 0.680576212759542, scale1 * 0.243298503519296), 0, 0, 0);
            poly106.AddVertexAt(7, new Point2d(scale1 * 0.677537094637375, scale1 * 0.244560906431583), 0, 0, 0);
            poly106.AddVertexAt(8, new Point2d(scale1 * 0.675666868100657, scale1 * 0.246945445265898), 0, 0, 0);
            poly106.AddVertexAt(9, new Point2d(scale1 * 0.675666868100657, scale1 * 0.248628649148944), 0, 0, 0);
            poly106.AddVertexAt(10, new Point2d(scale1 * 0.679173542857003, scale1 * 0.252275590895543), 0, 0, 0);
            poly106.AddVertexAt(11, new Point2d(scale1 * 0.685953114052606, scale1 * 0.253818527788335), 0, 0, 0);
            poly106.AddVertexAt(12, new Point2d(scale1 * 0.692732685248208, scale1 * 0.252275590895543), 0, 0, 0);
            poly106.Closed = true;
            poly106.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly106.Layer = "0";
            poly106.Color = color_GP;
            poly106.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly106);
            Polyline poly107 = new Polyline();
            poly107.AddVertexAt(0, new Point2d(scale1 * 0.662107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly107.AddVertexAt(1, new Point2d(scale1 * 0.659302385904375, scale1 * 0.241895833616759), 0, 0, 0);
            poly107.AddVertexAt(2, new Point2d(scale1 * 0.660237499172734, scale1 * 0.240072362743458), 0, 0, 0);
            poly107.AddVertexAt(3, new Point2d(scale1 * 0.661640169075273, scale1 * 0.237968357889651), 0, 0, 0);
            poly107.AddVertexAt(4, new Point2d(scale1 * 0.664679287197439, scale1 * 0.237687823909144), 0, 0, 0);
            poly107.AddVertexAt(5, new Point2d(scale1 * 0.668185961953786, scale1 * 0.23937102779219), 0, 0, 0);
            poly107.AddVertexAt(6, new Point2d(scale1 * 0.668185961953786, scale1 * 0.241615299636252), 0, 0, 0);
            poly107.AddVertexAt(7, new Point2d(scale1 * 0.665848178782888, scale1 * 0.244140105460821), 0, 0, 0);
            poly107.AddVertexAt(8, new Point2d(scale1 * 0.662107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly107.Closed = true;
            poly107.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly107.Layer = "0";
            poly107.Color = color_GP;
            poly107.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly107);
            Polyline poly108 = new Polyline();
            poly108.AddVertexAt(0, new Point2d(scale1 * 0.616520953876953, scale1 * 0.248628649148944), 0, 0, 0);
            poly108.AddVertexAt(1, new Point2d(scale1 * 0.623066746755466, scale1 * 0.250452120022245), 0, 0, 0);
            poly108.AddVertexAt(2, new Point2d(scale1 * 0.635924554195402, scale1 * 0.249750785070974), 0, 0, 0);
            poly108.AddVertexAt(3, new Point2d(scale1 * 0.642470347073914, scale1 * 0.248628649148944), 0, 0, 0);
            poly108.AddVertexAt(4, new Point2d(scale1 * 0.643639238659363, scale1 * 0.246805178275643), 0, 0, 0);
            poly108.AddVertexAt(5, new Point2d(scale1 * 0.643873016976453, scale1 * 0.244981707402343), 0, 0, 0);
            poly108.AddVertexAt(6, new Point2d(scale1 * 0.644106795293542, scale1 * 0.242737435558282), 0, 0, 0);
            poly108.AddVertexAt(7, new Point2d(scale1 * 0.64597702183026, scale1 * 0.240773697694727), 0, 0, 0);
            poly108.AddVertexAt(8, new Point2d(scale1 * 0.644574351927722, scale1 * 0.237407289928637), 0, 0, 0);
            poly108.AddVertexAt(9, new Point2d(scale1 * 0.639431228951748, scale1 * 0.235583819055336), 0, 0, 0);
            poly108.AddVertexAt(10, new Point2d(scale1 * 0.629378761316889, scale1 * 0.235303285074829), 0, 0, 0);
            poly108.AddVertexAt(11, new Point2d(scale1 * 0.624469416658004, scale1 * 0.236144887016352), 0, 0, 0);
            poly108.AddVertexAt(12, new Point2d(scale1 * 0.620962741901658, scale1 * 0.237828090899398), 0, 0, 0);
            poly108.AddVertexAt(13, new Point2d(scale1 * 0.617689845462402, scale1 * 0.241054231675236), 0, 0, 0);
            poly108.AddVertexAt(14, new Point2d(scale1 * 0.615819618925684, scale1 * 0.244420639441328), 0, 0, 0);
            poly108.Closed = true;
            poly108.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly108.Layer = "0";
            poly108.Color = color_GP;
            poly108.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly108);
            Polyline poly109 = new Polyline();
            poly109.AddVertexAt(0, new Point2d(scale1 * 0.560285741001707, scale1 * 0.3), 0, 0, 0);
            poly109.AddVertexAt(1, new Point2d(scale1 * 0.561640169075273, scale1 * 0.297968357889651), 0, 0, 0);
            poly109.AddVertexAt(2, new Point2d(scale1 * 0.564679287197439, scale1 * 0.297687823909144), 0, 0, 0);
            poly109.AddVertexAt(3, new Point2d(scale1 * 0.568185961953785, scale1 * 0.29937102779219), 0, 0, 0);
            poly109.AddVertexAt(4, new Point2d(scale1 * 0.568185961953785, scale1 * 0.3), 0, 0, 0);
            poly109.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly109.Layer = "0";
            poly109.Color = color_GP;
            poly109.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly109);
            Polyline poly110 = new Polyline();
            poly110.AddVertexAt(0, new Point2d(scale1 * 0.54565464779079, scale1 * 0.3), 0, 0, 0);
            poly110.AddVertexAt(1, new Point2d(scale1 * 0.544574351927722, scale1 * 0.297407289928637), 0, 0, 0);
            poly110.AddVertexAt(2, new Point2d(scale1 * 0.539431228951748, scale1 * 0.295583819055336), 0, 0, 0);
            poly110.AddVertexAt(3, new Point2d(scale1 * 0.529378761316889, scale1 * 0.295303285074829), 0, 0, 0);
            poly110.AddVertexAt(4, new Point2d(scale1 * 0.524469416658004, scale1 * 0.296144887016352), 0, 0, 0);
            poly110.AddVertexAt(5, new Point2d(scale1 * 0.520962741901658, scale1 * 0.297828090899398), 0, 0, 0);
            poly110.AddVertexAt(6, new Point2d(scale1 * 0.518759355857569, scale1 * 0.3), 0, 0, 0);
            poly110.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly110.Layer = "0";
            poly110.Color = color_GP;
            poly110.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly110);
            Polyline poly111 = new Polyline();
            poly111.AddVertexAt(0, new Point2d(scale1 * 0.558367272636016, scale1 * 0.282994061761134), 0, 0, 0);
            poly111.AddVertexAt(1, new Point2d(scale1 * 0.556497046099298, scale1 * 0.282853794770882), 0, 0, 0);
            poly111.AddVertexAt(2, new Point2d(scale1 * 0.554159262928401, scale1 * 0.28355512972215), 0, 0, 0);
            poly111.AddVertexAt(3, new Point2d(scale1 * 0.552289036391683, scale1 * 0.285518867585705), 0, 0, 0);
            poly111.AddVertexAt(4, new Point2d(scale1 * 0.553925484611311, scale1 * 0.287061804478495), 0, 0, 0);
            poly111.AddVertexAt(5, new Point2d(scale1 * 0.558367272636016, scale1 * 0.286220202536974), 0, 0, 0);
            poly111.AddVertexAt(6, new Point2d(scale1 * 0.559536164221465, scale1 * 0.284817532634435), 0, 0, 0);
            poly111.Closed = true;
            poly111.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly111.Layer = "0";
            poly111.Color = color_GP;
            poly111.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly111);
            Polyline poly112 = new Polyline();
            poly112.AddVertexAt(0, new Point2d(scale1 * 0.567952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly112.AddVertexAt(1, new Point2d(scale1 * 0.566081957099978, scale1 * 0.282012192829359), 0, 0, 0);
            poly112.AddVertexAt(2, new Point2d(scale1 * 0.565848178782888, scale1 * 0.280048454965803), 0, 0, 0);
            poly112.AddVertexAt(3, new Point2d(scale1 * 0.568419740270875, scale1 * 0.278505518073012), 0, 0, 0);
            poly112.AddVertexAt(4, new Point2d(scale1 * 0.571692636710131, scale1 * 0.278926319043773), 0, 0, 0);
            poly112.AddVertexAt(5, new Point2d(scale1 * 0.572393971661401, scale1 * 0.281030323897581), 0, 0, 0);
            poly112.AddVertexAt(6, new Point2d(scale1 * 0.571926415027221, scale1 * 0.282713527780627), 0, 0, 0);
            poly112.AddVertexAt(7, new Point2d(scale1 * 0.567952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly112.Closed = true;
            poly112.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly112.Layer = "0";
            poly112.Color = color_GP;
            poly112.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly112);
            Polyline poly113 = new Polyline();
            poly113.AddVertexAt(0, new Point2d(scale1 * 0.516577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly113.AddVertexAt(1, new Point2d(scale1 * 0.514239328956148, scale1 * 0.285251085149766), 0, 0, 0);
            poly113.AddVertexAt(2, new Point2d(scale1 * 0.515641998858686, scale1 * 0.283427614276465), 0, 0, 0);
            poly113.AddVertexAt(3, new Point2d(scale1 * 0.519616230249212, scale1 * 0.282726279325197), 0, 0, 0);
            poly113.AddVertexAt(4, new Point2d(scale1 * 0.522655348371378, scale1 * 0.283988682237481), 0, 0, 0);
            poly113.AddVertexAt(5, new Point2d(scale1 * 0.522655348371378, scale1 * 0.286653755052303), 0, 0, 0);
            poly113.AddVertexAt(6, new Point2d(scale1 * 0.516577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly113.Closed = true;
            poly113.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly113.Layer = "0";
            poly113.Color = color_GP;
            poly113.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly113);
            Polyline poly114 = new Polyline();
            poly114.AddVertexAt(0, new Point2d(scale1 * 0.511080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly114.AddVertexAt(1, new Point2d(scale1 * 0.509476961807309, scale1 * 0.263254670769046), 0, 0, 0);
            poly114.AddVertexAt(2, new Point2d(scale1 * 0.509710740124399, scale1 * 0.259607729022446), 0, 0, 0);
            poly114.AddVertexAt(3, new Point2d(scale1 * 0.510645853392758, scale1 * 0.256802389217368), 0, 0, 0);
            poly114.AddVertexAt(4, new Point2d(scale1 * 0.514853863100373, scale1 * 0.255960787275844), 0, 0, 0);
            poly114.AddVertexAt(5, new Point2d(scale1 * 0.519061872807989, scale1 * 0.256521855236861), 0, 0, 0);
            poly114.AddVertexAt(6, new Point2d(scale1 * 0.522568547564335, scale1 * 0.258625860090668), 0, 0, 0);
            poly114.AddVertexAt(7, new Point2d(scale1 * 0.519061872807989, scale1 * 0.263535204749552), 0, 0, 0);
            poly114.AddVertexAt(8, new Point2d(scale1 * 0.511080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly114.Closed = true;
            poly114.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly114.Layer = "0";
            poly114.Color = color_GP;
            poly114.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly114);
            Polyline poly115 = new Polyline();
            poly115.AddVertexAt(0, new Point2d(scale1 * 0.587975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly115.AddVertexAt(1, new Point2d(scale1 * 0.584702548523924, scale1 * 0.288936281530065), 0, 0, 0);
            poly115.AddVertexAt(2, new Point2d(scale1 * 0.582598543670117, scale1 * 0.283746402890672), 0, 0, 0);
            poly115.AddVertexAt(3, new Point2d(scale1 * 0.582598543670117, scale1 * 0.281221597066103), 0, 0, 0);
            poly115.AddVertexAt(4, new Point2d(scale1 * 0.581897208718847, scale1 * 0.275751184446203), 0, 0, 0);
            poly115.AddVertexAt(5, new Point2d(scale1 * 0.585403883475194, scale1 * 0.273366645611888), 0, 0, 0);
            poly115.AddVertexAt(6, new Point2d(scale1 * 0.590313228134079, scale1 * 0.272525043670365), 0, 0, 0);
            poly115.AddVertexAt(7, new Point2d(scale1 * 0.601768365671475, scale1 * 0.272945844641126), 0, 0, 0);
            poly115.AddVertexAt(8, new Point2d(scale1 * 0.604807483793642, scale1 * 0.274488781533918), 0, 0, 0);
            poly115.AddVertexAt(9, new Point2d(scale1 * 0.605742597062001, scale1 * 0.278275990270772), 0, 0, 0);
            poly115.AddVertexAt(10, new Point2d(scale1 * 0.602703478939834, scale1 * 0.28037999512458), 0, 0, 0);
            poly115.AddVertexAt(11, new Point2d(scale1 * 0.599898139134758, scale1 * 0.282764533958894), 0, 0, 0);
            poly115.AddVertexAt(12, new Point2d(scale1 * 0.59756035596386, scale1 * 0.286271208715241), 0, 0, 0);
            poly115.AddVertexAt(13, new Point2d(scale1 * 0.595690129427142, scale1 * 0.288796014539811), 0, 0, 0);
            poly115.AddVertexAt(14, new Point2d(scale1 * 0.591482119719527, scale1 * 0.289777883471588), 0, 0, 0);
            poly115.AddVertexAt(15, new Point2d(scale1 * 0.587975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly115.Closed = true;
            poly115.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly115.Layer = "0";
            poly115.Color = color_GP;
            poly115.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly115);
            Polyline poly116 = new Polyline();
            poly116.AddVertexAt(0, new Point2d(scale1 * 0.558367272636016, scale1 * 0.274437775355651), 0, 0, 0);
            poly116.AddVertexAt(1, new Point2d(scale1 * 0.566081957099978, scale1 * 0.272754571472604), 0, 0, 0);
            poly116.AddVertexAt(2, new Point2d(scale1 * 0.569354853539234, scale1 * 0.268125760794228), 0, 0, 0);
            poly116.AddVertexAt(3, new Point2d(scale1 * 0.566081957099978, scale1 * 0.263637217106105), 0, 0, 0);
            poly116.AddVertexAt(4, new Point2d(scale1 * 0.558367272636016, scale1 * 0.261954013223059), 0, 0, 0);
            poly116.AddVertexAt(5, new Point2d(scale1 * 0.549717474903696, scale1 * 0.261392945262043), 0, 0, 0);
            poly116.AddVertexAt(6, new Point2d(scale1 * 0.545041908561901, scale1 * 0.26237481419382), 0, 0, 0);
            poly116.AddVertexAt(7, new Point2d(scale1 * 0.541301455488465, scale1 * 0.26489962001839), 0, 0, 0);
            poly116.AddVertexAt(8, new Point2d(scale1 * 0.530781431219427, scale1 * 0.268967362735751), 0, 0, 0);
            poly116.AddVertexAt(9, new Point2d(scale1 * 0.525170751609273, scale1 * 0.272894838462859), 0, 0, 0);
            poly116.AddVertexAt(10, new Point2d(scale1 * 0.526573421511812, scale1 * 0.27471830933616), 0, 0, 0);
            poly116.AddVertexAt(11, new Point2d(scale1 * 0.534989440927042, scale1 * 0.278365251082757), 0, 0, 0);
            poly116.AddVertexAt(12, new Point2d(scale1 * 0.541535233805555, scale1 * 0.279066586034028), 0, 0, 0);
            poly116.AddVertexAt(13, new Point2d(scale1 * 0.54574324351317, scale1 * 0.277804183121743), 0, 0, 0);
            poly116.AddVertexAt(14, new Point2d(scale1 * 0.551353923123324, scale1 * 0.276541780209458), 0, 0, 0);
            poly116.AddVertexAt(15, new Point2d(scale1 * 0.555795711148029, scale1 * 0.275840445258188), 0, 0, 0);
            poly116.Closed = true;
            poly116.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly116.Layer = "0";
            poly116.Color = color_GP;
            poly116.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly116);
            Polyline poly117 = new Polyline();
            poly117.AddVertexAt(0, new Point2d(scale1 * 0.578898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly117.AddVertexAt(1, new Point2d(scale1 * 0.572741026607184, scale1 * 0.262017770945895), 0, 0, 0);
            poly117.AddVertexAt(2, new Point2d(scale1 * 0.56743251686181, scale1 * 0.259722492923566), 0, 0, 0);
            poly117.AddVertexAt(3, new Point2d(scale1 * 0.566370814912736, scale1 * 0.256662122227115), 0, 0, 0);
            poly117.AddVertexAt(4, new Point2d(scale1 * 0.57167932465811, scale1 * 0.255769514107321), 0, 0, 0);
            poly117.AddVertexAt(5, new Point2d(scale1 * 0.578261876742373, scale1 * 0.2567896376728), 0, 0, 0);
            poly117.AddVertexAt(6, new Point2d(scale1 * 0.582084003759042, scale1 * 0.259467462032192), 0, 0, 0);
            poly117.AddVertexAt(7, new Point2d(scale1 * 0.578898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly117.Closed = true;
            poly117.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly117.Layer = "0";
            poly117.Color = color_GP;
            poly117.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly117);
            Polyline poly118 = new Polyline();
            poly118.AddVertexAt(0, new Point2d(scale1 * 0.596473138321644, scale1 * 0.248628649148944), 0, 0, 0);
            poly118.AddVertexAt(1, new Point2d(scale1 * 0.596706916638733, scale1 * 0.243298503519296), 0, 0, 0);
            poly118.AddVertexAt(2, new Point2d(scale1 * 0.595070468419106, scale1 * 0.242036100607014), 0, 0, 0);
            poly118.AddVertexAt(3, new Point2d(scale1 * 0.59086245871149, scale1 * 0.240633430704475), 0, 0, 0);
            poly118.AddVertexAt(4, new Point2d(scale1 * 0.586888227320964, scale1 * 0.240773697694727), 0, 0, 0);
            poly118.AddVertexAt(5, new Point2d(scale1 * 0.583849109198798, scale1 * 0.242036100607014), 0, 0, 0);
            poly118.AddVertexAt(6, new Point2d(scale1 * 0.580576212759542, scale1 * 0.243298503519296), 0, 0, 0);
            poly118.AddVertexAt(7, new Point2d(scale1 * 0.577537094637375, scale1 * 0.244560906431583), 0, 0, 0);
            poly118.AddVertexAt(8, new Point2d(scale1 * 0.575666868100657, scale1 * 0.246945445265898), 0, 0, 0);
            poly118.AddVertexAt(9, new Point2d(scale1 * 0.575666868100657, scale1 * 0.248628649148944), 0, 0, 0);
            poly118.AddVertexAt(10, new Point2d(scale1 * 0.579173542857003, scale1 * 0.252275590895543), 0, 0, 0);
            poly118.AddVertexAt(11, new Point2d(scale1 * 0.585953114052606, scale1 * 0.253818527788335), 0, 0, 0);
            poly118.AddVertexAt(12, new Point2d(scale1 * 0.592732685248208, scale1 * 0.252275590895543), 0, 0, 0);
            poly118.Closed = true;
            poly118.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly118.Layer = "0";
            poly118.Color = color_GP;
            poly118.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly118);
            Polyline poly119 = new Polyline();
            poly119.AddVertexAt(0, new Point2d(scale1 * 0.562107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly119.AddVertexAt(1, new Point2d(scale1 * 0.559302385904375, scale1 * 0.241895833616759), 0, 0, 0);
            poly119.AddVertexAt(2, new Point2d(scale1 * 0.560237499172734, scale1 * 0.240072362743458), 0, 0, 0);
            poly119.AddVertexAt(3, new Point2d(scale1 * 0.561640169075273, scale1 * 0.237968357889651), 0, 0, 0);
            poly119.AddVertexAt(4, new Point2d(scale1 * 0.564679287197439, scale1 * 0.237687823909144), 0, 0, 0);
            poly119.AddVertexAt(5, new Point2d(scale1 * 0.568185961953785, scale1 * 0.23937102779219), 0, 0, 0);
            poly119.AddVertexAt(6, new Point2d(scale1 * 0.568185961953785, scale1 * 0.241615299636252), 0, 0, 0);
            poly119.AddVertexAt(7, new Point2d(scale1 * 0.565848178782888, scale1 * 0.244140105460821), 0, 0, 0);
            poly119.AddVertexAt(8, new Point2d(scale1 * 0.562107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly119.Closed = true;
            poly119.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly119.Layer = "0";
            poly119.Color = color_GP;
            poly119.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly119);
            Polyline poly120 = new Polyline();
            poly120.AddVertexAt(0, new Point2d(scale1 * 0.516520953876953, scale1 * 0.248628649148944), 0, 0, 0);
            poly120.AddVertexAt(1, new Point2d(scale1 * 0.523066746755466, scale1 * 0.250452120022245), 0, 0, 0);
            poly120.AddVertexAt(2, new Point2d(scale1 * 0.535924554195401, scale1 * 0.249750785070974), 0, 0, 0);
            poly120.AddVertexAt(3, new Point2d(scale1 * 0.542470347073914, scale1 * 0.248628649148944), 0, 0, 0);
            poly120.AddVertexAt(4, new Point2d(scale1 * 0.543639238659362, scale1 * 0.246805178275643), 0, 0, 0);
            poly120.AddVertexAt(5, new Point2d(scale1 * 0.543873016976453, scale1 * 0.244981707402343), 0, 0, 0);
            poly120.AddVertexAt(6, new Point2d(scale1 * 0.544106795293542, scale1 * 0.242737435558282), 0, 0, 0);
            poly120.AddVertexAt(7, new Point2d(scale1 * 0.54597702183026, scale1 * 0.240773697694727), 0, 0, 0);
            poly120.AddVertexAt(8, new Point2d(scale1 * 0.544574351927722, scale1 * 0.237407289928637), 0, 0, 0);
            poly120.AddVertexAt(9, new Point2d(scale1 * 0.539431228951748, scale1 * 0.235583819055336), 0, 0, 0);
            poly120.AddVertexAt(10, new Point2d(scale1 * 0.529378761316889, scale1 * 0.235303285074829), 0, 0, 0);
            poly120.AddVertexAt(11, new Point2d(scale1 * 0.524469416658004, scale1 * 0.236144887016352), 0, 0, 0);
            poly120.AddVertexAt(12, new Point2d(scale1 * 0.520962741901658, scale1 * 0.237828090899398), 0, 0, 0);
            poly120.AddVertexAt(13, new Point2d(scale1 * 0.517689845462401, scale1 * 0.241054231675236), 0, 0, 0);
            poly120.AddVertexAt(14, new Point2d(scale1 * 0.515819618925684, scale1 * 0.244420639441328), 0, 0, 0);
            poly120.Closed = true;
            poly120.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly120.Layer = "0";
            poly120.Color = color_GP;
            poly120.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly120);
            Polyline poly121 = new Polyline();
            poly121.AddVertexAt(0, new Point2d(scale1 * 0.460285741001707, scale1 * 0.3), 0, 0, 0);
            poly121.AddVertexAt(1, new Point2d(scale1 * 0.461640169075273, scale1 * 0.297968357889651), 0, 0, 0);
            poly121.AddVertexAt(2, new Point2d(scale1 * 0.464679287197439, scale1 * 0.297687823909144), 0, 0, 0);
            poly121.AddVertexAt(3, new Point2d(scale1 * 0.468185961953785, scale1 * 0.29937102779219), 0, 0, 0);
            poly121.AddVertexAt(4, new Point2d(scale1 * 0.468185961953785, scale1 * 0.3), 0, 0, 0);
            poly121.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly121.Layer = "0";
            poly121.Color = color_GP;
            poly121.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly121);
            Polyline poly122 = new Polyline();
            poly122.AddVertexAt(0, new Point2d(scale1 * 0.44565464779079, scale1 * 0.3), 0, 0, 0);
            poly122.AddVertexAt(1, new Point2d(scale1 * 0.444574351927722, scale1 * 0.297407289928637), 0, 0, 0);
            poly122.AddVertexAt(2, new Point2d(scale1 * 0.439431228951747, scale1 * 0.295583819055336), 0, 0, 0);
            poly122.AddVertexAt(3, new Point2d(scale1 * 0.429378761316889, scale1 * 0.295303285074829), 0, 0, 0);
            poly122.AddVertexAt(4, new Point2d(scale1 * 0.424469416658004, scale1 * 0.296144887016352), 0, 0, 0);
            poly122.AddVertexAt(5, new Point2d(scale1 * 0.420962741901658, scale1 * 0.297828090899398), 0, 0, 0);
            poly122.AddVertexAt(6, new Point2d(scale1 * 0.418759355857569, scale1 * 0.3), 0, 0, 0);
            poly122.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly122.Layer = "0";
            poly122.Color = color_GP;
            poly122.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly122);
            Polyline poly123 = new Polyline();
            poly123.AddVertexAt(0, new Point2d(scale1 * 0.458367272636016, scale1 * 0.282994061761134), 0, 0, 0);
            poly123.AddVertexAt(1, new Point2d(scale1 * 0.456497046099298, scale1 * 0.282853794770882), 0, 0, 0);
            poly123.AddVertexAt(2, new Point2d(scale1 * 0.454159262928401, scale1 * 0.28355512972215), 0, 0, 0);
            poly123.AddVertexAt(3, new Point2d(scale1 * 0.452289036391683, scale1 * 0.285518867585705), 0, 0, 0);
            poly123.AddVertexAt(4, new Point2d(scale1 * 0.453925484611311, scale1 * 0.287061804478495), 0, 0, 0);
            poly123.AddVertexAt(5, new Point2d(scale1 * 0.458367272636016, scale1 * 0.286220202536974), 0, 0, 0);
            poly123.AddVertexAt(6, new Point2d(scale1 * 0.459536164221465, scale1 * 0.284817532634435), 0, 0, 0);
            poly123.Closed = true;
            poly123.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly123.Layer = "0";
            poly123.Color = color_GP;
            poly123.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly123);
            Polyline poly124 = new Polyline();
            poly124.AddVertexAt(0, new Point2d(scale1 * 0.467952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly124.AddVertexAt(1, new Point2d(scale1 * 0.466081957099978, scale1 * 0.282012192829359), 0, 0, 0);
            poly124.AddVertexAt(2, new Point2d(scale1 * 0.465848178782888, scale1 * 0.280048454965803), 0, 0, 0);
            poly124.AddVertexAt(3, new Point2d(scale1 * 0.468419740270875, scale1 * 0.278505518073012), 0, 0, 0);
            poly124.AddVertexAt(4, new Point2d(scale1 * 0.471692636710132, scale1 * 0.278926319043773), 0, 0, 0);
            poly124.AddVertexAt(5, new Point2d(scale1 * 0.472393971661401, scale1 * 0.281030323897581), 0, 0, 0);
            poly124.AddVertexAt(6, new Point2d(scale1 * 0.471926415027221, scale1 * 0.282713527780627), 0, 0, 0);
            poly124.AddVertexAt(7, new Point2d(scale1 * 0.467952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly124.Closed = true;
            poly124.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly124.Layer = "0";
            poly124.Color = color_GP;
            poly124.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly124);
            Polyline poly125 = new Polyline();
            poly125.AddVertexAt(0, new Point2d(scale1 * 0.416577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly125.AddVertexAt(1, new Point2d(scale1 * 0.414239328956148, scale1 * 0.285251085149766), 0, 0, 0);
            poly125.AddVertexAt(2, new Point2d(scale1 * 0.415641998858686, scale1 * 0.283427614276465), 0, 0, 0);
            poly125.AddVertexAt(3, new Point2d(scale1 * 0.419616230249211, scale1 * 0.282726279325197), 0, 0, 0);
            poly125.AddVertexAt(4, new Point2d(scale1 * 0.422655348371378, scale1 * 0.283988682237481), 0, 0, 0);
            poly125.AddVertexAt(5, new Point2d(scale1 * 0.422655348371378, scale1 * 0.286653755052303), 0, 0, 0);
            poly125.AddVertexAt(6, new Point2d(scale1 * 0.416577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly125.Closed = true;
            poly125.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly125.Layer = "0";
            poly125.Color = color_GP;
            poly125.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly125);
            Polyline poly126 = new Polyline();
            poly126.AddVertexAt(0, new Point2d(scale1 * 0.411080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly126.AddVertexAt(1, new Point2d(scale1 * 0.409476961807309, scale1 * 0.263254670769046), 0, 0, 0);
            poly126.AddVertexAt(2, new Point2d(scale1 * 0.409710740124399, scale1 * 0.259607729022446), 0, 0, 0);
            poly126.AddVertexAt(3, new Point2d(scale1 * 0.410645853392758, scale1 * 0.256802389217368), 0, 0, 0);
            poly126.AddVertexAt(4, new Point2d(scale1 * 0.414853863100374, scale1 * 0.255960787275844), 0, 0, 0);
            poly126.AddVertexAt(5, new Point2d(scale1 * 0.419061872807989, scale1 * 0.256521855236861), 0, 0, 0);
            poly126.AddVertexAt(6, new Point2d(scale1 * 0.422568547564335, scale1 * 0.258625860090668), 0, 0, 0);
            poly126.AddVertexAt(7, new Point2d(scale1 * 0.419061872807989, scale1 * 0.263535204749552), 0, 0, 0);
            poly126.AddVertexAt(8, new Point2d(scale1 * 0.411080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly126.Closed = true;
            poly126.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly126.Layer = "0";
            poly126.Color = color_GP;
            poly126.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly126);
            Polyline poly127 = new Polyline();
            poly127.AddVertexAt(0, new Point2d(scale1 * 0.487975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly127.AddVertexAt(1, new Point2d(scale1 * 0.484702548523924, scale1 * 0.288936281530065), 0, 0, 0);
            poly127.AddVertexAt(2, new Point2d(scale1 * 0.482598543670117, scale1 * 0.283746402890672), 0, 0, 0);
            poly127.AddVertexAt(3, new Point2d(scale1 * 0.482598543670117, scale1 * 0.281221597066103), 0, 0, 0);
            poly127.AddVertexAt(4, new Point2d(scale1 * 0.481897208718848, scale1 * 0.275751184446203), 0, 0, 0);
            poly127.AddVertexAt(5, new Point2d(scale1 * 0.485403883475193, scale1 * 0.273366645611888), 0, 0, 0);
            poly127.AddVertexAt(6, new Point2d(scale1 * 0.490313228134079, scale1 * 0.272525043670365), 0, 0, 0);
            poly127.AddVertexAt(7, new Point2d(scale1 * 0.501768365671475, scale1 * 0.272945844641126), 0, 0, 0);
            poly127.AddVertexAt(8, new Point2d(scale1 * 0.504807483793642, scale1 * 0.274488781533918), 0, 0, 0);
            poly127.AddVertexAt(9, new Point2d(scale1 * 0.505742597062001, scale1 * 0.278275990270772), 0, 0, 0);
            poly127.AddVertexAt(10, new Point2d(scale1 * 0.502703478939834, scale1 * 0.28037999512458), 0, 0, 0);
            poly127.AddVertexAt(11, new Point2d(scale1 * 0.499898139134758, scale1 * 0.282764533958894), 0, 0, 0);
            poly127.AddVertexAt(12, new Point2d(scale1 * 0.49756035596386, scale1 * 0.286271208715241), 0, 0, 0);
            poly127.AddVertexAt(13, new Point2d(scale1 * 0.495690129427142, scale1 * 0.288796014539811), 0, 0, 0);
            poly127.AddVertexAt(14, new Point2d(scale1 * 0.491482119719527, scale1 * 0.289777883471588), 0, 0, 0);
            poly127.AddVertexAt(15, new Point2d(scale1 * 0.487975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly127.Closed = true;
            poly127.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly127.Layer = "0";
            poly127.Color = color_GP;
            poly127.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly127);
            Polyline poly128 = new Polyline();
            poly128.AddVertexAt(0, new Point2d(scale1 * 0.458367272636016, scale1 * 0.274437775355651), 0, 0, 0);
            poly128.AddVertexAt(1, new Point2d(scale1 * 0.466081957099978, scale1 * 0.272754571472604), 0, 0, 0);
            poly128.AddVertexAt(2, new Point2d(scale1 * 0.469354853539234, scale1 * 0.268125760794228), 0, 0, 0);
            poly128.AddVertexAt(3, new Point2d(scale1 * 0.466081957099978, scale1 * 0.263637217106105), 0, 0, 0);
            poly128.AddVertexAt(4, new Point2d(scale1 * 0.458367272636016, scale1 * 0.261954013223059), 0, 0, 0);
            poly128.AddVertexAt(5, new Point2d(scale1 * 0.449717474903696, scale1 * 0.261392945262043), 0, 0, 0);
            poly128.AddVertexAt(6, new Point2d(scale1 * 0.445041908561901, scale1 * 0.26237481419382), 0, 0, 0);
            poly128.AddVertexAt(7, new Point2d(scale1 * 0.441301455488465, scale1 * 0.26489962001839), 0, 0, 0);
            poly128.AddVertexAt(8, new Point2d(scale1 * 0.430781431219427, scale1 * 0.268967362735751), 0, 0, 0);
            poly128.AddVertexAt(9, new Point2d(scale1 * 0.425170751609273, scale1 * 0.272894838462859), 0, 0, 0);
            poly128.AddVertexAt(10, new Point2d(scale1 * 0.426573421511812, scale1 * 0.27471830933616), 0, 0, 0);
            poly128.AddVertexAt(11, new Point2d(scale1 * 0.434989440927043, scale1 * 0.278365251082757), 0, 0, 0);
            poly128.AddVertexAt(12, new Point2d(scale1 * 0.441535233805555, scale1 * 0.279066586034028), 0, 0, 0);
            poly128.AddVertexAt(13, new Point2d(scale1 * 0.445743243513171, scale1 * 0.277804183121743), 0, 0, 0);
            poly128.AddVertexAt(14, new Point2d(scale1 * 0.451353923123324, scale1 * 0.276541780209458), 0, 0, 0);
            poly128.AddVertexAt(15, new Point2d(scale1 * 0.455795711148029, scale1 * 0.275840445258188), 0, 0, 0);
            poly128.Closed = true;
            poly128.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly128.Layer = "0";
            poly128.Color = color_GP;
            poly128.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly128);
            Polyline poly129 = new Polyline();
            poly129.AddVertexAt(0, new Point2d(scale1 * 0.478898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly129.AddVertexAt(1, new Point2d(scale1 * 0.472741026607183, scale1 * 0.262017770945895), 0, 0, 0);
            poly129.AddVertexAt(2, new Point2d(scale1 * 0.46743251686181, scale1 * 0.259722492923566), 0, 0, 0);
            poly129.AddVertexAt(3, new Point2d(scale1 * 0.466370814912736, scale1 * 0.256662122227115), 0, 0, 0);
            poly129.AddVertexAt(4, new Point2d(scale1 * 0.47167932465811, scale1 * 0.255769514107321), 0, 0, 0);
            poly129.AddVertexAt(5, new Point2d(scale1 * 0.478261876742373, scale1 * 0.2567896376728), 0, 0, 0);
            poly129.AddVertexAt(6, new Point2d(scale1 * 0.482084003759041, scale1 * 0.259467462032192), 0, 0, 0);
            poly129.AddVertexAt(7, new Point2d(scale1 * 0.478898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly129.Closed = true;
            poly129.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly129.Layer = "0";
            poly129.Color = color_GP;
            poly129.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly129);
            Polyline poly130 = new Polyline();
            poly130.AddVertexAt(0, new Point2d(scale1 * 0.496473138321644, scale1 * 0.248628649148944), 0, 0, 0);
            poly130.AddVertexAt(1, new Point2d(scale1 * 0.496706916638733, scale1 * 0.243298503519296), 0, 0, 0);
            poly130.AddVertexAt(2, new Point2d(scale1 * 0.495070468419106, scale1 * 0.242036100607014), 0, 0, 0);
            poly130.AddVertexAt(3, new Point2d(scale1 * 0.49086245871149, scale1 * 0.240633430704475), 0, 0, 0);
            poly130.AddVertexAt(4, new Point2d(scale1 * 0.486888227320964, scale1 * 0.240773697694727), 0, 0, 0);
            poly130.AddVertexAt(5, new Point2d(scale1 * 0.483849109198798, scale1 * 0.242036100607014), 0, 0, 0);
            poly130.AddVertexAt(6, new Point2d(scale1 * 0.480576212759541, scale1 * 0.243298503519296), 0, 0, 0);
            poly130.AddVertexAt(7, new Point2d(scale1 * 0.477537094637375, scale1 * 0.244560906431583), 0, 0, 0);
            poly130.AddVertexAt(8, new Point2d(scale1 * 0.475666868100657, scale1 * 0.246945445265898), 0, 0, 0);
            poly130.AddVertexAt(9, new Point2d(scale1 * 0.475666868100657, scale1 * 0.248628649148944), 0, 0, 0);
            poly130.AddVertexAt(10, new Point2d(scale1 * 0.479173542857003, scale1 * 0.252275590895543), 0, 0, 0);
            poly130.AddVertexAt(11, new Point2d(scale1 * 0.485953114052605, scale1 * 0.253818527788335), 0, 0, 0);
            poly130.AddVertexAt(12, new Point2d(scale1 * 0.492732685248208, scale1 * 0.252275590895543), 0, 0, 0);
            poly130.Closed = true;
            poly130.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly130.Layer = "0";
            poly130.Color = color_GP;
            poly130.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly130);
            Polyline poly131 = new Polyline();
            poly131.AddVertexAt(0, new Point2d(scale1 * 0.462107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly131.AddVertexAt(1, new Point2d(scale1 * 0.459302385904375, scale1 * 0.241895833616759), 0, 0, 0);
            poly131.AddVertexAt(2, new Point2d(scale1 * 0.460237499172734, scale1 * 0.240072362743458), 0, 0, 0);
            poly131.AddVertexAt(3, new Point2d(scale1 * 0.461640169075273, scale1 * 0.237968357889651), 0, 0, 0);
            poly131.AddVertexAt(4, new Point2d(scale1 * 0.464679287197439, scale1 * 0.237687823909144), 0, 0, 0);
            poly131.AddVertexAt(5, new Point2d(scale1 * 0.468185961953785, scale1 * 0.23937102779219), 0, 0, 0);
            poly131.AddVertexAt(6, new Point2d(scale1 * 0.468185961953785, scale1 * 0.241615299636252), 0, 0, 0);
            poly131.AddVertexAt(7, new Point2d(scale1 * 0.465848178782888, scale1 * 0.244140105460821), 0, 0, 0);
            poly131.AddVertexAt(8, new Point2d(scale1 * 0.462107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly131.Closed = true;
            poly131.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly131.Layer = "0";
            poly131.Color = color_GP;
            poly131.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly131);
            Polyline poly132 = new Polyline();
            poly132.AddVertexAt(0, new Point2d(scale1 * 0.416520953876953, scale1 * 0.248628649148944), 0, 0, 0);
            poly132.AddVertexAt(1, new Point2d(scale1 * 0.423066746755466, scale1 * 0.250452120022245), 0, 0, 0);
            poly132.AddVertexAt(2, new Point2d(scale1 * 0.435924554195401, scale1 * 0.249750785070974), 0, 0, 0);
            poly132.AddVertexAt(3, new Point2d(scale1 * 0.442470347073914, scale1 * 0.248628649148944), 0, 0, 0);
            poly132.AddVertexAt(4, new Point2d(scale1 * 0.443639238659363, scale1 * 0.246805178275643), 0, 0, 0);
            poly132.AddVertexAt(5, new Point2d(scale1 * 0.443873016976452, scale1 * 0.244981707402343), 0, 0, 0);
            poly132.AddVertexAt(6, new Point2d(scale1 * 0.444106795293542, scale1 * 0.242737435558282), 0, 0, 0);
            poly132.AddVertexAt(7, new Point2d(scale1 * 0.44597702183026, scale1 * 0.240773697694727), 0, 0, 0);
            poly132.AddVertexAt(8, new Point2d(scale1 * 0.444574351927722, scale1 * 0.237407289928637), 0, 0, 0);
            poly132.AddVertexAt(9, new Point2d(scale1 * 0.439431228951747, scale1 * 0.235583819055336), 0, 0, 0);
            poly132.AddVertexAt(10, new Point2d(scale1 * 0.429378761316889, scale1 * 0.235303285074829), 0, 0, 0);
            poly132.AddVertexAt(11, new Point2d(scale1 * 0.424469416658004, scale1 * 0.236144887016352), 0, 0, 0);
            poly132.AddVertexAt(12, new Point2d(scale1 * 0.420962741901658, scale1 * 0.237828090899398), 0, 0, 0);
            poly132.AddVertexAt(13, new Point2d(scale1 * 0.417689845462402, scale1 * 0.241054231675236), 0, 0, 0);
            poly132.AddVertexAt(14, new Point2d(scale1 * 0.415819618925684, scale1 * 0.244420639441328), 0, 0, 0);
            poly132.Closed = true;
            poly132.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly132.Layer = "0";
            poly132.Color = color_GP;
            poly132.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly132);
            Polyline poly133 = new Polyline();
            poly133.AddVertexAt(0, new Point2d(scale1 * 0.360285741001707, scale1 * 0.3), 0, 0, 0);
            poly133.AddVertexAt(1, new Point2d(scale1 * 0.361640169075273, scale1 * 0.297968357889651), 0, 0, 0);
            poly133.AddVertexAt(2, new Point2d(scale1 * 0.364679287197439, scale1 * 0.297687823909144), 0, 0, 0);
            poly133.AddVertexAt(3, new Point2d(scale1 * 0.368185961953785, scale1 * 0.29937102779219), 0, 0, 0);
            poly133.AddVertexAt(4, new Point2d(scale1 * 0.368185961953785, scale1 * 0.3), 0, 0, 0);
            poly133.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly133.Layer = "0";
            poly133.Color = color_GP;
            poly133.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly133);
            Polyline poly134 = new Polyline();
            poly134.AddVertexAt(0, new Point2d(scale1 * 0.34565464779079, scale1 * 0.3), 0, 0, 0);
            poly134.AddVertexAt(1, new Point2d(scale1 * 0.344574351927722, scale1 * 0.297407289928637), 0, 0, 0);
            poly134.AddVertexAt(2, new Point2d(scale1 * 0.339431228951747, scale1 * 0.295583819055336), 0, 0, 0);
            poly134.AddVertexAt(3, new Point2d(scale1 * 0.329378761316889, scale1 * 0.295303285074829), 0, 0, 0);
            poly134.AddVertexAt(4, new Point2d(scale1 * 0.324469416658004, scale1 * 0.296144887016352), 0, 0, 0);
            poly134.AddVertexAt(5, new Point2d(scale1 * 0.320962741901658, scale1 * 0.297828090899398), 0, 0, 0);
            poly134.AddVertexAt(6, new Point2d(scale1 * 0.318759355857569, scale1 * 0.3), 0, 0, 0);
            poly134.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly134.Layer = "0";
            poly134.Color = color_GP;
            poly134.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly134);
            Polyline poly135 = new Polyline();
            poly135.AddVertexAt(0, new Point2d(scale1 * 0.358367272636016, scale1 * 0.282994061761134), 0, 0, 0);
            poly135.AddVertexAt(1, new Point2d(scale1 * 0.356497046099298, scale1 * 0.282853794770882), 0, 0, 0);
            poly135.AddVertexAt(2, new Point2d(scale1 * 0.354159262928401, scale1 * 0.28355512972215), 0, 0, 0);
            poly135.AddVertexAt(3, new Point2d(scale1 * 0.352289036391683, scale1 * 0.285518867585705), 0, 0, 0);
            poly135.AddVertexAt(4, new Point2d(scale1 * 0.353925484611311, scale1 * 0.287061804478495), 0, 0, 0);
            poly135.AddVertexAt(5, new Point2d(scale1 * 0.358367272636016, scale1 * 0.286220202536974), 0, 0, 0);
            poly135.AddVertexAt(6, new Point2d(scale1 * 0.359536164221465, scale1 * 0.284817532634435), 0, 0, 0);
            poly135.Closed = true;
            poly135.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly135.Layer = "0";
            poly135.Color = color_GP;
            poly135.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly135);
            Polyline poly136 = new Polyline();
            poly136.AddVertexAt(0, new Point2d(scale1 * 0.367952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly136.AddVertexAt(1, new Point2d(scale1 * 0.366081957099978, scale1 * 0.282012192829359), 0, 0, 0);
            poly136.AddVertexAt(2, new Point2d(scale1 * 0.365848178782888, scale1 * 0.280048454965803), 0, 0, 0);
            poly136.AddVertexAt(3, new Point2d(scale1 * 0.368419740270875, scale1 * 0.278505518073012), 0, 0, 0);
            poly136.AddVertexAt(4, new Point2d(scale1 * 0.371692636710131, scale1 * 0.278926319043773), 0, 0, 0);
            poly136.AddVertexAt(5, new Point2d(scale1 * 0.372393971661401, scale1 * 0.281030323897581), 0, 0, 0);
            poly136.AddVertexAt(6, new Point2d(scale1 * 0.371926415027221, scale1 * 0.282713527780627), 0, 0, 0);
            poly136.AddVertexAt(7, new Point2d(scale1 * 0.367952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly136.Closed = true;
            poly136.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly136.Layer = "0";
            poly136.Color = color_GP;
            poly136.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly136);
            Polyline poly137 = new Polyline();
            poly137.AddVertexAt(0, new Point2d(scale1 * 0.316577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly137.AddVertexAt(1, new Point2d(scale1 * 0.314239328956148, scale1 * 0.285251085149766), 0, 0, 0);
            poly137.AddVertexAt(2, new Point2d(scale1 * 0.315641998858686, scale1 * 0.283427614276465), 0, 0, 0);
            poly137.AddVertexAt(3, new Point2d(scale1 * 0.319616230249212, scale1 * 0.282726279325197), 0, 0, 0);
            poly137.AddVertexAt(4, new Point2d(scale1 * 0.322655348371378, scale1 * 0.283988682237481), 0, 0, 0);
            poly137.AddVertexAt(5, new Point2d(scale1 * 0.322655348371378, scale1 * 0.286653755052303), 0, 0, 0);
            poly137.AddVertexAt(6, new Point2d(scale1 * 0.316577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly137.Closed = true;
            poly137.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly137.Layer = "0";
            poly137.Color = color_GP;
            poly137.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly137);
            Polyline poly138 = new Polyline();
            poly138.AddVertexAt(0, new Point2d(scale1 * 0.311080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly138.AddVertexAt(1, new Point2d(scale1 * 0.309476961807309, scale1 * 0.263254670769046), 0, 0, 0);
            poly138.AddVertexAt(2, new Point2d(scale1 * 0.309710740124399, scale1 * 0.259607729022446), 0, 0, 0);
            poly138.AddVertexAt(3, new Point2d(scale1 * 0.310645853392758, scale1 * 0.256802389217368), 0, 0, 0);
            poly138.AddVertexAt(4, new Point2d(scale1 * 0.314853863100373, scale1 * 0.255960787275844), 0, 0, 0);
            poly138.AddVertexAt(5, new Point2d(scale1 * 0.319061872807989, scale1 * 0.256521855236861), 0, 0, 0);
            poly138.AddVertexAt(6, new Point2d(scale1 * 0.322568547564335, scale1 * 0.258625860090668), 0, 0, 0);
            poly138.AddVertexAt(7, new Point2d(scale1 * 0.319061872807989, scale1 * 0.263535204749552), 0, 0, 0);
            poly138.AddVertexAt(8, new Point2d(scale1 * 0.311080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly138.Closed = true;
            poly138.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly138.Layer = "0";
            poly138.Color = color_GP;
            poly138.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly138);
            Polyline poly139 = new Polyline();
            poly139.AddVertexAt(0, new Point2d(scale1 * 0.387975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly139.AddVertexAt(1, new Point2d(scale1 * 0.384702548523924, scale1 * 0.288936281530065), 0, 0, 0);
            poly139.AddVertexAt(2, new Point2d(scale1 * 0.382598543670117, scale1 * 0.283746402890672), 0, 0, 0);
            poly139.AddVertexAt(3, new Point2d(scale1 * 0.382598543670117, scale1 * 0.281221597066103), 0, 0, 0);
            poly139.AddVertexAt(4, new Point2d(scale1 * 0.381897208718848, scale1 * 0.275751184446203), 0, 0, 0);
            poly139.AddVertexAt(5, new Point2d(scale1 * 0.385403883475194, scale1 * 0.273366645611888), 0, 0, 0);
            poly139.AddVertexAt(6, new Point2d(scale1 * 0.390313228134078, scale1 * 0.272525043670365), 0, 0, 0);
            poly139.AddVertexAt(7, new Point2d(scale1 * 0.401768365671475, scale1 * 0.272945844641126), 0, 0, 0);
            poly139.AddVertexAt(8, new Point2d(scale1 * 0.404807483793642, scale1 * 0.274488781533918), 0, 0, 0);
            poly139.AddVertexAt(9, new Point2d(scale1 * 0.405742597062001, scale1 * 0.278275990270772), 0, 0, 0);
            poly139.AddVertexAt(10, new Point2d(scale1 * 0.402703478939834, scale1 * 0.28037999512458), 0, 0, 0);
            poly139.AddVertexAt(11, new Point2d(scale1 * 0.399898139134758, scale1 * 0.282764533958894), 0, 0, 0);
            poly139.AddVertexAt(12, new Point2d(scale1 * 0.39756035596386, scale1 * 0.286271208715241), 0, 0, 0);
            poly139.AddVertexAt(13, new Point2d(scale1 * 0.395690129427142, scale1 * 0.288796014539811), 0, 0, 0);
            poly139.AddVertexAt(14, new Point2d(scale1 * 0.391482119719527, scale1 * 0.289777883471588), 0, 0, 0);
            poly139.AddVertexAt(15, new Point2d(scale1 * 0.387975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly139.Closed = true;
            poly139.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly139.Layer = "0";
            poly139.Color = color_GP;
            poly139.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly139);
            Polyline poly140 = new Polyline();
            poly140.AddVertexAt(0, new Point2d(scale1 * 0.358367272636016, scale1 * 0.274437775355651), 0, 0, 0);
            poly140.AddVertexAt(1, new Point2d(scale1 * 0.366081957099978, scale1 * 0.272754571472604), 0, 0, 0);
            poly140.AddVertexAt(2, new Point2d(scale1 * 0.369354853539234, scale1 * 0.268125760794228), 0, 0, 0);
            poly140.AddVertexAt(3, new Point2d(scale1 * 0.366081957099978, scale1 * 0.263637217106105), 0, 0, 0);
            poly140.AddVertexAt(4, new Point2d(scale1 * 0.358367272636016, scale1 * 0.261954013223059), 0, 0, 0);
            poly140.AddVertexAt(5, new Point2d(scale1 * 0.349717474903696, scale1 * 0.261392945262043), 0, 0, 0);
            poly140.AddVertexAt(6, new Point2d(scale1 * 0.345041908561901, scale1 * 0.26237481419382), 0, 0, 0);
            poly140.AddVertexAt(7, new Point2d(scale1 * 0.341301455488465, scale1 * 0.26489962001839), 0, 0, 0);
            poly140.AddVertexAt(8, new Point2d(scale1 * 0.330781431219427, scale1 * 0.268967362735751), 0, 0, 0);
            poly140.AddVertexAt(9, new Point2d(scale1 * 0.325170751609273, scale1 * 0.272894838462859), 0, 0, 0);
            poly140.AddVertexAt(10, new Point2d(scale1 * 0.326573421511812, scale1 * 0.27471830933616), 0, 0, 0);
            poly140.AddVertexAt(11, new Point2d(scale1 * 0.334989440927042, scale1 * 0.278365251082757), 0, 0, 0);
            poly140.AddVertexAt(12, new Point2d(scale1 * 0.341535233805555, scale1 * 0.279066586034028), 0, 0, 0);
            poly140.AddVertexAt(13, new Point2d(scale1 * 0.34574324351317, scale1 * 0.277804183121743), 0, 0, 0);
            poly140.AddVertexAt(14, new Point2d(scale1 * 0.351353923123324, scale1 * 0.276541780209458), 0, 0, 0);
            poly140.AddVertexAt(15, new Point2d(scale1 * 0.355795711148029, scale1 * 0.275840445258188), 0, 0, 0);
            poly140.Closed = true;
            poly140.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly140.Layer = "0";
            poly140.Color = color_GP;
            poly140.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly140);
            Polyline poly141 = new Polyline();
            poly141.AddVertexAt(0, new Point2d(scale1 * 0.378898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly141.AddVertexAt(1, new Point2d(scale1 * 0.372741026607184, scale1 * 0.262017770945895), 0, 0, 0);
            poly141.AddVertexAt(2, new Point2d(scale1 * 0.36743251686181, scale1 * 0.259722492923566), 0, 0, 0);
            poly141.AddVertexAt(3, new Point2d(scale1 * 0.366370814912736, scale1 * 0.256662122227115), 0, 0, 0);
            poly141.AddVertexAt(4, new Point2d(scale1 * 0.37167932465811, scale1 * 0.255769514107321), 0, 0, 0);
            poly141.AddVertexAt(5, new Point2d(scale1 * 0.378261876742373, scale1 * 0.2567896376728), 0, 0, 0);
            poly141.AddVertexAt(6, new Point2d(scale1 * 0.382084003759042, scale1 * 0.259467462032192), 0, 0, 0);
            poly141.AddVertexAt(7, new Point2d(scale1 * 0.378898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly141.Closed = true;
            poly141.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly141.Layer = "0";
            poly141.Color = color_GP;
            poly141.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly141);
            Polyline poly142 = new Polyline();
            poly142.AddVertexAt(0, new Point2d(scale1 * 0.396473138321644, scale1 * 0.248628649148944), 0, 0, 0);
            poly142.AddVertexAt(1, new Point2d(scale1 * 0.396706916638733, scale1 * 0.243298503519296), 0, 0, 0);
            poly142.AddVertexAt(2, new Point2d(scale1 * 0.395070468419106, scale1 * 0.242036100607014), 0, 0, 0);
            poly142.AddVertexAt(3, new Point2d(scale1 * 0.39086245871149, scale1 * 0.240633430704475), 0, 0, 0);
            poly142.AddVertexAt(4, new Point2d(scale1 * 0.386888227320964, scale1 * 0.240773697694727), 0, 0, 0);
            poly142.AddVertexAt(5, new Point2d(scale1 * 0.383849109198798, scale1 * 0.242036100607014), 0, 0, 0);
            poly142.AddVertexAt(6, new Point2d(scale1 * 0.380576212759542, scale1 * 0.243298503519296), 0, 0, 0);
            poly142.AddVertexAt(7, new Point2d(scale1 * 0.377537094637375, scale1 * 0.244560906431583), 0, 0, 0);
            poly142.AddVertexAt(8, new Point2d(scale1 * 0.375666868100657, scale1 * 0.246945445265898), 0, 0, 0);
            poly142.AddVertexAt(9, new Point2d(scale1 * 0.375666868100657, scale1 * 0.248628649148944), 0, 0, 0);
            poly142.AddVertexAt(10, new Point2d(scale1 * 0.379173542857003, scale1 * 0.252275590895543), 0, 0, 0);
            poly142.AddVertexAt(11, new Point2d(scale1 * 0.385953114052606, scale1 * 0.253818527788335), 0, 0, 0);
            poly142.AddVertexAt(12, new Point2d(scale1 * 0.392732685248208, scale1 * 0.252275590895543), 0, 0, 0);
            poly142.Closed = true;
            poly142.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly142.Layer = "0";
            poly142.Color = color_GP;
            poly142.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly142);
            Polyline poly143 = new Polyline();
            poly143.AddVertexAt(0, new Point2d(scale1 * 0.362107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly143.AddVertexAt(1, new Point2d(scale1 * 0.359302385904375, scale1 * 0.241895833616759), 0, 0, 0);
            poly143.AddVertexAt(2, new Point2d(scale1 * 0.360237499172734, scale1 * 0.240072362743458), 0, 0, 0);
            poly143.AddVertexAt(3, new Point2d(scale1 * 0.361640169075273, scale1 * 0.237968357889651), 0, 0, 0);
            poly143.AddVertexAt(4, new Point2d(scale1 * 0.364679287197439, scale1 * 0.237687823909144), 0, 0, 0);
            poly143.AddVertexAt(5, new Point2d(scale1 * 0.368185961953785, scale1 * 0.23937102779219), 0, 0, 0);
            poly143.AddVertexAt(6, new Point2d(scale1 * 0.368185961953785, scale1 * 0.241615299636252), 0, 0, 0);
            poly143.AddVertexAt(7, new Point2d(scale1 * 0.365848178782888, scale1 * 0.244140105460821), 0, 0, 0);
            poly143.AddVertexAt(8, new Point2d(scale1 * 0.362107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly143.Closed = true;
            poly143.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly143.Layer = "0";
            poly143.Color = color_GP;
            poly143.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly143);
            Polyline poly144 = new Polyline();
            poly144.AddVertexAt(0, new Point2d(scale1 * 0.316520953876953, scale1 * 0.248628649148944), 0, 0, 0);
            poly144.AddVertexAt(1, new Point2d(scale1 * 0.323066746755466, scale1 * 0.250452120022245), 0, 0, 0);
            poly144.AddVertexAt(2, new Point2d(scale1 * 0.335924554195401, scale1 * 0.249750785070974), 0, 0, 0);
            poly144.AddVertexAt(3, new Point2d(scale1 * 0.342470347073914, scale1 * 0.248628649148944), 0, 0, 0);
            poly144.AddVertexAt(4, new Point2d(scale1 * 0.343639238659363, scale1 * 0.246805178275643), 0, 0, 0);
            poly144.AddVertexAt(5, new Point2d(scale1 * 0.343873016976453, scale1 * 0.244981707402343), 0, 0, 0);
            poly144.AddVertexAt(6, new Point2d(scale1 * 0.344106795293542, scale1 * 0.242737435558282), 0, 0, 0);
            poly144.AddVertexAt(7, new Point2d(scale1 * 0.34597702183026, scale1 * 0.240773697694727), 0, 0, 0);
            poly144.AddVertexAt(8, new Point2d(scale1 * 0.344574351927722, scale1 * 0.237407289928637), 0, 0, 0);
            poly144.AddVertexAt(9, new Point2d(scale1 * 0.339431228951747, scale1 * 0.235583819055336), 0, 0, 0);
            poly144.AddVertexAt(10, new Point2d(scale1 * 0.329378761316889, scale1 * 0.235303285074829), 0, 0, 0);
            poly144.AddVertexAt(11, new Point2d(scale1 * 0.324469416658004, scale1 * 0.236144887016352), 0, 0, 0);
            poly144.AddVertexAt(12, new Point2d(scale1 * 0.320962741901658, scale1 * 0.237828090899398), 0, 0, 0);
            poly144.AddVertexAt(13, new Point2d(scale1 * 0.317689845462402, scale1 * 0.241054231675236), 0, 0, 0);
            poly144.AddVertexAt(14, new Point2d(scale1 * 0.315819618925684, scale1 * 0.244420639441328), 0, 0, 0);
            poly144.Closed = true;
            poly144.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly144.Layer = "0";
            poly144.Color = color_GP;
            poly144.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly144);
            Polyline poly145 = new Polyline();
            poly145.AddVertexAt(0, new Point2d(scale1 * 0.260285741001707, scale1 * 0.3), 0, 0, 0);
            poly145.AddVertexAt(1, new Point2d(scale1 * 0.261640169075273, scale1 * 0.297968357889651), 0, 0, 0);
            poly145.AddVertexAt(2, new Point2d(scale1 * 0.264679287197439, scale1 * 0.297687823909144), 0, 0, 0);
            poly145.AddVertexAt(3, new Point2d(scale1 * 0.268185961953785, scale1 * 0.29937102779219), 0, 0, 0);
            poly145.AddVertexAt(4, new Point2d(scale1 * 0.268185961953785, scale1 * 0.3), 0, 0, 0);
            poly145.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly145.Layer = "0";
            poly145.Color = color_GP;
            poly145.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly145);
            Polyline poly146 = new Polyline();
            poly146.AddVertexAt(0, new Point2d(scale1 * 0.24565464779079, scale1 * 0.3), 0, 0, 0);
            poly146.AddVertexAt(1, new Point2d(scale1 * 0.244574351927722, scale1 * 0.297407289928637), 0, 0, 0);
            poly146.AddVertexAt(2, new Point2d(scale1 * 0.239431228951748, scale1 * 0.295583819055336), 0, 0, 0);
            poly146.AddVertexAt(3, new Point2d(scale1 * 0.229378761316889, scale1 * 0.295303285074829), 0, 0, 0);
            poly146.AddVertexAt(4, new Point2d(scale1 * 0.224469416658004, scale1 * 0.296144887016352), 0, 0, 0);
            poly146.AddVertexAt(5, new Point2d(scale1 * 0.220962741901658, scale1 * 0.297828090899398), 0, 0, 0);
            poly146.AddVertexAt(6, new Point2d(scale1 * 0.218759355857569, scale1 * 0.3), 0, 0, 0);
            poly146.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly146.Layer = "0";
            poly146.Color = color_GP;
            poly146.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly146);
            Polyline poly147 = new Polyline();
            poly147.AddVertexAt(0, new Point2d(scale1 * 0.258367272636016, scale1 * 0.282994061761134), 0, 0, 0);
            poly147.AddVertexAt(1, new Point2d(scale1 * 0.256497046099298, scale1 * 0.282853794770882), 0, 0, 0);
            poly147.AddVertexAt(2, new Point2d(scale1 * 0.254159262928401, scale1 * 0.28355512972215), 0, 0, 0);
            poly147.AddVertexAt(3, new Point2d(scale1 * 0.252289036391683, scale1 * 0.285518867585705), 0, 0, 0);
            poly147.AddVertexAt(4, new Point2d(scale1 * 0.253925484611311, scale1 * 0.287061804478495), 0, 0, 0);
            poly147.AddVertexAt(5, new Point2d(scale1 * 0.258367272636016, scale1 * 0.286220202536974), 0, 0, 0);
            poly147.AddVertexAt(6, new Point2d(scale1 * 0.259536164221465, scale1 * 0.284817532634435), 0, 0, 0);
            poly147.Closed = true;
            poly147.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly147.Layer = "0";
            poly147.Color = color_GP;
            poly147.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly147);
            Polyline poly148 = new Polyline();
            poly148.AddVertexAt(0, new Point2d(scale1 * 0.267952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly148.AddVertexAt(1, new Point2d(scale1 * 0.266081957099978, scale1 * 0.282012192829359), 0, 0, 0);
            poly148.AddVertexAt(2, new Point2d(scale1 * 0.265848178782888, scale1 * 0.280048454965803), 0, 0, 0);
            poly148.AddVertexAt(3, new Point2d(scale1 * 0.268419740270875, scale1 * 0.278505518073012), 0, 0, 0);
            poly148.AddVertexAt(4, new Point2d(scale1 * 0.271692636710131, scale1 * 0.278926319043773), 0, 0, 0);
            poly148.AddVertexAt(5, new Point2d(scale1 * 0.272393971661401, scale1 * 0.281030323897581), 0, 0, 0);
            poly148.AddVertexAt(6, new Point2d(scale1 * 0.271926415027221, scale1 * 0.282713527780627), 0, 0, 0);
            poly148.AddVertexAt(7, new Point2d(scale1 * 0.267952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly148.Closed = true;
            poly148.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly148.Layer = "0";
            poly148.Color = color_GP;
            poly148.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly148);
            Polyline poly149 = new Polyline();
            poly149.AddVertexAt(0, new Point2d(scale1 * 0.216577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly149.AddVertexAt(1, new Point2d(scale1 * 0.214239328956148, scale1 * 0.285251085149766), 0, 0, 0);
            poly149.AddVertexAt(2, new Point2d(scale1 * 0.215641998858686, scale1 * 0.283427614276465), 0, 0, 0);
            poly149.AddVertexAt(3, new Point2d(scale1 * 0.219616230249212, scale1 * 0.282726279325197), 0, 0, 0);
            poly149.AddVertexAt(4, new Point2d(scale1 * 0.222655348371378, scale1 * 0.283988682237481), 0, 0, 0);
            poly149.AddVertexAt(5, new Point2d(scale1 * 0.222655348371378, scale1 * 0.286653755052303), 0, 0, 0);
            poly149.AddVertexAt(6, new Point2d(scale1 * 0.216577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly149.Closed = true;
            poly149.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly149.Layer = "0";
            poly149.Color = color_GP;
            poly149.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly149);
            Polyline poly150 = new Polyline();
            poly150.AddVertexAt(0, new Point2d(scale1 * 0.211080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly150.AddVertexAt(1, new Point2d(scale1 * 0.209476961807309, scale1 * 0.263254670769046), 0, 0, 0);
            poly150.AddVertexAt(2, new Point2d(scale1 * 0.209710740124399, scale1 * 0.259607729022446), 0, 0, 0);
            poly150.AddVertexAt(3, new Point2d(scale1 * 0.210645853392758, scale1 * 0.256802389217368), 0, 0, 0);
            poly150.AddVertexAt(4, new Point2d(scale1 * 0.214853863100373, scale1 * 0.255960787275844), 0, 0, 0);
            poly150.AddVertexAt(5, new Point2d(scale1 * 0.219061872807989, scale1 * 0.256521855236861), 0, 0, 0);
            poly150.AddVertexAt(6, new Point2d(scale1 * 0.222568547564335, scale1 * 0.258625860090668), 0, 0, 0);
            poly150.AddVertexAt(7, new Point2d(scale1 * 0.219061872807989, scale1 * 0.263535204749552), 0, 0, 0);
            poly150.AddVertexAt(8, new Point2d(scale1 * 0.211080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly150.Closed = true;
            poly150.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly150.Layer = "0";
            poly150.Color = color_GP;
            poly150.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly150);
            Polyline poly151 = new Polyline();
            poly151.AddVertexAt(0, new Point2d(scale1 * 0.287975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly151.AddVertexAt(1, new Point2d(scale1 * 0.284702548523924, scale1 * 0.288936281530065), 0, 0, 0);
            poly151.AddVertexAt(2, new Point2d(scale1 * 0.282598543670117, scale1 * 0.283746402890672), 0, 0, 0);
            poly151.AddVertexAt(3, new Point2d(scale1 * 0.282598543670117, scale1 * 0.281221597066103), 0, 0, 0);
            poly151.AddVertexAt(4, new Point2d(scale1 * 0.281897208718848, scale1 * 0.275751184446203), 0, 0, 0);
            poly151.AddVertexAt(5, new Point2d(scale1 * 0.285403883475194, scale1 * 0.273366645611888), 0, 0, 0);
            poly151.AddVertexAt(6, new Point2d(scale1 * 0.290313228134079, scale1 * 0.272525043670365), 0, 0, 0);
            poly151.AddVertexAt(7, new Point2d(scale1 * 0.301768365671475, scale1 * 0.272945844641126), 0, 0, 0);
            poly151.AddVertexAt(8, new Point2d(scale1 * 0.304807483793642, scale1 * 0.274488781533918), 0, 0, 0);
            poly151.AddVertexAt(9, new Point2d(scale1 * 0.305742597062001, scale1 * 0.278275990270772), 0, 0, 0);
            poly151.AddVertexAt(10, new Point2d(scale1 * 0.302703478939834, scale1 * 0.28037999512458), 0, 0, 0);
            poly151.AddVertexAt(11, new Point2d(scale1 * 0.299898139134757, scale1 * 0.282764533958894), 0, 0, 0);
            poly151.AddVertexAt(12, new Point2d(scale1 * 0.29756035596386, scale1 * 0.286271208715241), 0, 0, 0);
            poly151.AddVertexAt(13, new Point2d(scale1 * 0.295690129427142, scale1 * 0.288796014539811), 0, 0, 0);
            poly151.AddVertexAt(14, new Point2d(scale1 * 0.291482119719527, scale1 * 0.289777883471588), 0, 0, 0);
            poly151.AddVertexAt(15, new Point2d(scale1 * 0.287975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly151.Closed = true;
            poly151.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly151.Layer = "0";
            poly151.Color = color_GP;
            poly151.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly151);
            Polyline poly152 = new Polyline();
            poly152.AddVertexAt(0, new Point2d(scale1 * 0.258367272636016, scale1 * 0.274437775355651), 0, 0, 0);
            poly152.AddVertexAt(1, new Point2d(scale1 * 0.266081957099978, scale1 * 0.272754571472604), 0, 0, 0);
            poly152.AddVertexAt(2, new Point2d(scale1 * 0.269354853539234, scale1 * 0.268125760794228), 0, 0, 0);
            poly152.AddVertexAt(3, new Point2d(scale1 * 0.266081957099978, scale1 * 0.263637217106105), 0, 0, 0);
            poly152.AddVertexAt(4, new Point2d(scale1 * 0.258367272636016, scale1 * 0.261954013223059), 0, 0, 0);
            poly152.AddVertexAt(5, new Point2d(scale1 * 0.249717474903696, scale1 * 0.261392945262043), 0, 0, 0);
            poly152.AddVertexAt(6, new Point2d(scale1 * 0.245041908561901, scale1 * 0.26237481419382), 0, 0, 0);
            poly152.AddVertexAt(7, new Point2d(scale1 * 0.241301455488465, scale1 * 0.26489962001839), 0, 0, 0);
            poly152.AddVertexAt(8, new Point2d(scale1 * 0.230781431219427, scale1 * 0.268967362735751), 0, 0, 0);
            poly152.AddVertexAt(9, new Point2d(scale1 * 0.225170751609273, scale1 * 0.272894838462859), 0, 0, 0);
            poly152.AddVertexAt(10, new Point2d(scale1 * 0.226573421511812, scale1 * 0.27471830933616), 0, 0, 0);
            poly152.AddVertexAt(11, new Point2d(scale1 * 0.234989440927042, scale1 * 0.278365251082757), 0, 0, 0);
            poly152.AddVertexAt(12, new Point2d(scale1 * 0.241535233805555, scale1 * 0.279066586034028), 0, 0, 0);
            poly152.AddVertexAt(13, new Point2d(scale1 * 0.24574324351317, scale1 * 0.277804183121743), 0, 0, 0);
            poly152.AddVertexAt(14, new Point2d(scale1 * 0.251353923123324, scale1 * 0.276541780209458), 0, 0, 0);
            poly152.AddVertexAt(15, new Point2d(scale1 * 0.255795711148029, scale1 * 0.275840445258188), 0, 0, 0);
            poly152.Closed = true;
            poly152.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly152.Layer = "0";
            poly152.Color = color_GP;
            poly152.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly152);
            Polyline poly153 = new Polyline();
            poly153.AddVertexAt(0, new Point2d(scale1 * 0.278898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly153.AddVertexAt(1, new Point2d(scale1 * 0.272741026607183, scale1 * 0.262017770945895), 0, 0, 0);
            poly153.AddVertexAt(2, new Point2d(scale1 * 0.267432516861809, scale1 * 0.259722492923566), 0, 0, 0);
            poly153.AddVertexAt(3, new Point2d(scale1 * 0.266370814912736, scale1 * 0.256662122227115), 0, 0, 0);
            poly153.AddVertexAt(4, new Point2d(scale1 * 0.27167932465811, scale1 * 0.255769514107321), 0, 0, 0);
            poly153.AddVertexAt(5, new Point2d(scale1 * 0.278261876742373, scale1 * 0.2567896376728), 0, 0, 0);
            poly153.AddVertexAt(6, new Point2d(scale1 * 0.282084003759042, scale1 * 0.259467462032192), 0, 0, 0);
            poly153.AddVertexAt(7, new Point2d(scale1 * 0.278898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly153.Closed = true;
            poly153.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly153.Layer = "0";
            poly153.Color = color_GP;
            poly153.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly153);
            Polyline poly154 = new Polyline();
            poly154.AddVertexAt(0, new Point2d(scale1 * 0.296473138321644, scale1 * 0.248628649148944), 0, 0, 0);
            poly154.AddVertexAt(1, new Point2d(scale1 * 0.296706916638733, scale1 * 0.243298503519296), 0, 0, 0);
            poly154.AddVertexAt(2, new Point2d(scale1 * 0.295070468419106, scale1 * 0.242036100607014), 0, 0, 0);
            poly154.AddVertexAt(3, new Point2d(scale1 * 0.29086245871149, scale1 * 0.240633430704475), 0, 0, 0);
            poly154.AddVertexAt(4, new Point2d(scale1 * 0.286888227320964, scale1 * 0.240773697694727), 0, 0, 0);
            poly154.AddVertexAt(5, new Point2d(scale1 * 0.283849109198798, scale1 * 0.242036100607014), 0, 0, 0);
            poly154.AddVertexAt(6, new Point2d(scale1 * 0.280576212759541, scale1 * 0.243298503519296), 0, 0, 0);
            poly154.AddVertexAt(7, new Point2d(scale1 * 0.277537094637375, scale1 * 0.244560906431583), 0, 0, 0);
            poly154.AddVertexAt(8, new Point2d(scale1 * 0.275666868100657, scale1 * 0.246945445265898), 0, 0, 0);
            poly154.AddVertexAt(9, new Point2d(scale1 * 0.275666868100657, scale1 * 0.248628649148944), 0, 0, 0);
            poly154.AddVertexAt(10, new Point2d(scale1 * 0.279173542857003, scale1 * 0.252275590895543), 0, 0, 0);
            poly154.AddVertexAt(11, new Point2d(scale1 * 0.285953114052605, scale1 * 0.253818527788335), 0, 0, 0);
            poly154.AddVertexAt(12, new Point2d(scale1 * 0.292732685248208, scale1 * 0.252275590895543), 0, 0, 0);
            poly154.Closed = true;
            poly154.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly154.Layer = "0";
            poly154.Color = color_GP;
            poly154.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly154);
            Polyline poly155 = new Polyline();
            poly155.AddVertexAt(0, new Point2d(scale1 * 0.262107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly155.AddVertexAt(1, new Point2d(scale1 * 0.259302385904375, scale1 * 0.241895833616759), 0, 0, 0);
            poly155.AddVertexAt(2, new Point2d(scale1 * 0.260237499172734, scale1 * 0.240072362743458), 0, 0, 0);
            poly155.AddVertexAt(3, new Point2d(scale1 * 0.261640169075273, scale1 * 0.237968357889651), 0, 0, 0);
            poly155.AddVertexAt(4, new Point2d(scale1 * 0.264679287197439, scale1 * 0.237687823909144), 0, 0, 0);
            poly155.AddVertexAt(5, new Point2d(scale1 * 0.268185961953785, scale1 * 0.23937102779219), 0, 0, 0);
            poly155.AddVertexAt(6, new Point2d(scale1 * 0.268185961953785, scale1 * 0.241615299636252), 0, 0, 0);
            poly155.AddVertexAt(7, new Point2d(scale1 * 0.265848178782888, scale1 * 0.244140105460821), 0, 0, 0);
            poly155.AddVertexAt(8, new Point2d(scale1 * 0.262107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly155.Closed = true;
            poly155.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly155.Layer = "0";
            poly155.Color = color_GP;
            poly155.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly155);
            Polyline poly156 = new Polyline();
            poly156.AddVertexAt(0, new Point2d(scale1 * 0.216520953876953, scale1 * 0.248628649148944), 0, 0, 0);
            poly156.AddVertexAt(1, new Point2d(scale1 * 0.223066746755466, scale1 * 0.250452120022245), 0, 0, 0);
            poly156.AddVertexAt(2, new Point2d(scale1 * 0.235924554195401, scale1 * 0.249750785070974), 0, 0, 0);
            poly156.AddVertexAt(3, new Point2d(scale1 * 0.242470347073914, scale1 * 0.248628649148944), 0, 0, 0);
            poly156.AddVertexAt(4, new Point2d(scale1 * 0.243639238659363, scale1 * 0.246805178275643), 0, 0, 0);
            poly156.AddVertexAt(5, new Point2d(scale1 * 0.243873016976452, scale1 * 0.244981707402343), 0, 0, 0);
            poly156.AddVertexAt(6, new Point2d(scale1 * 0.244106795293542, scale1 * 0.242737435558282), 0, 0, 0);
            poly156.AddVertexAt(7, new Point2d(scale1 * 0.24597702183026, scale1 * 0.240773697694727), 0, 0, 0);
            poly156.AddVertexAt(8, new Point2d(scale1 * 0.244574351927722, scale1 * 0.237407289928637), 0, 0, 0);
            poly156.AddVertexAt(9, new Point2d(scale1 * 0.239431228951748, scale1 * 0.235583819055336), 0, 0, 0);
            poly156.AddVertexAt(10, new Point2d(scale1 * 0.229378761316889, scale1 * 0.235303285074829), 0, 0, 0);
            poly156.AddVertexAt(11, new Point2d(scale1 * 0.224469416658004, scale1 * 0.236144887016352), 0, 0, 0);
            poly156.AddVertexAt(12, new Point2d(scale1 * 0.220962741901658, scale1 * 0.237828090899398), 0, 0, 0);
            poly156.AddVertexAt(13, new Point2d(scale1 * 0.217689845462402, scale1 * 0.241054231675236), 0, 0, 0);
            poly156.AddVertexAt(14, new Point2d(scale1 * 0.215819618925684, scale1 * 0.244420639441328), 0, 0, 0);
            poly156.Closed = true;
            poly156.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly156.Layer = "0";
            poly156.Color = color_GP;
            poly156.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly156);
            Polyline poly157 = new Polyline();
            poly157.AddVertexAt(0, new Point2d(scale1 * 0.160285741001707, scale1 * 0.3), 0, 0, 0);
            poly157.AddVertexAt(1, new Point2d(scale1 * 0.161640169075273, scale1 * 0.297968357889651), 0, 0, 0);
            poly157.AddVertexAt(2, new Point2d(scale1 * 0.164679287197439, scale1 * 0.297687823909144), 0, 0, 0);
            poly157.AddVertexAt(3, new Point2d(scale1 * 0.168185961953785, scale1 * 0.29937102779219), 0, 0, 0);
            poly157.AddVertexAt(4, new Point2d(scale1 * 0.168185961953785, scale1 * 0.3), 0, 0, 0);
            poly157.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly157.Layer = "0";
            poly157.Color = color_GP;
            poly157.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly157);
            Polyline poly158 = new Polyline();
            poly158.AddVertexAt(0, new Point2d(scale1 * 0.14565464779079, scale1 * 0.3), 0, 0, 0);
            poly158.AddVertexAt(1, new Point2d(scale1 * 0.144574351927722, scale1 * 0.297407289928637), 0, 0, 0);
            poly158.AddVertexAt(2, new Point2d(scale1 * 0.139431228951747, scale1 * 0.295583819055336), 0, 0, 0);
            poly158.AddVertexAt(3, new Point2d(scale1 * 0.129378761316889, scale1 * 0.295303285074829), 0, 0, 0);
            poly158.AddVertexAt(4, new Point2d(scale1 * 0.124469416658004, scale1 * 0.296144887016352), 0, 0, 0);
            poly158.AddVertexAt(5, new Point2d(scale1 * 0.120962741901658, scale1 * 0.297828090899398), 0, 0, 0);
            poly158.AddVertexAt(6, new Point2d(scale1 * 0.118759355857569, scale1 * 0.3), 0, 0, 0);
            poly158.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly158.Layer = "0";
            poly158.Color = color_GP;
            poly158.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly158);
            Polyline poly159 = new Polyline();
            poly159.AddVertexAt(0, new Point2d(scale1 * 0.158367272636016, scale1 * 0.282994061761134), 0, 0, 0);
            poly159.AddVertexAt(1, new Point2d(scale1 * 0.156497046099298, scale1 * 0.282853794770882), 0, 0, 0);
            poly159.AddVertexAt(2, new Point2d(scale1 * 0.154159262928401, scale1 * 0.28355512972215), 0, 0, 0);
            poly159.AddVertexAt(3, new Point2d(scale1 * 0.152289036391683, scale1 * 0.285518867585705), 0, 0, 0);
            poly159.AddVertexAt(4, new Point2d(scale1 * 0.153925484611311, scale1 * 0.287061804478495), 0, 0, 0);
            poly159.AddVertexAt(5, new Point2d(scale1 * 0.158367272636016, scale1 * 0.286220202536974), 0, 0, 0);
            poly159.AddVertexAt(6, new Point2d(scale1 * 0.159536164221465, scale1 * 0.284817532634435), 0, 0, 0);
            poly159.Closed = true;
            poly159.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly159.Layer = "0";
            poly159.Color = color_GP;
            poly159.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly159);
            Polyline poly160 = new Polyline();
            poly160.AddVertexAt(0, new Point2d(scale1 * 0.167952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly160.AddVertexAt(1, new Point2d(scale1 * 0.166081957099978, scale1 * 0.282012192829359), 0, 0, 0);
            poly160.AddVertexAt(2, new Point2d(scale1 * 0.165848178782888, scale1 * 0.280048454965803), 0, 0, 0);
            poly160.AddVertexAt(3, new Point2d(scale1 * 0.168419740270875, scale1 * 0.278505518073012), 0, 0, 0);
            poly160.AddVertexAt(4, new Point2d(scale1 * 0.171692636710131, scale1 * 0.278926319043773), 0, 0, 0);
            poly160.AddVertexAt(5, new Point2d(scale1 * 0.172393971661401, scale1 * 0.281030323897581), 0, 0, 0);
            poly160.AddVertexAt(6, new Point2d(scale1 * 0.171926415027221, scale1 * 0.282713527780627), 0, 0, 0);
            poly160.AddVertexAt(7, new Point2d(scale1 * 0.167952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly160.Closed = true;
            poly160.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly160.Layer = "0";
            poly160.Color = color_GP;
            poly160.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly160);
            Polyline poly161 = new Polyline();
            poly161.AddVertexAt(0, new Point2d(scale1 * 0.116577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly161.AddVertexAt(1, new Point2d(scale1 * 0.114239328956148, scale1 * 0.285251085149766), 0, 0, 0);
            poly161.AddVertexAt(2, new Point2d(scale1 * 0.115641998858686, scale1 * 0.283427614276465), 0, 0, 0);
            poly161.AddVertexAt(3, new Point2d(scale1 * 0.119616230249212, scale1 * 0.282726279325197), 0, 0, 0);
            poly161.AddVertexAt(4, new Point2d(scale1 * 0.122655348371378, scale1 * 0.283988682237481), 0, 0, 0);
            poly161.AddVertexAt(5, new Point2d(scale1 * 0.122655348371378, scale1 * 0.286653755052303), 0, 0, 0);
            poly161.AddVertexAt(6, new Point2d(scale1 * 0.116577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly161.Closed = true;
            poly161.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly161.Layer = "0";
            poly161.Color = color_GP;
            poly161.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly161);
            Polyline poly162 = new Polyline();
            poly162.AddVertexAt(0, new Point2d(scale1 * 0.111080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly162.AddVertexAt(1, new Point2d(scale1 * 0.109476961807309, scale1 * 0.263254670769046), 0, 0, 0);
            poly162.AddVertexAt(2, new Point2d(scale1 * 0.109710740124399, scale1 * 0.259607729022446), 0, 0, 0);
            poly162.AddVertexAt(3, new Point2d(scale1 * 0.110645853392758, scale1 * 0.256802389217368), 0, 0, 0);
            poly162.AddVertexAt(4, new Point2d(scale1 * 0.114853863100373, scale1 * 0.255960787275844), 0, 0, 0);
            poly162.AddVertexAt(5, new Point2d(scale1 * 0.119061872807989, scale1 * 0.256521855236861), 0, 0, 0);
            poly162.AddVertexAt(6, new Point2d(scale1 * 0.122568547564335, scale1 * 0.258625860090668), 0, 0, 0);
            poly162.AddVertexAt(7, new Point2d(scale1 * 0.119061872807989, scale1 * 0.263535204749552), 0, 0, 0);
            poly162.AddVertexAt(8, new Point2d(scale1 * 0.111080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly162.Closed = true;
            poly162.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly162.Layer = "0";
            poly162.Color = color_GP;
            poly162.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly162);
            Polyline poly163 = new Polyline();
            poly163.AddVertexAt(0, new Point2d(scale1 * 0.187975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly163.AddVertexAt(1, new Point2d(scale1 * 0.184702548523924, scale1 * 0.288936281530065), 0, 0, 0);
            poly163.AddVertexAt(2, new Point2d(scale1 * 0.182598543670117, scale1 * 0.283746402890672), 0, 0, 0);
            poly163.AddVertexAt(3, new Point2d(scale1 * 0.182598543670117, scale1 * 0.281221597066103), 0, 0, 0);
            poly163.AddVertexAt(4, new Point2d(scale1 * 0.181897208718848, scale1 * 0.275751184446203), 0, 0, 0);
            poly163.AddVertexAt(5, new Point2d(scale1 * 0.185403883475194, scale1 * 0.273366645611888), 0, 0, 0);
            poly163.AddVertexAt(6, new Point2d(scale1 * 0.190313228134078, scale1 * 0.272525043670365), 0, 0, 0);
            poly163.AddVertexAt(7, new Point2d(scale1 * 0.201768365671475, scale1 * 0.272945844641126), 0, 0, 0);
            poly163.AddVertexAt(8, new Point2d(scale1 * 0.204807483793642, scale1 * 0.274488781533918), 0, 0, 0);
            poly163.AddVertexAt(9, new Point2d(scale1 * 0.205742597062001, scale1 * 0.278275990270772), 0, 0, 0);
            poly163.AddVertexAt(10, new Point2d(scale1 * 0.202703478939834, scale1 * 0.28037999512458), 0, 0, 0);
            poly163.AddVertexAt(11, new Point2d(scale1 * 0.199898139134758, scale1 * 0.282764533958894), 0, 0, 0);
            poly163.AddVertexAt(12, new Point2d(scale1 * 0.19756035596386, scale1 * 0.286271208715241), 0, 0, 0);
            poly163.AddVertexAt(13, new Point2d(scale1 * 0.195690129427142, scale1 * 0.288796014539811), 0, 0, 0);
            poly163.AddVertexAt(14, new Point2d(scale1 * 0.191482119719527, scale1 * 0.289777883471588), 0, 0, 0);
            poly163.AddVertexAt(15, new Point2d(scale1 * 0.187975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly163.Closed = true;
            poly163.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly163.Layer = "0";
            poly163.Color = color_GP;
            poly163.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly163);
            Polyline poly164 = new Polyline();
            poly164.AddVertexAt(0, new Point2d(scale1 * 0.158367272636016, scale1 * 0.274437775355651), 0, 0, 0);
            poly164.AddVertexAt(1, new Point2d(scale1 * 0.166081957099978, scale1 * 0.272754571472604), 0, 0, 0);
            poly164.AddVertexAt(2, new Point2d(scale1 * 0.169354853539234, scale1 * 0.268125760794228), 0, 0, 0);
            poly164.AddVertexAt(3, new Point2d(scale1 * 0.166081957099978, scale1 * 0.263637217106105), 0, 0, 0);
            poly164.AddVertexAt(4, new Point2d(scale1 * 0.158367272636016, scale1 * 0.261954013223059), 0, 0, 0);
            poly164.AddVertexAt(5, new Point2d(scale1 * 0.149717474903696, scale1 * 0.261392945262043), 0, 0, 0);
            poly164.AddVertexAt(6, new Point2d(scale1 * 0.145041908561901, scale1 * 0.26237481419382), 0, 0, 0);
            poly164.AddVertexAt(7, new Point2d(scale1 * 0.141301455488465, scale1 * 0.26489962001839), 0, 0, 0);
            poly164.AddVertexAt(8, new Point2d(scale1 * 0.130781431219427, scale1 * 0.268967362735751), 0, 0, 0);
            poly164.AddVertexAt(9, new Point2d(scale1 * 0.125170751609273, scale1 * 0.272894838462859), 0, 0, 0);
            poly164.AddVertexAt(10, new Point2d(scale1 * 0.126573421511812, scale1 * 0.27471830933616), 0, 0, 0);
            poly164.AddVertexAt(11, new Point2d(scale1 * 0.134989440927042, scale1 * 0.278365251082757), 0, 0, 0);
            poly164.AddVertexAt(12, new Point2d(scale1 * 0.141535233805555, scale1 * 0.279066586034028), 0, 0, 0);
            poly164.AddVertexAt(13, new Point2d(scale1 * 0.14574324351317, scale1 * 0.277804183121743), 0, 0, 0);
            poly164.AddVertexAt(14, new Point2d(scale1 * 0.151353923123324, scale1 * 0.276541780209458), 0, 0, 0);
            poly164.AddVertexAt(15, new Point2d(scale1 * 0.155795711148029, scale1 * 0.275840445258188), 0, 0, 0);
            poly164.Closed = true;
            poly164.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly164.Layer = "0";
            poly164.Color = color_GP;
            poly164.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly164);
            Polyline poly165 = new Polyline();
            poly165.AddVertexAt(0, new Point2d(scale1 * 0.178898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly165.AddVertexAt(1, new Point2d(scale1 * 0.172741026607183, scale1 * 0.262017770945895), 0, 0, 0);
            poly165.AddVertexAt(2, new Point2d(scale1 * 0.167432516861809, scale1 * 0.259722492923566), 0, 0, 0);
            poly165.AddVertexAt(3, new Point2d(scale1 * 0.166370814912736, scale1 * 0.256662122227115), 0, 0, 0);
            poly165.AddVertexAt(4, new Point2d(scale1 * 0.17167932465811, scale1 * 0.255769514107321), 0, 0, 0);
            poly165.AddVertexAt(5, new Point2d(scale1 * 0.178261876742373, scale1 * 0.2567896376728), 0, 0, 0);
            poly165.AddVertexAt(6, new Point2d(scale1 * 0.182084003759041, scale1 * 0.259467462032192), 0, 0, 0);
            poly165.AddVertexAt(7, new Point2d(scale1 * 0.178898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly165.Closed = true;
            poly165.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly165.Layer = "0";
            poly165.Color = color_GP;
            poly165.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly165);
            Polyline poly166 = new Polyline();
            poly166.AddVertexAt(0, new Point2d(scale1 * 0.196473138321644, scale1 * 0.248628649148944), 0, 0, 0);
            poly166.AddVertexAt(1, new Point2d(scale1 * 0.196706916638733, scale1 * 0.243298503519296), 0, 0, 0);
            poly166.AddVertexAt(2, new Point2d(scale1 * 0.195070468419106, scale1 * 0.242036100607014), 0, 0, 0);
            poly166.AddVertexAt(3, new Point2d(scale1 * 0.19086245871149, scale1 * 0.240633430704475), 0, 0, 0);
            poly166.AddVertexAt(4, new Point2d(scale1 * 0.186888227320964, scale1 * 0.240773697694727), 0, 0, 0);
            poly166.AddVertexAt(5, new Point2d(scale1 * 0.183849109198798, scale1 * 0.242036100607014), 0, 0, 0);
            poly166.AddVertexAt(6, new Point2d(scale1 * 0.180576212759542, scale1 * 0.243298503519296), 0, 0, 0);
            poly166.AddVertexAt(7, new Point2d(scale1 * 0.177537094637375, scale1 * 0.244560906431583), 0, 0, 0);
            poly166.AddVertexAt(8, new Point2d(scale1 * 0.175666868100657, scale1 * 0.246945445265898), 0, 0, 0);
            poly166.AddVertexAt(9, new Point2d(scale1 * 0.175666868100657, scale1 * 0.248628649148944), 0, 0, 0);
            poly166.AddVertexAt(10, new Point2d(scale1 * 0.179173542857003, scale1 * 0.252275590895543), 0, 0, 0);
            poly166.AddVertexAt(11, new Point2d(scale1 * 0.185953114052605, scale1 * 0.253818527788335), 0, 0, 0);
            poly166.AddVertexAt(12, new Point2d(scale1 * 0.192732685248208, scale1 * 0.252275590895543), 0, 0, 0);
            poly166.Closed = true;
            poly166.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly166.Layer = "0";
            poly166.Color = color_GP;
            poly166.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly166);
            Polyline poly167 = new Polyline();
            poly167.AddVertexAt(0, new Point2d(scale1 * 0.162107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly167.AddVertexAt(1, new Point2d(scale1 * 0.159302385904375, scale1 * 0.241895833616759), 0, 0, 0);
            poly167.AddVertexAt(2, new Point2d(scale1 * 0.160237499172734, scale1 * 0.240072362743458), 0, 0, 0);
            poly167.AddVertexAt(3, new Point2d(scale1 * 0.161640169075273, scale1 * 0.237968357889651), 0, 0, 0);
            poly167.AddVertexAt(4, new Point2d(scale1 * 0.164679287197439, scale1 * 0.237687823909144), 0, 0, 0);
            poly167.AddVertexAt(5, new Point2d(scale1 * 0.168185961953785, scale1 * 0.23937102779219), 0, 0, 0);
            poly167.AddVertexAt(6, new Point2d(scale1 * 0.168185961953785, scale1 * 0.241615299636252), 0, 0, 0);
            poly167.AddVertexAt(7, new Point2d(scale1 * 0.165848178782888, scale1 * 0.244140105460821), 0, 0, 0);
            poly167.AddVertexAt(8, new Point2d(scale1 * 0.162107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly167.Closed = true;
            poly167.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly167.Layer = "0";
            poly167.Color = color_GP;
            poly167.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly167);
            Polyline poly168 = new Polyline();
            poly168.AddVertexAt(0, new Point2d(scale1 * 0.116520953876953, scale1 * 0.248628649148944), 0, 0, 0);
            poly168.AddVertexAt(1, new Point2d(scale1 * 0.123066746755466, scale1 * 0.250452120022245), 0, 0, 0);
            poly168.AddVertexAt(2, new Point2d(scale1 * 0.135924554195401, scale1 * 0.249750785070974), 0, 0, 0);
            poly168.AddVertexAt(3, new Point2d(scale1 * 0.142470347073914, scale1 * 0.248628649148944), 0, 0, 0);
            poly168.AddVertexAt(4, new Point2d(scale1 * 0.143639238659363, scale1 * 0.246805178275643), 0, 0, 0);
            poly168.AddVertexAt(5, new Point2d(scale1 * 0.143873016976452, scale1 * 0.244981707402343), 0, 0, 0);
            poly168.AddVertexAt(6, new Point2d(scale1 * 0.144106795293542, scale1 * 0.242737435558282), 0, 0, 0);
            poly168.AddVertexAt(7, new Point2d(scale1 * 0.14597702183026, scale1 * 0.240773697694727), 0, 0, 0);
            poly168.AddVertexAt(8, new Point2d(scale1 * 0.144574351927722, scale1 * 0.237407289928637), 0, 0, 0);
            poly168.AddVertexAt(9, new Point2d(scale1 * 0.139431228951747, scale1 * 0.235583819055336), 0, 0, 0);
            poly168.AddVertexAt(10, new Point2d(scale1 * 0.129378761316889, scale1 * 0.235303285074829), 0, 0, 0);
            poly168.AddVertexAt(11, new Point2d(scale1 * 0.124469416658004, scale1 * 0.236144887016352), 0, 0, 0);
            poly168.AddVertexAt(12, new Point2d(scale1 * 0.120962741901658, scale1 * 0.237828090899398), 0, 0, 0);
            poly168.AddVertexAt(13, new Point2d(scale1 * 0.117689845462402, scale1 * 0.241054231675236), 0, 0, 0);
            poly168.AddVertexAt(14, new Point2d(scale1 * 0.115819618925684, scale1 * 0.244420639441328), 0, 0, 0);
            poly168.Closed = true;
            poly168.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly168.Layer = "0";
            poly168.Color = color_GP;
            poly168.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly168);
            Polyline poly169 = new Polyline();
            poly169.AddVertexAt(0, new Point2d(scale1 * 0.0602857410017066, scale1 * 0.3), 0, 0, 0);
            poly169.AddVertexAt(1, new Point2d(scale1 * 0.0616401690752726, scale1 * 0.297968357889651), 0, 0, 0);
            poly169.AddVertexAt(2, new Point2d(scale1 * 0.0646792871974391, scale1 * 0.297687823909144), 0, 0, 0);
            poly169.AddVertexAt(3, new Point2d(scale1 * 0.0681859619537855, scale1 * 0.29937102779219), 0, 0, 0);
            poly169.AddVertexAt(4, new Point2d(scale1 * 0.0681859619537855, scale1 * 0.3), 0, 0, 0);
            poly169.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly169.Layer = "0";
            poly169.Color = color_GP;
            poly169.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly169);
            Polyline poly170 = new Polyline();
            poly170.AddVertexAt(0, new Point2d(scale1 * 0.0456546477907903, scale1 * 0.3), 0, 0, 0);
            poly170.AddVertexAt(1, new Point2d(scale1 * 0.0445743519277215, scale1 * 0.297407289928637), 0, 0, 0);
            poly170.AddVertexAt(2, new Point2d(scale1 * 0.0394312289517476, scale1 * 0.295583819055336), 0, 0, 0);
            poly170.AddVertexAt(3, new Point2d(scale1 * 0.0293787613168888, scale1 * 0.295303285074829), 0, 0, 0);
            poly170.AddVertexAt(4, new Point2d(scale1 * 0.024469416658004, scale1 * 0.296144887016352), 0, 0, 0);
            poly170.AddVertexAt(5, new Point2d(scale1 * 0.020962741901658, scale1 * 0.297828090899398), 0, 0, 0);
            poly170.AddVertexAt(6, new Point2d(scale1 * 0.0187593558575691, scale1 * 0.3), 0, 0, 0);
            poly170.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly170.Layer = "0";
            poly170.Color = color_GP;
            poly170.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly170);
            Polyline poly171 = new Polyline();
            poly171.AddVertexAt(0, new Point2d(scale1 * 0.0583672726360163, scale1 * 0.282994061761134), 0, 0, 0);
            poly171.AddVertexAt(1, new Point2d(scale1 * 0.0564970460992982, scale1 * 0.282853794770882), 0, 0, 0);
            poly171.AddVertexAt(2, new Point2d(scale1 * 0.0541592629284009, scale1 * 0.28355512972215), 0, 0, 0);
            poly171.AddVertexAt(3, new Point2d(scale1 * 0.052289036391683, scale1 * 0.285518867585705), 0, 0, 0);
            poly171.AddVertexAt(4, new Point2d(scale1 * 0.0539254846113113, scale1 * 0.287061804478495), 0, 0, 0);
            poly171.AddVertexAt(5, new Point2d(scale1 * 0.0583672726360163, scale1 * 0.286220202536974), 0, 0, 0);
            poly171.AddVertexAt(6, new Point2d(scale1 * 0.0595361642214651, scale1 * 0.284817532634435), 0, 0, 0);
            poly171.Closed = true;
            poly171.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly171.Layer = "0";
            poly171.Color = color_GP;
            poly171.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly171);
            Polyline poly172 = new Polyline();
            poly172.AddVertexAt(0, new Point2d(scale1 * 0.0679521836366956, scale1 * 0.28355512972215), 0, 0, 0);
            poly172.AddVertexAt(1, new Point2d(scale1 * 0.0660819570999778, scale1 * 0.282012192829359), 0, 0, 0);
            poly172.AddVertexAt(2, new Point2d(scale1 * 0.065848178782888, scale1 * 0.280048454965803), 0, 0, 0);
            poly172.AddVertexAt(3, new Point2d(scale1 * 0.0684197402708751, scale1 * 0.278505518073012), 0, 0, 0);
            poly172.AddVertexAt(4, new Point2d(scale1 * 0.0716926367101314, scale1 * 0.278926319043773), 0, 0, 0);
            poly172.AddVertexAt(5, new Point2d(scale1 * 0.0723939716614006, scale1 * 0.281030323897581), 0, 0, 0);
            poly172.AddVertexAt(6, new Point2d(scale1 * 0.0719264150272212, scale1 * 0.282713527780627), 0, 0, 0);
            poly172.AddVertexAt(7, new Point2d(scale1 * 0.0679521836366956, scale1 * 0.28355512972215), 0, 0, 0);
            poly172.Closed = true;
            poly172.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly172.Layer = "0";
            poly172.Color = color_GP;
            poly172.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly172);
            Polyline poly173 = new Polyline();
            poly173.AddVertexAt(0, new Point2d(scale1 * 0.016577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly173.AddVertexAt(1, new Point2d(scale1 * 0.0142393289561475, scale1 * 0.285251085149766), 0, 0, 0);
            poly173.AddVertexAt(2, new Point2d(scale1 * 0.015641998858686, scale1 * 0.283427614276465), 0, 0, 0);
            poly173.AddVertexAt(3, new Point2d(scale1 * 0.0196162302492116, scale1 * 0.282726279325197), 0, 0, 0);
            poly173.AddVertexAt(4, new Point2d(scale1 * 0.0226553483713783, scale1 * 0.283988682237481), 0, 0, 0);
            poly173.AddVertexAt(5, new Point2d(scale1 * 0.0226553483713783, scale1 * 0.286653755052303), 0, 0, 0);
            poly173.AddVertexAt(6, new Point2d(scale1 * 0.016577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly173.Closed = true;
            poly173.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly173.Layer = "0";
            poly173.Color = color_GP;
            poly173.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly173);
            Polyline poly174 = new Polyline();
            poly174.AddVertexAt(0, new Point2d(scale1 * 0.0110807582435306, scale1 * 0.266735842436258), 0, 0, 0);
            poly174.AddVertexAt(1, new Point2d(scale1 * 0.00947696180730939, scale1 * 0.263254670769046), 0, 0, 0);
            poly174.AddVertexAt(2, new Point2d(scale1 * 0.00971074012439921, scale1 * 0.259607729022446), 0, 0, 0);
            poly174.AddVertexAt(3, new Point2d(scale1 * 0.010645853392758, scale1 * 0.256802389217368), 0, 0, 0);
            poly174.AddVertexAt(4, new Point2d(scale1 * 0.0148538631003734, scale1 * 0.255960787275844), 0, 0, 0);
            poly174.AddVertexAt(5, new Point2d(scale1 * 0.0190618728079888, scale1 * 0.256521855236861), 0, 0, 0);
            poly174.AddVertexAt(6, new Point2d(scale1 * 0.0225685475643347, scale1 * 0.258625860090668), 0, 0, 0);
            poly174.AddVertexAt(7, new Point2d(scale1 * 0.0190618728079888, scale1 * 0.263535204749552), 0, 0, 0);
            poly174.AddVertexAt(8, new Point2d(scale1 * 0.0110807582435306, scale1 * 0.266735842436258), 0, 0, 0);
            poly174.Closed = true;
            poly174.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly174.Layer = "0";
            poly174.Color = color_GP;
            poly174.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly174);
            Polyline poly175 = new Polyline();
            poly175.AddVertexAt(0, new Point2d(scale1 * 0.0879754449631811, scale1 * 0.290759752403364), 0, 0, 0);
            poly175.AddVertexAt(1, new Point2d(scale1 * 0.0847025485239243, scale1 * 0.288936281530065), 0, 0, 0);
            poly175.AddVertexAt(2, new Point2d(scale1 * 0.0825985436701167, scale1 * 0.283746402890672), 0, 0, 0);
            poly175.AddVertexAt(3, new Point2d(scale1 * 0.0825985436701167, scale1 * 0.281221597066103), 0, 0, 0);
            poly175.AddVertexAt(4, new Point2d(scale1 * 0.0818972087188476, scale1 * 0.275751184446203), 0, 0, 0);
            poly175.AddVertexAt(5, new Point2d(scale1 * 0.0854038834751936, scale1 * 0.273366645611888), 0, 0, 0);
            poly175.AddVertexAt(6, new Point2d(scale1 * 0.0903132281340784, scale1 * 0.272525043670365), 0, 0, 0);
            poly175.AddVertexAt(7, new Point2d(scale1 * 0.101768365671475, scale1 * 0.272945844641126), 0, 0, 0);
            poly175.AddVertexAt(8, new Point2d(scale1 * 0.104807483793642, scale1 * 0.274488781533918), 0, 0, 0);
            poly175.AddVertexAt(9, new Point2d(scale1 * 0.105742597062001, scale1 * 0.278275990270772), 0, 0, 0);
            poly175.AddVertexAt(10, new Point2d(scale1 * 0.102703478939834, scale1 * 0.28037999512458), 0, 0, 0);
            poly175.AddVertexAt(11, new Point2d(scale1 * 0.0998981391347575, scale1 * 0.282764533958894), 0, 0, 0);
            poly175.AddVertexAt(12, new Point2d(scale1 * 0.09756035596386, scale1 * 0.286271208715241), 0, 0, 0);
            poly175.AddVertexAt(13, new Point2d(scale1 * 0.0956901294271417, scale1 * 0.288796014539811), 0, 0, 0);
            poly175.AddVertexAt(14, new Point2d(scale1 * 0.0914821197195268, scale1 * 0.289777883471588), 0, 0, 0);
            poly175.AddVertexAt(15, new Point2d(scale1 * 0.0879754449631811, scale1 * 0.290759752403364), 0, 0, 0);
            poly175.Closed = true;
            poly175.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly175.Layer = "0";
            poly175.Color = color_GP;
            poly175.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly175);
            Polyline poly176 = new Polyline();
            poly176.AddVertexAt(0, new Point2d(scale1 * 0.0583672726360163, scale1 * 0.274437775355651), 0, 0, 0);
            poly176.AddVertexAt(1, new Point2d(scale1 * 0.0660819570999778, scale1 * 0.272754571472604), 0, 0, 0);
            poly176.AddVertexAt(2, new Point2d(scale1 * 0.0693548535392339, scale1 * 0.268125760794228), 0, 0, 0);
            poly176.AddVertexAt(3, new Point2d(scale1 * 0.0660819570999778, scale1 * 0.263637217106105), 0, 0, 0);
            poly176.AddVertexAt(4, new Point2d(scale1 * 0.0583672726360163, scale1 * 0.261954013223059), 0, 0, 0);
            poly176.AddVertexAt(5, new Point2d(scale1 * 0.0497174749036959, scale1 * 0.261392945262043), 0, 0, 0);
            poly176.AddVertexAt(6, new Point2d(scale1 * 0.0450419085619012, scale1 * 0.26237481419382), 0, 0, 0);
            poly176.AddVertexAt(7, new Point2d(scale1 * 0.0413014554884654, scale1 * 0.26489962001839), 0, 0, 0);
            poly176.AddVertexAt(8, new Point2d(scale1 * 0.030781431219427, scale1 * 0.268967362735751), 0, 0, 0);
            poly176.AddVertexAt(9, new Point2d(scale1 * 0.0251707516092734, scale1 * 0.272894838462859), 0, 0, 0);
            poly176.AddVertexAt(10, new Point2d(scale1 * 0.0265734215118119, scale1 * 0.27471830933616), 0, 0, 0);
            poly176.AddVertexAt(11, new Point2d(scale1 * 0.0349894409270424, scale1 * 0.278365251082757), 0, 0, 0);
            poly176.AddVertexAt(12, new Point2d(scale1 * 0.041535233805555, scale1 * 0.279066586034028), 0, 0, 0);
            poly176.AddVertexAt(13, new Point2d(scale1 * 0.0457432435131704, scale1 * 0.277804183121743), 0, 0, 0);
            poly176.AddVertexAt(14, new Point2d(scale1 * 0.0513539231233242, scale1 * 0.276541780209458), 0, 0, 0);
            poly176.AddVertexAt(15, new Point2d(scale1 * 0.0557957111480292, scale1 * 0.275840445258188), 0, 0, 0);
            poly176.Closed = true;
            poly176.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly176.Layer = "0";
            poly176.Color = color_GP;
            poly176.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly176);
            Polyline poly177 = new Polyline();
            poly177.AddVertexAt(0, new Point2d(scale1 * 0.0788988979118168, scale1 * 0.261890255500208), 0, 0, 0);
            poly177.AddVertexAt(1, new Point2d(scale1 * 0.0727410266071835, scale1 * 0.262017770945895), 0, 0, 0);
            poly177.AddVertexAt(2, new Point2d(scale1 * 0.0674325168618095, scale1 * 0.259722492923566), 0, 0, 0);
            poly177.AddVertexAt(3, new Point2d(scale1 * 0.0663708149127362, scale1 * 0.256662122227115), 0, 0, 0);
            poly177.AddVertexAt(4, new Point2d(scale1 * 0.07167932465811, scale1 * 0.255769514107321), 0, 0, 0);
            poly177.AddVertexAt(5, new Point2d(scale1 * 0.0782618767423733, scale1 * 0.2567896376728), 0, 0, 0);
            poly177.AddVertexAt(6, new Point2d(scale1 * 0.0820840037590416, scale1 * 0.259467462032192), 0, 0, 0);
            poly177.AddVertexAt(7, new Point2d(scale1 * 0.0788988979118168, scale1 * 0.261890255500208), 0, 0, 0);
            poly177.Closed = true;
            poly177.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly177.Layer = "0";
            poly177.Color = color_GP;
            poly177.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly177);
            Polyline poly178 = new Polyline();
            poly178.AddVertexAt(0, new Point2d(scale1 * 0.0964731383216439, scale1 * 0.248628649148944), 0, 0, 0);
            poly178.AddVertexAt(1, new Point2d(scale1 * 0.0967069166387331, scale1 * 0.243298503519296), 0, 0, 0);
            poly178.AddVertexAt(2, new Point2d(scale1 * 0.0950704684191055, scale1 * 0.242036100607014), 0, 0, 0);
            poly178.AddVertexAt(3, new Point2d(scale1 * 0.0908624587114899, scale1 * 0.240633430704475), 0, 0, 0);
            poly178.AddVertexAt(4, new Point2d(scale1 * 0.0868882273209644, scale1 * 0.240773697694727), 0, 0, 0);
            poly178.AddVertexAt(5, new Point2d(scale1 * 0.0838491091987978, scale1 * 0.242036100607014), 0, 0, 0);
            poly178.AddVertexAt(6, new Point2d(scale1 * 0.0805762127595415, scale1 * 0.243298503519296), 0, 0, 0);
            poly178.AddVertexAt(7, new Point2d(scale1 * 0.0775370946373748, scale1 * 0.244560906431583), 0, 0, 0);
            poly178.AddVertexAt(8, new Point2d(scale1 * 0.0756668681006569, scale1 * 0.246945445265898), 0, 0, 0);
            poly178.AddVertexAt(9, new Point2d(scale1 * 0.0756668681006569, scale1 * 0.248628649148944), 0, 0, 0);
            poly178.AddVertexAt(10, new Point2d(scale1 * 0.0791735428570031, scale1 * 0.252275590895543), 0, 0, 0);
            poly178.AddVertexAt(11, new Point2d(scale1 * 0.0859531140526055, scale1 * 0.253818527788335), 0, 0, 0);
            poly178.AddVertexAt(12, new Point2d(scale1 * 0.0927326852482082, scale1 * 0.252275590895543), 0, 0, 0);
            poly178.Closed = true;
            poly178.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly178.Layer = "0";
            poly178.Color = color_GP;
            poly178.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly178);
            Polyline poly179 = new Polyline();
            poly179.AddVertexAt(0, new Point2d(scale1 * 0.062107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly179.AddVertexAt(1, new Point2d(scale1 * 0.0593023859043753, scale1 * 0.241895833616759), 0, 0, 0);
            poly179.AddVertexAt(2, new Point2d(scale1 * 0.0602374991727341, scale1 * 0.240072362743458), 0, 0, 0);
            poly179.AddVertexAt(3, new Point2d(scale1 * 0.0616401690752726, scale1 * 0.237968357889651), 0, 0, 0);
            poly179.AddVertexAt(4, new Point2d(scale1 * 0.0646792871974391, scale1 * 0.237687823909144), 0, 0, 0);
            poly179.AddVertexAt(5, new Point2d(scale1 * 0.0681859619537855, scale1 * 0.23937102779219), 0, 0, 0);
            poly179.AddVertexAt(6, new Point2d(scale1 * 0.0681859619537855, scale1 * 0.241615299636252), 0, 0, 0);
            poly179.AddVertexAt(7, new Point2d(scale1 * 0.065848178782888, scale1 * 0.244140105460821), 0, 0, 0);
            poly179.AddVertexAt(8, new Point2d(scale1 * 0.062107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly179.Closed = true;
            poly179.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly179.Layer = "0";
            poly179.Color = color_GP;
            poly179.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly179);
            Polyline poly180 = new Polyline();
            poly180.AddVertexAt(0, new Point2d(scale1 * 0.0165209538769528, scale1 * 0.248628649148944), 0, 0, 0);
            poly180.AddVertexAt(1, new Point2d(scale1 * 0.0230667467554657, scale1 * 0.250452120022245), 0, 0, 0);
            poly180.AddVertexAt(2, new Point2d(scale1 * 0.0359245541954014, scale1 * 0.249750785070974), 0, 0, 0);
            poly180.AddVertexAt(3, new Point2d(scale1 * 0.0424703470739138, scale1 * 0.248628649148944), 0, 0, 0);
            poly180.AddVertexAt(4, new Point2d(scale1 * 0.0436392386593627, scale1 * 0.246805178275643), 0, 0, 0);
            poly180.AddVertexAt(5, new Point2d(scale1 * 0.0438730169764525, scale1 * 0.244981707402343), 0, 0, 0);
            poly180.AddVertexAt(6, new Point2d(scale1 * 0.0441067952935421, scale1 * 0.242737435558282), 0, 0, 0);
            poly180.AddVertexAt(7, new Point2d(scale1 * 0.0459770218302602, scale1 * 0.240773697694727), 0, 0, 0);
            poly180.AddVertexAt(8, new Point2d(scale1 * 0.0445743519277215, scale1 * 0.237407289928637), 0, 0, 0);
            poly180.AddVertexAt(9, new Point2d(scale1 * 0.0394312289517476, scale1 * 0.235583819055336), 0, 0, 0);
            poly180.AddVertexAt(10, new Point2d(scale1 * 0.0293787613168888, scale1 * 0.235303285074829), 0, 0, 0);
            poly180.AddVertexAt(11, new Point2d(scale1 * 0.024469416658004, scale1 * 0.236144887016352), 0, 0, 0);
            poly180.AddVertexAt(12, new Point2d(scale1 * 0.020962741901658, scale1 * 0.237828090899398), 0, 0, 0);
            poly180.AddVertexAt(13, new Point2d(scale1 * 0.0176898454624017, scale1 * 0.241054231675236), 0, 0, 0);
            poly180.AddVertexAt(14, new Point2d(scale1 * 0.0158196189256838, scale1 * 0.244420639441328), 0, 0, 0);
            poly180.Closed = true;
            poly180.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly180.Layer = "0";
            poly180.Color = color_GP;
            poly180.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly180);

            #endregion


        }


        private void add_pattern_SP(BlockTableRecord bltrec1, double scale1, double graph_vexag, double stick_vexag, Polyline poly1, BlockTableRecord BTrecord, Autodesk.AutoCAD.DatabaseServices.Transaction Trans1)
        {


            Autodesk.AutoCAD.Colors.Color color1 = Autodesk.AutoCAD.Colors.Color.FromRgb(0, 255, 0);




            int nr_rows = 0;
            int nr_col = 0;

            double Xmax = 0.177187500001552 * scale1;
            double Ymax = 0.0645833333333333 * scale1;
            double r1 = 0.0015 * scale1;



            double x1 = poly1.GetPoint2dAt(3).X;
            double y1 = poly1.GetPoint2dAt(3).Y;
            double stick_width = poly1.GetPoint2dAt(1).X - poly1.GetPoint2dAt(0).X;
            double stick_height = poly1.GetPoint2dAt(1).Y - poly1.GetPoint2dAt(2).Y;

            if (stick_height < Ymax)
            {
                nr_rows = 1;
            }
            else
            {
                double nr2 = Math.Floor(stick_height / Ymax);
                nr_rows = Convert.ToInt32(nr2);
            }


            if (stick_width < Xmax)
            {
                nr_col = 1;
            }
            else
            {
                double nr2 = Math.Floor(stick_width / Xmax);
                nr_col = Convert.ToInt32(nr2);
            }








            for (int m = 0; m < nr_col; ++m)
            {
                for (int n = 0; n < nr_rows; ++n)
                {

                    double x2 = x1;
                    double y2 = y1;



                    Circle cerc1 = new Circle(new Point3d(0.0771875000015522 * scale1 + m * Xmax, 0.00208333333333333 * scale1 + n * Ymax, 0), Vector3d.ZAxis, r1);
                    cerc1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(x2, y2, 0))));
                    BTrecord.AppendEntity(cerc1);
                    Trans1.AddNewlyCreatedDBObject(cerc1, true);
                    Hatch hatch1 = CreateHatch(cerc1, "SOLID", 1, 0);
                    hatch1.Layer = "0";
                    hatch1.LineWeight = LineWeight.LineWeight000;
                    hatch1.Color = color1;
                    bltrec1.AppendEntity(hatch1);
                    cerc1.Erase();

                    cerc1 = new Circle(new Point3d(0.108497459879921 * scale1 + m * Xmax, 0.0111742424157759 * scale1 + n * Ymax, 0), Vector3d.ZAxis, r1);
                    cerc1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(x2, y2, 0))));
                    BTrecord.AppendEntity(cerc1);
                    Trans1.AddNewlyCreatedDBObject(cerc1, true);
                    hatch1 = CreateHatch(cerc1, "SOLID", 1, 0);
                    hatch1.Layer = "0";
                    hatch1.LineWeight = LineWeight.LineWeight000;
                    hatch1.Color = color1;
                    bltrec1.AppendEntity(hatch1);
                    cerc1.Erase();

                    cerc1 = new Circle(new Point3d(0.139687500001552 * scale1 + m * Xmax, 0.0395833333333333 * scale1 + n * Ymax, 0), Vector3d.ZAxis, r1);
                    cerc1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(x2, y2, 0))));
                    BTrecord.AppendEntity(cerc1);
                    Trans1.AddNewlyCreatedDBObject(cerc1, true);
                    hatch1 = CreateHatch(cerc1, "SOLID", 1, 0);
                    hatch1.Layer = "0";
                    hatch1.LineWeight = LineWeight.LineWeight000;
                    hatch1.Color = color1;
                    bltrec1.AppendEntity(hatch1);
                    cerc1.Erase();

                    cerc1 = new Circle(new Point3d(0.177187500001552 * scale1 + m * Xmax, 0.0520833333333333 * scale1 + n * Ymax, 0), Vector3d.ZAxis, r1);
                    cerc1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(x2, y2, 0))));
                    BTrecord.AppendEntity(cerc1);
                    Trans1.AddNewlyCreatedDBObject(cerc1, true);
                    hatch1 = CreateHatch(cerc1, "SOLID", 1, 0);
                    hatch1.Layer = "0";
                    hatch1.LineWeight = LineWeight.LineWeight000;
                    hatch1.Color = color1;
                    bltrec1.AppendEntity(hatch1);
                    cerc1.Erase();

                    cerc1 = new Circle(new Point3d(0.164687500001552 * scale1 + m * Xmax, 0.0145833333333333 * scale1 + n * Ymax, 0), Vector3d.ZAxis, r1);
                    cerc1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(x2, y2, 0))));
                    BTrecord.AppendEntity(cerc1);
                    Trans1.AddNewlyCreatedDBObject(cerc1, true);
                    hatch1 = CreateHatch(cerc1, "SOLID", 1, 0);
                    hatch1.Layer = "0";
                    hatch1.LineWeight = LineWeight.LineWeight000;
                    hatch1.Color = color1;
                    bltrec1.AppendEntity(hatch1);
                    cerc1.Erase();

                    cerc1 = new Circle(new Point3d(0.102187500001552 * scale1 + m * Xmax, 0.0395833333333333 * scale1 + n * Ymax, 0), Vector3d.ZAxis, r1);
                    cerc1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(x2, y2, 0))));
                    BTrecord.AppendEntity(cerc1);
                    Trans1.AddNewlyCreatedDBObject(cerc1, true);
                    hatch1 = CreateHatch(cerc1, "SOLID", 1, 0);
                    hatch1.Layer = "0";
                    hatch1.LineWeight = LineWeight.LineWeight000;
                    hatch1.Color = color1;
                    bltrec1.AppendEntity(hatch1);
                    cerc1.Erase();

                    cerc1 = new Circle(new Point3d(0.114687500001552 * scale1 + m * Xmax, 0.0645833333333333 * scale1 + n * Ymax, 0), Vector3d.ZAxis, r1);
                    cerc1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(x2, y2, 0))));
                    BTrecord.AppendEntity(cerc1);
                    Trans1.AddNewlyCreatedDBObject(cerc1, true);
                    hatch1 = CreateHatch(cerc1, "SOLID", 1, 0);
                    hatch1.Layer = "0";
                    hatch1.LineWeight = LineWeight.LineWeight000;
                    hatch1.Color = color1;
                    bltrec1.AppendEntity(hatch1);
                    cerc1.Erase();

                    cerc1 = new Circle(new Point3d(0.0857900086014221 * scale1 + m * Xmax, 0.0532196969725192 * scale1 + n * Ymax, 0), Vector3d.ZAxis, r1);
                    cerc1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(x2, y2, 0))));
                    BTrecord.AppendEntity(cerc1);
                    Trans1.AddNewlyCreatedDBObject(cerc1, true);
                    hatch1 = CreateHatch(cerc1, "SOLID", 1, 0);
                    hatch1.Layer = "0";
                    hatch1.LineWeight = LineWeight.LineWeight000;
                    hatch1.Color = color1;
                    bltrec1.AppendEntity(hatch1);
                    cerc1.Erase();

                    cerc1 = new Circle(new Point3d(0.0521875000015522 * scale1 + m * Xmax, 0.0270833333333333 * scale1 + n * Ymax, 0), Vector3d.ZAxis, r1);
                    cerc1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(x2, y2, 0))));
                    BTrecord.AppendEntity(cerc1);
                    Trans1.AddNewlyCreatedDBObject(cerc1, true);
                    hatch1 = CreateHatch(cerc1, "SOLID", 1, 0);
                    hatch1.Layer = "0";
                    hatch1.LineWeight = LineWeight.LineWeight000;
                    hatch1.Color = color1;
                    bltrec1.AppendEntity(hatch1);
                    cerc1.Erase();

                    cerc1 = new Circle(new Point3d(0.0146875000015522 * scale1 + m * Xmax, 0.0145833333333333 * scale1 + n * Ymax, 0), Vector3d.ZAxis, r1);
                    cerc1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(x2, y2, 0))));
                    BTrecord.AppendEntity(cerc1);
                    Trans1.AddNewlyCreatedDBObject(cerc1, true);
                    hatch1 = CreateHatch(cerc1, "SOLID", 1, 0);
                    hatch1.Layer = "0";
                    hatch1.LineWeight = LineWeight.LineWeight000;
                    hatch1.Color = color1;
                    bltrec1.AppendEntity(hatch1);
                    cerc1.Erase();

                    cerc1 = new Circle(new Point3d(0.0271875000015522 * scale1 + m * Xmax, 0.0395833333333333 * scale1 + n * Ymax, 0), Vector3d.ZAxis, r1);
                    cerc1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(x2, y2, 0))));
                    BTrecord.AppendEntity(cerc1);
                    Trans1.AddNewlyCreatedDBObject(cerc1, true);
                    hatch1 = CreateHatch(cerc1, "SOLID", 1, 0);
                    hatch1.Layer = "0";
                    hatch1.LineWeight = LineWeight.LineWeight000;
                    hatch1.Color = color1;
                    bltrec1.AppendEntity(hatch1);
                    cerc1.Erase();

                    cerc1 = new Circle(new Point3d(0.0396875000015522 * scale1 + m * Xmax, 0.0645833333333333 * scale1 + n * Ymax, 0), Vector3d.ZAxis, r1);
                    cerc1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(x2, y2, 0))));
                    BTrecord.AppendEntity(cerc1);
                    Trans1.AddNewlyCreatedDBObject(cerc1, true);
                    hatch1 = CreateHatch(cerc1, "SOLID", 1, 0);
                    hatch1.Layer = "0";
                    hatch1.LineWeight = LineWeight.LineWeight000;
                    hatch1.Color = color1;
                    bltrec1.AppendEntity(hatch1);
                    cerc1.Erase();

                    cerc1 = new Circle(new Point3d(0.0021875000015522 * scale1 + m * Xmax, 0.0520833333333333 * scale1 + n * Ymax, 0), Vector3d.ZAxis, r1);
                    cerc1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(x2, y2, 0))));
                    BTrecord.AppendEntity(cerc1);
                    Trans1.AddNewlyCreatedDBObject(cerc1, true);
                    hatch1 = CreateHatch(cerc1, "SOLID", 1, 0);
                    hatch1.Layer = "0";
                    hatch1.LineWeight = LineWeight.LineWeight000;
                    hatch1.Color = color1;
                    bltrec1.AppendEntity(hatch1);
                    cerc1.Erase();





                }
            }

        }





        private void add_pattern_Shale(BlockTableRecord bltrec1, double scale1, double graph_vexag, double stick_vexag, Polyline poly1, BlockTableRecord BTrecord, Autodesk.AutoCAD.DatabaseServices.Transaction Trans1)
        {


            Autodesk.AutoCAD.Colors.Color color1 = Autodesk.AutoCAD.Colors.Color.FromRgb(0, 255, 0);




            int nr_rows = 0;

            double spc_h_edge = 0;
            double spc_v_edge = 0;
            double spc_ver = scale1 * 0.009;
            double hole = scale1 * 0.036;

            double x1 = poly1.GetPoint2dAt(3).X + spc_h_edge;
            double y1 = poly1.GetPoint2dAt(3).Y + spc_v_edge;
            double stick_width = poly1.GetPoint2dAt(1).X - poly1.GetPoint2dAt(0).X;
            double stick_height = poly1.GetPoint2dAt(1).Y - poly1.GetPoint2dAt(2).Y;

            if (stick_height < spc_ver)
            {
                nr_rows = 1;
            }
            else
            {
                double nr2 = Math.Floor(stick_height / spc_ver);
                nr_rows = Convert.ToInt32(nr2) - 1;
            }





            double dif_len_v = stick_height - ((nr_rows - 1) * spc_ver);



            if (nr_rows > 0 && stick_width > 4 * hole)
            {
                Polyline poly_left_right = new Polyline();
                int line_no = 1;
                bool draw_second = false;

                for (int n = 0; n < nr_rows; ++n)
                {

                    double y2 = y1 + n * spc_ver + dif_len_v / 2;
                    double y3 = y2;
                    double x2 = x1;
                    double x3 = x1 + stick_width;
                    double y4 = y3;
                    double y5 = y3;
                    double x4 = x1;
                    double x5 = x1 + stick_width;


                    switch (line_no)
                    {
                        case 1:
                            x2 = x1 + hole;
                            x3 = x1 + stick_width;
                            draw_second = false;
                            break;
                        case 2:
                            x2 = x1;
                            x3 = x1 + hole;
                            draw_second = true;
                            break;
                        case 3:
                            x2 = x1;
                            x3 = x1 + 2 * hole;
                            draw_second = true;
                            break;
                        case 4:
                            x2 = x1;
                            x3 = x1 + stick_width - hole;
                            draw_second = false;
                            break;
                        default:
                            break;
                    }

                    ++line_no;
                    if (line_no == 5) line_no = 1;

                    poly_left_right = new Polyline();
                    poly_left_right.AddVertexAt(0, new Point2d(x2, y2), 0, 0, 0);
                    poly_left_right.AddVertexAt(1, new Point2d(x3, y3), 0, 0, 0);
                    poly_left_right.Layer = "0";
                    poly_left_right.Color = color1;
                    poly_left_right.LineWeight = LineWeight.LineWeight000;
                    bltrec1.AppendEntity(poly_left_right);

                    if (draw_second == true)
                    {
                        x4 = x3 + hole;
                        x5 = x1 + stick_width;
                        poly_left_right = new Polyline();
                        poly_left_right.AddVertexAt(0, new Point2d(x4, y4), 0, 0, 0);
                        poly_left_right.AddVertexAt(1, new Point2d(x5, y5), 0, 0, 0);
                        poly_left_right.Layer = "0";
                        poly_left_right.Color = color1;
                        poly_left_right.LineWeight = LineWeight.LineWeight000;
                        bltrec1.AppendEntity(poly_left_right);
                    }

                }
            }
        }


        private void add_pattern_Shale_legend(BlockTableRecord bltrec1, double scale1, double graph_vexag, double stick_vexag, Polyline poly1, BlockTableRecord BTrecord, Autodesk.AutoCAD.DatabaseServices.Transaction Trans1)
        {


            Autodesk.AutoCAD.Colors.Color color1 = Autodesk.AutoCAD.Colors.Color.FromRgb(0, 255, 0);




            int nr_rows = 32;
            int nr_col = 10;

            double spc_ver = 0.009;
            double hole = 0.036;

            double x1 = poly1.GetPoint2dAt(3).X;
            double ins_y = poly1.GetPoint2dAt(3).Y + 0.0105;
            double xtra1 = 0.024;
            double xtra3 = 0.096;
            double xtra4 = 0.06;



            Polyline poly_left_right = new Polyline();

            int row_no = 1;

            for (int n = 0; n < nr_rows; ++n)
            {
                double x2 = 0;
                double x3 = 0;

                double y2 = ins_y + n * spc_ver;
                double y3 = y2;



                switch (row_no)
                {
                    case 1:

                        for (int m = 0; m < nr_col; m++)
                        {

                            x2 = x1 + hole + m * 4 * hole;
                            x3 = x2 + 3 * hole;

                            poly_left_right = new Polyline();
                            poly_left_right.AddVertexAt(0, new Point2d(x2, y2), 0, 0, 0);
                            poly_left_right.AddVertexAt(1, new Point2d(x3, y3), 0, 0, 0);
                            poly_left_right.Layer = "0";
                            poly_left_right.Color = color1;
                            poly_left_right.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(poly_left_right);

                        }
                        poly_left_right = new Polyline();
                        poly_left_right.AddVertexAt(0, new Point2d(x3 + hole, y3), 0, 0, 0);
                        poly_left_right.AddVertexAt(1, new Point2d(x3 + hole + xtra1, y3), 0, 0, 0);
                        poly_left_right.Layer = "0";
                        poly_left_right.Color = color1;
                        poly_left_right.LineWeight = LineWeight.LineWeight000;
                        bltrec1.AppendEntity(poly_left_right);



                        break;
                    case 2:
                        for (int m = 0; m < nr_col; m++)
                        {

                            x2 = x1 + 2 * hole + m * 4 * hole;
                            x3 = x2 + 3 * hole;

                            poly_left_right = new Polyline();
                            poly_left_right.AddVertexAt(0, new Point2d(x2, y2), 0, 0, 0);
                            poly_left_right.AddVertexAt(1, new Point2d(x3, y3), 0, 0, 0);
                            poly_left_right.Layer = "0";
                            poly_left_right.Color = color1;
                            poly_left_right.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(poly_left_right);
                        }
                        poly_left_right = new Polyline();
                        poly_left_right.AddVertexAt(0, new Point2d(x1, y3), 0, 0, 0);
                        poly_left_right.AddVertexAt(1, new Point2d(x1 + hole, y3), 0, 0, 0);
                        poly_left_right.Layer = "0";
                        poly_left_right.Color = color1;
                        poly_left_right.LineWeight = LineWeight.LineWeight000;
                        bltrec1.AppendEntity(poly_left_right);



                        break;
                    case 3:
                        for (int m = 0; m < nr_col - 1; m++)
                        {

                            x2 = x1 + 3 * hole + m * 4 * hole;
                            x3 = x2 + 3 * hole;

                            poly_left_right = new Polyline();
                            poly_left_right.AddVertexAt(0, new Point2d(x2, y2), 0, 0, 0);
                            poly_left_right.AddVertexAt(1, new Point2d(x3, y3), 0, 0, 0);
                            poly_left_right.Layer = "0";
                            poly_left_right.Color = color1;
                            poly_left_right.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(poly_left_right);
                        }
                        poly_left_right = new Polyline();
                        poly_left_right.AddVertexAt(0, new Point2d(x1, y3), 0, 0, 0);
                        poly_left_right.AddVertexAt(1, new Point2d(x1 + 2 * hole, y3), 0, 0, 0);
                        poly_left_right.Layer = "0";
                        poly_left_right.Color = color1;
                        poly_left_right.LineWeight = LineWeight.LineWeight000;
                        bltrec1.AppendEntity(poly_left_right);

                        poly_left_right = new Polyline();
                        poly_left_right.AddVertexAt(0, new Point2d(x3 + hole, y3), 0, 0, 0);
                        poly_left_right.AddVertexAt(1, new Point2d(x3 + hole + xtra3, y3), 0, 0, 0);
                        poly_left_right.Layer = "0";
                        poly_left_right.Color = color1;
                        poly_left_right.LineWeight = LineWeight.LineWeight000;
                        bltrec1.AppendEntity(poly_left_right);

                        break;
                    case 4:
                        for (int m = 0; m < nr_col; m++)
                        {

                            x2 = x1 + m * 4 * hole;
                            x3 = x2 + 3 * hole;

                            poly_left_right = new Polyline();
                            poly_left_right.AddVertexAt(0, new Point2d(x2, y2), 0, 0, 0);
                            poly_left_right.AddVertexAt(1, new Point2d(x3, y3), 0, 0, 0);
                            poly_left_right.Layer = "0";
                            poly_left_right.Color = color1;
                            poly_left_right.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(poly_left_right);
                        }
                        poly_left_right = new Polyline();
                        poly_left_right.AddVertexAt(0, new Point2d(x3 + hole, y3), 0, 0, 0);
                        poly_left_right.AddVertexAt(1, new Point2d(x3 + hole + xtra4, y3), 0, 0, 0);
                        poly_left_right.Layer = "0";
                        poly_left_right.Color = color1;
                        poly_left_right.LineWeight = LineWeight.LineWeight000;
                        bltrec1.AppendEntity(poly_left_right);

                        break;
                    default:
                        break;
                }

                ++row_no;
                if (row_no == 5) row_no = 1;







            }
        }



        private void add_pattern_GM(BlockTableRecord bltrec1, double scale1, double graph_vexag, double stick_vexag, Polyline poly1, BlockTableRecord BTrecord, Autodesk.AutoCAD.DatabaseServices.Transaction Trans1)
        {



            Autodesk.AutoCAD.Colors.Color color_gm = Autodesk.AutoCAD.Colors.Color.FromRgb(127, 95, 93);



            double pattern_width = 0.18 * scale1;
            double pattern_height = 0.11 * scale1;

            double spc_h_edge = 0;
            double spc_v_edge = 0;
            double spc_hor = 0;
            double spc_ver = 0;

            int nr_col = 0;
            int nr_rows = 1;

            double x1 = poly1.GetPoint2dAt(3).X + spc_h_edge;
            double y1 = poly1.GetPoint2dAt(3).Y + spc_v_edge;


            double stick_width = poly1.GetPoint2dAt(1).X - poly1.GetPoint2dAt(0).X;
            double rectangle_height = poly1.GetPoint2dAt(1).Y - poly1.GetPoint2dAt(2).Y;

            if (rectangle_height >= pattern_height + 2 * spc_v_edge)
            {
                double nr1 = Math.Floor((rectangle_height - 2 * spc_v_edge) / (pattern_height + spc_ver));
                nr_rows = Convert.ToInt32(nr1);

                if (stick_width - 2 * spc_h_edge < pattern_width + 2 * spc_hor)
                {
                    nr_col = 1;
                }
                else
                {
                    double nr2 = Math.Floor((stick_width - 2 * spc_h_edge) / (pattern_width + spc_hor));

                    nr_col = Convert.ToInt32(nr2);
                }

                double dif_len = stick_width - (nr_col * (pattern_width + spc_hor) - spc_hor);
                double dif_hght = rectangle_height - (nr_rows * (pattern_height + spc_ver) - spc_ver);

                Point3d pt_ins = poly1.GetPoint3dAt(3);

                if (nr_rows > 0 && nr_col > 0)
                {
                    for (int m = 0; m < nr_col; ++m)
                    {
                        for (int n = 0; n < nr_rows; ++n)
                        {
                            double x2 = x1 + dif_len / 2 + m * (pattern_width + spc_hor);
                            double y2 = y1 + dif_hght / 2 + n * (pattern_height + spc_ver);

                            double x3 = x2 + pattern_width;
                            double y3 = y2;


                            Polyline polygm1 = get_poly_gm1(scale1);
                            polygm1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));
                            polygm1.Layer = "0";
                            polygm1.Color = color_gm;
                            polygm1.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygm1);

                            Polyline polygm2 = get_poly_gm2(scale1);
                            polygm2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygm2.Layer = "0";
                            polygm2.Color = color_gm;
                            polygm2.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygm2);

                            Polyline polygm3 = get_poly_gm3(scale1);
                            polygm3.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygm3.Layer = "0";
                            polygm3.Color = color_gm;
                            polygm3.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygm3);

                            Polyline polygm4 = get_poly_gm4(scale1);
                            polygm4.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygm4.Layer = "0";
                            polygm4.Color = color_gm;
                            polygm4.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygm4);

                            Polyline polygm5 = get_poly_gm5(scale1);
                            polygm5.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygm5.Layer = "0";
                            polygm5.Color = color_gm;
                            polygm5.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygm5);

                            Polyline polygm6 = get_poly_gm6(scale1);
                            polygm6.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygm6.Layer = "0";
                            polygm6.Color = color_gm;
                            polygm6.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygm6);

                            Polyline polygm7 = get_poly_gm7(scale1);
                            polygm7.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygm7.Layer = "0";
                            polygm7.Color = color_gm;
                            polygm7.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygm7);

                            Polyline polygm8 = get_poly_gm8(scale1);
                            polygm8.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygm8.Layer = "0";
                            polygm8.Color = color_gm;
                            polygm8.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygm8);

                            Polyline polygm9 = get_poly_gm9(scale1);
                            polygm9.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygm9.Layer = "0";
                            polygm9.Color = color_gm;
                            polygm9.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygm9);

                            Polyline polygm10 = get_poly_gm10(scale1);
                            polygm10.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygm10.Layer = "0";
                            polygm10.Color = color_gm;
                            polygm10.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygm10);

                        }
                    }



                }
            }
            else
            {
            }


        }



        private void add_pattern_GM_legend(BlockTableRecord bltrec1, double scale1, double graph_vexag, double stick_vexag, Polyline poly1, BlockTableRecord BTrecord, Autodesk.AutoCAD.DatabaseServices.Transaction Trans1)
        {



            Autodesk.AutoCAD.Colors.Color color_gm = Autodesk.AutoCAD.Colors.Color.FromRgb(127, 95, 93);



            double pattern_width = 0.18 * scale1;
            double pattern_height = 0.11 * scale1;

            double spc_h_edge = 0;
            double spc_v_edge = 0;
            double spc_hor = 0;
            double spc_ver = 0;

            int nr_col = 0;
            int nr_rows = 1;

            double x1 = poly1.GetPoint2dAt(3).X + spc_h_edge;
            double y1 = poly1.GetPoint2dAt(3).Y + spc_v_edge;


            double stick_width = poly1.GetPoint2dAt(1).X - poly1.GetPoint2dAt(0).X;
            double rectangle_height = poly1.GetPoint2dAt(1).Y - poly1.GetPoint2dAt(2).Y;

            if (rectangle_height >= pattern_height + 2 * spc_v_edge)
            {
                double nr1 = Math.Floor((rectangle_height - 2 * spc_v_edge) / (pattern_height + spc_ver));
                nr_rows = Convert.ToInt32(nr1);

                if (stick_width - 2 * spc_h_edge < pattern_width + 2 * spc_hor)
                {
                    nr_col = 1;
                }
                else
                {
                    double nr2 = Math.Floor((stick_width - 2 * spc_h_edge) / (pattern_width + spc_hor));

                    nr_col = Convert.ToInt32(nr2);
                }

                double dif_len = stick_width - (nr_col * (pattern_width + spc_hor) - spc_hor);
                double dif_hght = rectangle_height - (nr_rows * (pattern_height + spc_ver) - spc_ver);

                Point3d pt_ins = poly1.GetPoint3dAt(3);

                if (nr_rows > 0 && nr_col > 0)
                {
                    for (int m = 0; m < nr_col; ++m)
                    {
                        for (int n = 0; n < nr_rows; ++n)
                        {
                            double x2 = x1 + dif_len / 2 + m * (pattern_width + spc_hor);
                            double y2 = y1 + dif_hght / 2 + n * (pattern_height + spc_ver);

                            double x3 = x2 + pattern_width;
                            double y3 = y2;


                            Polyline polygm1 = get_poly_gm1(scale1);
                            polygm1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));
                            polygm1.Layer = "0";
                            polygm1.Color = color_gm;
                            polygm1.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygm1);

                            Polyline polygm2 = get_poly_gm2(scale1);
                            polygm2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygm2.Layer = "0";
                            polygm2.Color = color_gm;
                            polygm2.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygm2);

                            Polyline polygm3 = get_poly_gm3(scale1);
                            polygm3.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygm3.Layer = "0";
                            polygm3.Color = color_gm;
                            polygm3.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygm3);

                            Polyline polygm4 = get_poly_gm4(scale1);
                            polygm4.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygm4.Layer = "0";
                            polygm4.Color = color_gm;
                            polygm4.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygm4);

                            Polyline polygm5 = get_poly_gm5(scale1);
                            polygm5.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygm5.Layer = "0";
                            polygm5.Color = color_gm;
                            polygm5.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygm5);

                            Polyline polygm6 = get_poly_gm6(scale1);
                            polygm6.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygm6.Layer = "0";
                            polygm6.Color = color_gm;
                            polygm6.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygm6);

                            Polyline polygm7 = get_poly_gm7(scale1);
                            polygm7.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygm7.Layer = "0";
                            polygm7.Color = color_gm;
                            polygm7.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygm7);

                            Polyline polygm8 = get_poly_gm8(scale1);
                            polygm8.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygm8.Layer = "0";
                            polygm8.Color = color_gm;
                            polygm8.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygm8);

                            Polyline polygm9 = get_poly_gm9(scale1);
                            polygm9.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygm9.Layer = "0";
                            polygm9.Color = color_gm;
                            polygm9.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygm9);

                            Polyline polygm10 = get_poly_gm10(scale1);
                            polygm10.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygm10.Layer = "0";
                            polygm10.Color = color_gm;
                            polygm10.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygm10);

                        }
                    }



                }
            }
            else
            {
            }


            #region polyextra

            Polyline polyextra1 = new Polyline();
            polyextra1.AddVertexAt(0, new Point2d(scale1 * 1.49617536339454, scale1 * 0.3), 0, 0, 0);
            polyextra1.AddVertexAt(1, new Point2d(scale1 * 1.4979215182454, scale1 * 0.297485535959897), 0, 0, 0);
            polyextra1.AddVertexAt(2, new Point2d(scale1 * 1.49587551722873, scale1 * 0.293655420459897), 0, 0, 0);
            polyextra1.AddVertexAt(3, new Point2d(scale1 * 1.4926019155954, scale1 * 0.293360796259897), 0, 0, 0);
            polyextra1.AddVertexAt(4, new Point2d(scale1 * 1.48850991357873, scale1 * 0.294833917559897), 0, 0, 0);
            polyextra1.AddVertexAt(5, new Point2d(scale1 * 1.4852363119454, scale1 * 0.298958657259897), 0, 0, 0);
            polyextra1.AddVertexAt(6, new Point2d(scale1 * 1.48615669023306, scale1 * 0.3), 0, 0, 0);
            polyextra1.Closed = true;
            polyextra1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra1.Layer = "0";
            polyextra1.Color = color_gm;
            polyextra1.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra1);

            Polyline polyextra2 = new Polyline();
            polyextra2.AddVertexAt(0, new Point2d(scale1 * 1.48073510971774, scale1 * 0.248283283431083), 0, 0, 0);
            polyextra2.AddVertexAt(1, new Point2d(scale1 * 1.47255110566271, scale1 * 0.250345653295517), 0, 0, 0);
            polyextra2.AddVertexAt(2, new Point2d(scale1 * 1.46600390241792, scale1 * 0.25564889010042), 0, 0, 0);
            polyextra2.AddVertexAt(3, new Point2d(scale1 * 1.44758989328829, scale1 * 0.264192993827164), 0, 0, 0);
            polyextra2.AddVertexAt(4, new Point2d(scale1 * 1.43776908841915, scale1 * 0.272442473303527), 0, 0, 0);
            polyextra2.AddVertexAt(5, new Point2d(scale1 * 1.44022428963644, scale1 * 0.276272588763386), 0, 0, 0);
            polyextra2.AddVertexAt(6, new Point2d(scale1 * 1.45495549694014, scale1 * 0.283932819701731), 0, 0, 0);
            polyextra2.AddVertexAt(7, new Point2d(scale1 * 1.46641310261951, scale1 * 0.285405941028148), 0, 0, 0);
            polyextra2.AddVertexAt(8, new Point2d(scale1 * 1.47377870627136, scale1 * 0.282754322625697), 0, 0, 0);
            polyextra2.AddVertexAt(9, new Point2d(scale1 * 1.48359951113661, scale1 * 0.280102704223245), 0, 0, 0);
            polyextra2.AddVertexAt(10, new Point2d(scale1 * 1.49137431499005, scale1 * 0.278629582896829), 0, 0, 0);
            polyextra2.AddVertexAt(11, new Point2d(scale1 * 1.49587551722303, scale1 * 0.275683340225369), 0, 0, 0);
            polyextra2.AddVertexAt(12, new Point2d(scale1 * 1.5, scale1 * 0.274603475204557), 0, 0, 0);
            polyextra2.AddVertexAt(13, new Point2d(scale1 * 1.5, scale1 * 0.250541645527929), 0, 0, 0);
            polyextra2.AddVertexAt(14, new Point2d(scale1 * 1.49587551722303, scale1 * 0.249461780507118), 0, 0, 0);

            polyextra2.Closed = true;
            polyextra2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra2.Layer = "0";
            polyextra2.Color = color_gm;
            polyextra2.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra2);

            Polyline polyextra3 = new Polyline();
            polyextra3.AddVertexAt(0, new Point2d(scale1 * 1.48073510971774, scale1 * 0.248283283431083), 0, 0, 0);
            polyextra3.AddVertexAt(1, new Point2d(scale1 * 1.47255110566271, scale1 * 0.250345653295517), 0, 0, 0);
            polyextra3.AddVertexAt(2, new Point2d(scale1 * 1.46600390241792, scale1 * 0.25564889010042), 0, 0, 0);
            polyextra3.AddVertexAt(3, new Point2d(scale1 * 1.44758989328829, scale1 * 0.264192993827164), 0, 0, 0);
            polyextra3.AddVertexAt(4, new Point2d(scale1 * 1.43776908841915, scale1 * 0.272442473303527), 0, 0, 0);
            polyextra3.AddVertexAt(5, new Point2d(scale1 * 1.44022428963644, scale1 * 0.276272588763386), 0, 0, 0);
            polyextra3.AddVertexAt(6, new Point2d(scale1 * 1.45495549694014, scale1 * 0.283932819701731), 0, 0, 0);
            polyextra3.AddVertexAt(7, new Point2d(scale1 * 1.46641310261951, scale1 * 0.285405941028148), 0, 0, 0);
            polyextra3.AddVertexAt(8, new Point2d(scale1 * 1.47377870627136, scale1 * 0.282754322625697), 0, 0, 0);
            polyextra3.AddVertexAt(9, new Point2d(scale1 * 1.48359951113661, scale1 * 0.280102704223245), 0, 0, 0);
            polyextra3.AddVertexAt(10, new Point2d(scale1 * 1.49137431499005, scale1 * 0.278629582896829), 0, 0, 0);
            polyextra3.AddVertexAt(11, new Point2d(scale1 * 1.49587551722303, scale1 * 0.275683340225369), 0, 0, 0);
            polyextra3.AddVertexAt(12, new Point2d(scale1 * 1.5, scale1 * 0.274603475204557), 0, 0, 0);
            polyextra3.AddVertexAt(13, new Point2d(scale1 * 1.5, scale1 * 0.250541645527929), 0, 0, 0);
            polyextra3.AddVertexAt(14, new Point2d(scale1 * 1.49587551722303, scale1 * 0.249461780507118), 0, 0, 0);

            polyextra3.Closed = true;
            polyextra3.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra3.Layer = "0";
            polyextra3.Color = color_gm;
            polyextra3.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra3);

            Polyline polyextra4 = new Polyline();
            polyextra4.AddVertexAt(0, new Point2d(scale1 * 1.45679485542801, scale1 * 0.239821948474975), 0, 0, 0);
            polyextra4.AddVertexAt(1, new Point2d(scale1 * 1.46456965929468, scale1 * 0.238054202874975), 0, 0, 0);
            polyextra4.AddVertexAt(2, new Point2d(scale1 * 1.46661566031134, scale1 * 0.235107960174975), 0, 0, 0);
            polyextra4.AddVertexAt(3, new Point2d(scale1 * 1.46456965929468, scale1 * 0.231277844674975), 0, 0, 0);
            polyextra4.AddVertexAt(4, new Point2d(scale1 * 1.46129605766134, scale1 * 0.230983220474975), 0, 0, 0);
            polyextra4.AddVertexAt(5, new Point2d(scale1 * 1.45720405564468, scale1 * 0.232456341774975), 0, 0, 0);
            polyextra4.AddVertexAt(6, new Point2d(scale1 * 1.45393045401134, scale1 * 0.236581081474975), 0, 0, 0);
            polyextra4.Closed = true;
            polyextra4.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra4.Layer = "0";
            polyextra4.Color = color_gm;
            polyextra4.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra4);

            Polyline polyextra5 = new Polyline();
            polyextra5.AddVertexAt(0, new Point2d(scale1 * 1.47807326597801, scale1 * 0.229215474834975), 0, 0, 0);
            polyextra5.AddVertexAt(1, new Point2d(scale1 * 1.47766406577801, scale1 * 0.225090735104975), 0, 0, 0);
            polyextra5.AddVertexAt(2, new Point2d(scale1 * 1.48216526801134, scale1 * 0.221849868164975), 0, 0, 0);
            polyextra5.AddVertexAt(3, new Point2d(scale1 * 1.48789407084468, scale1 * 0.222733740974975), 0, 0, 0);
            polyextra5.AddVertexAt(4, new Point2d(scale1 * 1.48912167146134, scale1 * 0.227153104974975), 0, 0, 0);
            polyextra5.AddVertexAt(5, new Point2d(scale1 * 1.48830327106134, scale1 * 0.230688596174975), 0, 0, 0);
            polyextra5.AddVertexAt(6, new Point2d(scale1 * 1.48134686761134, scale1 * 0.232456341774975), 0, 0, 0);
            polyextra5.Closed = true;
            polyextra5.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra5.Layer = "0";
            polyextra5.Color = color_gm;
            polyextra5.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra5);

          Polyline  polyextra6 = new Polyline();
            polyextra6.AddVertexAt(0, new Point2d(scale1 * 1.5, scale1 * 0.174114035195047), 0, 0, 0);
            polyextra6.AddVertexAt(1, new Point2d(scale1 * 1.49539016077217, scale1 * 0.173755215127021), 0, 0, 0);
            polyextra6.AddVertexAt(2, new Point2d(scale1 * 1.48720615671715, scale1 * 0.175817584991455), 0, 0, 0);
            polyextra6.AddVertexAt(3, new Point2d(scale1 * 1.48065895347235, scale1 * 0.181120821796358), 0, 0, 0);
            polyextra6.AddVertexAt(4, new Point2d(scale1 * 1.46224494434272, scale1 * 0.189664925523102), 0, 0, 0);
            polyextra6.AddVertexAt(5, new Point2d(scale1 * 1.45242413947359, scale1 * 0.197914404999465), 0, 0, 0);
            polyextra6.AddVertexAt(6, new Point2d(scale1 * 1.45487934069087, scale1 * 0.201744520459324), 0, 0, 0);
            polyextra6.AddVertexAt(7, new Point2d(scale1 * 1.46961054799457, scale1 * 0.20940475139767), 0, 0, 0);
            polyextra6.AddVertexAt(8, new Point2d(scale1 * 1.48106815367394, scale1 * 0.210877872724086), 0, 0, 0);
            polyextra6.AddVertexAt(9, new Point2d(scale1 * 1.48843375732579, scale1 * 0.208226254321635), 0, 0, 0);
            polyextra6.AddVertexAt(10, new Point2d(scale1 * 1.49825456219104, scale1 * 0.205574635919183), 0, 0, 0);
            polyextra6.AddVertexAt(11, new Point2d(scale1 * 1.5, scale1 * 0.205243921253093), 0, 0, 0);
            polyextra6.Closed = true;
            polyextra6.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra6.Layer = "0";
            polyextra6.Color = color_gm;
            polyextra6.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra6);

            Polyline polyextra7 = new Polyline();
            polyextra7.AddVertexAt(0, new Point2d(scale1 * 1.43932973298399, scale1 * 0.131034696474671), 0, 0, 0);
            polyextra7.AddVertexAt(1, new Point2d(scale1 * 1.43605613135966, scale1 * 0.138105678874999), 0, 0, 0);
            polyextra7.AddVertexAt(2, new Point2d(scale1 * 1.4372837319683, scale1 * 0.146944406870753), 0, 0, 0);
            polyextra7.AddVertexAt(3, new Point2d(scale1 * 1.44874133764766, scale1 * 0.150774522330612), 0, 0, 0);
            polyextra7.AddVertexAt(4, new Point2d(scale1 * 1.4712473488048, scale1 * 0.149301401004195), 0, 0, 0);
            polyextra7.AddVertexAt(5, new Point2d(scale1 * 1.48270495448417, scale1 * 0.146944406870753), 0, 0, 0);
            polyextra7.AddVertexAt(6, new Point2d(scale1 * 1.48475095549986, scale1 * 0.143114291410893), 0, 0, 0);
            polyextra7.AddVertexAt(7, new Point2d(scale1 * 1.48516015570145, scale1 * 0.139284175932407), 0, 0, 0);
            polyextra7.AddVertexAt(8, new Point2d(scale1 * 1.48556935590304, scale1 * 0.134570187665522), 0, 0, 0);
            polyextra7.AddVertexAt(9, new Point2d(scale1 * 1.48884295752738, scale1 * 0.130445447936654), 0, 0, 0);
            polyextra7.AddVertexAt(10, new Point2d(scale1 * 1.48638775631009, scale1 * 0.123374465536326), 0, 0, 0);
            polyextra7.AddVertexAt(11, new Point2d(scale1 * 1.47738535184801, scale1 * 0.119544350076467), 0, 0, 0);
            polyextra7.AddVertexAt(12, new Point2d(scale1 * 1.45978974312544, scale1 * 0.11895510153845), 0, 0, 0);
            polyextra7.AddVertexAt(13, new Point2d(scale1 * 1.45119653886495, scale1 * 0.120722847133875), 0, 0, 0);
            polyextra7.AddVertexAt(14, new Point2d(scale1 * 1.44505853582174, scale1 * 0.124258338343352), 0, 0, 0);
            polyextra7.Closed = true;
            polyextra7.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra7.Layer = "0";
            polyextra7.Color = color_gm;
            polyextra7.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra7);

            Polyline polyextra8 = new Polyline();
            polyextra8.AddVertexAt(0, new Point2d(scale1 * 1.5, scale1 * 0.0641140351950464), 0, 0, 0);
            polyextra8.AddVertexAt(1, new Point2d(scale1 * 1.49539016077217, scale1 * 0.0637552151270211), 0, 0, 0);
            polyextra8.AddVertexAt(2, new Point2d(scale1 * 1.48720615671715, scale1 * 0.0658175849914548), 0, 0, 0);
            polyextra8.AddVertexAt(3, new Point2d(scale1 * 1.48065895347235, scale1 * 0.0711208217963577), 0, 0, 0);
            polyextra8.AddVertexAt(4, new Point2d(scale1 * 1.46224494434272, scale1 * 0.0796649255231023), 0, 0, 0);
            polyextra8.AddVertexAt(5, new Point2d(scale1 * 1.45242413947359, scale1 * 0.0879144049994647), 0, 0, 0);
            polyextra8.AddVertexAt(6, new Point2d(scale1 * 1.45487934069087, scale1 * 0.0917445204593244), 0, 0, 0);
            polyextra8.AddVertexAt(7, new Point2d(scale1 * 1.46961054799457, scale1 * 0.0994047513976692), 0, 0, 0);
            polyextra8.AddVertexAt(8, new Point2d(scale1 * 1.48106815367394, scale1 * 0.100877872724086), 0, 0, 0);
            polyextra8.AddVertexAt(9, new Point2d(scale1 * 1.48843375732579, scale1 * 0.0982262543216348), 0, 0, 0);
            polyextra8.AddVertexAt(10, new Point2d(scale1 * 1.49825456219104, scale1 * 0.0955746359191836), 0, 0, 0);
            polyextra8.AddVertexAt(11, new Point2d(scale1 * 1.5, scale1 * 0.0952439212530933), 0, 0, 0);
            polyextra8.Closed = true;
            polyextra8.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra8.Layer = "0";
            polyextra8.Color = color_gm;
            polyextra8.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra8);

            Polyline polyextra9 = new Polyline();
            polyextra9.AddVertexAt(0, new Point2d(scale1 * 1.43932973298399, scale1 * 0.0210346964746715), 0, 0, 0);
            polyextra9.AddVertexAt(1, new Point2d(scale1 * 1.43605613135966, scale1 * 0.0281056788749994), 0, 0, 0);
            polyextra9.AddVertexAt(2, new Point2d(scale1 * 1.4372837319683, scale1 * 0.0369444068707527), 0, 0, 0);
            polyextra9.AddVertexAt(3, new Point2d(scale1 * 1.44874133764766, scale1 * 0.040774522330612), 0, 0, 0);
            polyextra9.AddVertexAt(4, new Point2d(scale1 * 1.4712473488048, scale1 * 0.0393014010041952), 0, 0, 0);
            polyextra9.AddVertexAt(5, new Point2d(scale1 * 1.48270495448417, scale1 * 0.0369444068707527), 0, 0, 0);
            polyextra9.AddVertexAt(6, new Point2d(scale1 * 1.48475095549986, scale1 * 0.0331142914108931), 0, 0, 0);
            polyextra9.AddVertexAt(7, new Point2d(scale1 * 1.48516015570145, scale1 * 0.0292841759324074), 0, 0, 0);
            polyextra9.AddVertexAt(8, new Point2d(scale1 * 1.48556935590304, scale1 * 0.024570187665522), 0, 0, 0);
            polyextra9.AddVertexAt(9, new Point2d(scale1 * 1.48884295752738, scale1 * 0.020445447936654), 0, 0, 0);
            polyextra9.AddVertexAt(10, new Point2d(scale1 * 1.48638775631009, scale1 * 0.0133744655363262), 0, 0, 0);
            polyextra9.AddVertexAt(11, new Point2d(scale1 * 1.47738535184801, scale1 * 0.00954435007646692), 0, 0, 0);
            polyextra9.AddVertexAt(12, new Point2d(scale1 * 1.45978974312544, scale1 * 0.00895510153844947), 0, 0, 0);
            polyextra9.AddVertexAt(13, new Point2d(scale1 * 1.45119653886495, scale1 * 0.0107228471338749), 0, 0, 0);
            polyextra9.AddVertexAt(14, new Point2d(scale1 * 1.44505853582174, scale1 * 0.0142583383433519), 0, 0, 0);
            polyextra9.Closed = true;
            polyextra9.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra9.Layer = "0";
            polyextra9.Color = color_gm;
            polyextra9.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra9);

            Polyline polyextra10 = new Polyline();
            polyextra10.AddVertexAt(0, new Point2d(scale1 * 1.33563698629557, scale1 * 0.283954357764062), 0, 0, 0);
            polyextra10.AddVertexAt(1, new Point2d(scale1 * 1.33522778609557, scale1 * 0.279829618034062), 0, 0, 0);
            polyextra10.AddVertexAt(2, new Point2d(scale1 * 1.3397289883289, scale1 * 0.276588751094062), 0, 0, 0);
            polyextra10.AddVertexAt(3, new Point2d(scale1 * 1.34545779116223, scale1 * 0.277472623904062), 0, 0, 0);
            polyextra10.AddVertexAt(4, new Point2d(scale1 * 1.3466853917789, scale1 * 0.281891987904062), 0, 0, 0);
            polyextra10.AddVertexAt(5, new Point2d(scale1 * 1.3458669913789, scale1 * 0.285427479104062), 0, 0, 0);
            polyextra10.AddVertexAt(6, new Point2d(scale1 * 1.3389105879289, scale1 * 0.287195224704062), 0, 0, 0);
            polyextra10.Closed = true;
            polyextra10.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra10.Layer = "0";
            polyextra10.Color = color_gm;
            polyextra10.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra10);

            Polyline polyextra11 = new Polyline();
            polyextra11.AddVertexAt(0, new Point2d(scale1 * 1.40857292051223, scale1 * 0.290757499904062), 0, 0, 0);
            polyextra11.AddVertexAt(1, new Point2d(scale1 * 1.4110281217289, scale1 * 0.286927384504062), 0, 0, 0);
            polyextra11.AddVertexAt(2, new Point2d(scale1 * 1.41798452511223, scale1 * 0.285454263104062), 0, 0, 0);
            polyextra11.AddVertexAt(3, new Point2d(scale1 * 1.4233041277789, scale1 * 0.288105881504062), 0, 0, 0);
            polyextra11.AddVertexAt(4, new Point2d(scale1 * 1.4233041277789, scale1 * 0.293703742604062), 0, 0, 0);
            polyextra11.AddVertexAt(5, new Point2d(scale1 * 1.41266492261223, scale1 * 0.295176863904062), 0, 0, 0);
            polyextra11.Closed = true;
            polyextra11.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra11.Layer = "0";
            polyextra11.Color = color_gm;
            polyextra11.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra11);

            Polyline polyextra12 = new Polyline();
            polyextra12.AddVertexAt(0, new Point2d(scale1 * 1.37047649856761, scale1 * 0.3), 0, 0, 0);
            polyextra12.AddVertexAt(1, new Point2d(scale1 * 1.36823004232952, scale1 * 0.298498082924634), 0, 0, 0);
            polyextra12.AddVertexAt(2, new Point2d(scale1 * 1.3645472405036, scale1 * 0.287596985064447), 0, 0, 0);
            polyextra12.AddVertexAt(3, new Point2d(scale1 * 1.3645472405036, scale1 * 0.282293748259544), 0, 0, 0);
            polyextra12.AddVertexAt(4, new Point2d(scale1 * 1.36331963989496, scale1 * 0.27080340186134), 0, 0, 0);
            polyextra12.AddVertexAt(5, new Point2d(scale1 * 1.36945764293817, scale1 * 0.265794789344072), 0, 0, 0);
            polyextra12.AddVertexAt(6, new Point2d(scale1 * 1.37805084719866, scale1 * 0.264027043730021), 0, 0, 0);
            polyextra12.AddVertexAt(7, new Point2d(scale1 * 1.39810165713852, scale1 * 0.264910916537046), 0, 0, 0);
            polyextra12.AddVertexAt(8, new Point2d(scale1 * 1.40342125977467, scale1 * 0.268151783477515), 0, 0, 0);
            polyextra12.AddVertexAt(9, new Point2d(scale1 * 1.4050580605849, scale1 * 0.276106638666243), 0, 0, 0);
            polyextra12.AddVertexAt(10, new Point2d(scale1 * 1.39973845794875, scale1 * 0.280526002664119), 0, 0, 0);
            polyextra12.AddVertexAt(11, new Point2d(scale1 * 1.39482805551418, scale1 * 0.285534615200013), 0, 0, 0);
            polyextra12.AddVertexAt(12, new Point2d(scale1 * 1.39073605348667, scale1 * 0.29290022186935), 0, 0, 0);
            polyextra12.AddVertexAt(13, new Point2d(scale1 * 1.38746245186233, scale1 * 0.298203458655626), 0, 0, 0);
            polyextra12.AddVertexAt(14, new Point2d(scale1 * 1.38104623539793, scale1 * 0.3), 0, 0, 0);
            polyextra12.Closed = true;
            polyextra12.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra12.Layer = "0";
            polyextra12.Color = color_gm;
            polyextra12.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra12);

            Polyline polyextra13 = new Polyline();
            polyextra13.AddVertexAt(0, new Point2d(scale1 * 1.40023698231904, scale1 * 0.244555058125407), 0, 0, 0);
            polyextra13.AddVertexAt(1, new Point2d(scale1 * 1.40064618252063, scale1 * 0.236894827187061), 0, 0, 0);
            polyextra13.AddVertexAt(2, new Point2d(scale1 * 1.40228298333474, scale1 * 0.231002341844142), 0, 0, 0);
            polyextra13.AddVertexAt(3, new Point2d(scale1 * 1.40964858698659, scale1 * 0.229234596248716), 0, 0, 0);
            polyextra13.AddVertexAt(4, new Point2d(scale1 * 1.41701419063456, scale1 * 0.230413093324751), 0, 0, 0);
            polyextra13.AddVertexAt(5, new Point2d(scale1 * 1.42315219367777, scale1 * 0.234832457322627), 0, 0, 0);
            polyextra13.AddVertexAt(6, new Point2d(scale1 * 1.41701419063456, scale1 * 0.245144306644797), 0, 0, 0);
            polyextra13.AddVertexAt(7, new Point2d(scale1 * 1.40304423080757, scale1 * 0.251867096740753), 0, 0, 0);
            polyextra13.Closed = true;
            polyextra13.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra13.Layer = "0";
            polyextra13.Color = color_gm;
            polyextra13.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra13);

            Polyline polyextra14 = new Polyline();
            polyextra14.AddVertexAt(0, new Point2d(scale1 * 1.34729286879146, scale1 * 0.24195700775832), 0, 0, 0);
            polyextra14.AddVertexAt(1, new Point2d(scale1 * 1.33800097512081, scale1 * 0.23713588340208), 0, 0, 0);
            polyextra14.AddVertexAt(2, new Point2d(scale1 * 1.33614259638746, scale1 * 0.230707717593759), 0, 0, 0);
            polyextra14.AddVertexAt(3, new Point2d(scale1 * 1.34543449005811, scale1 * 0.228832835890353), 0, 0, 0);
            polyextra14.AddVertexAt(4, new Point2d(scale1 * 1.35695643820645, scale1 * 0.230975557826459), 0, 0, 0);
            polyextra14.AddVertexAt(5, new Point2d(scale1 * 1.36364660165117, scale1 * 0.236600202918053), 0, 0, 0);
            polyextra14.AddVertexAt(6, new Point2d(scale1 * 1.35807146544724, scale1 * 0.241689167525619), 0, 0, 0);
            polyextra14.Closed = true;
            polyextra14.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra14.Layer = "0";
            polyextra14.Color = color_gm;
            polyextra14.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra14);

            Polyline polyextra15 = new Polyline();
            polyextra15.AddVertexAt(0, new Point2d(scale1 * 1.30699297210124, scale1 * 0.240644590575248), 0, 0, 0);
            polyextra15.AddVertexAt(1, new Point2d(scale1 * 1.29880896804621, scale1 * 0.242706960439682), 0, 0, 0);
            polyextra15.AddVertexAt(2, new Point2d(scale1 * 1.29226176480142, scale1 * 0.248010197244585), 0, 0, 0);
            polyextra15.AddVertexAt(3, new Point2d(scale1 * 1.27384775567179, scale1 * 0.256554300971329), 0, 0, 0);
            polyextra15.AddVertexAt(4, new Point2d(scale1 * 1.26402695080265, scale1 * 0.264803780447692), 0, 0, 0);
            polyextra15.AddVertexAt(5, new Point2d(scale1 * 1.26648215201994, scale1 * 0.268633895907551), 0, 0, 0);
            polyextra15.AddVertexAt(6, new Point2d(scale1 * 1.28121335932364, scale1 * 0.276294126845896), 0, 0, 0);
            polyextra15.AddVertexAt(7, new Point2d(scale1 * 1.292670965003, scale1 * 0.277767248172313), 0, 0, 0);
            polyextra15.AddVertexAt(8, new Point2d(scale1 * 1.30003656865486, scale1 * 0.275115629769862), 0, 0, 0);
            polyextra15.AddVertexAt(9, new Point2d(scale1 * 1.30985737352011, scale1 * 0.27246401136741), 0, 0, 0);
            polyextra15.AddVertexAt(10, new Point2d(scale1 * 1.31763217737355, scale1 * 0.270990890040994), 0, 0, 0);
            polyextra15.AddVertexAt(11, new Point2d(scale1 * 1.32213337960653, scale1 * 0.268044647369534), 0, 0, 0);
            polyextra15.AddVertexAt(12, new Point2d(scale1 * 1.33563698630159, scale1 * 0.264509156178683), 0, 0, 0);
            polyextra15.AddVertexAt(13, new Point2d(scale1 * 1.34136578913933, scale1 * 0.254786555375904), 0, 0, 0);
            polyextra15.AddVertexAt(14, new Point2d(scale1 * 1.33563698630159, scale1 * 0.245358578842133), 0, 0, 0);
            polyextra15.AddVertexAt(15, new Point2d(scale1 * 1.32213337960653, scale1 * 0.241823087651283), 0, 0, 0);
            polyextra15.Closed = true;
            polyextra15.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra15.Layer = "0";
            polyextra15.Color = color_gm;
            polyextra15.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra15);

            Polyline polyextra16 = new Polyline();
            polyextra16.AddVertexAt(0, new Point2d(scale1 * 1.31435857574557, scale1 * 0.294560831404062), 0, 0, 0);
            polyextra16.AddVertexAt(1, new Point2d(scale1 * 1.32213337961223, scale1 * 0.292793085804062), 0, 0, 0);
            polyextra16.AddVertexAt(2, new Point2d(scale1 * 1.3241793806289, scale1 * 0.289846843104062), 0, 0, 0);
            polyextra16.AddVertexAt(3, new Point2d(scale1 * 1.32213337961223, scale1 * 0.286016727604062), 0, 0, 0);
            polyextra16.AddVertexAt(4, new Point2d(scale1 * 1.3188597779789, scale1 * 0.285722103404062), 0, 0, 0);
            polyextra16.AddVertexAt(5, new Point2d(scale1 * 1.31476777596223, scale1 * 0.287195224704062), 0, 0, 0);
            polyextra16.AddVertexAt(6, new Point2d(scale1 * 1.3114941743289, scale1 * 0.291319964404062), 0, 0, 0);
            polyextra16.Closed = true;
            polyextra16.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra16.Layer = "0";
            polyextra16.Color = color_gm;
            polyextra16.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra16);

            Polyline polyextra17 = new Polyline();
            polyextra17.AddVertexAt(0, new Point2d(scale1 * 1.13435857574557, scale1 * 0.294560831404062), 0, 0, 0);
            polyextra17.AddVertexAt(1, new Point2d(scale1 * 1.14213337961223, scale1 * 0.292793085804062), 0, 0, 0);
            polyextra17.AddVertexAt(2, new Point2d(scale1 * 1.1441793806289, scale1 * 0.289846843104062), 0, 0, 0);
            polyextra17.AddVertexAt(3, new Point2d(scale1 * 1.14213337961223, scale1 * 0.286016727604062), 0, 0, 0);
            polyextra17.AddVertexAt(4, new Point2d(scale1 * 1.1388597779789, scale1 * 0.285722103404062), 0, 0, 0);
            polyextra17.AddVertexAt(5, new Point2d(scale1 * 1.13476777596223, scale1 * 0.287195224704062), 0, 0, 0);
            polyextra17.AddVertexAt(6, new Point2d(scale1 * 1.1314941743289, scale1 * 0.291319964404062), 0, 0, 0);
            polyextra17.Closed = true;
            polyextra17.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra17.Layer = "0";
            polyextra17.Color = color_gm;
            polyextra17.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra17);

            Polyline polyextra18 = new Polyline();
            polyextra18.AddVertexAt(0, new Point2d(scale1 * 1.15563698629557, scale1 * 0.283954357764062), 0, 0, 0);
            polyextra18.AddVertexAt(1, new Point2d(scale1 * 1.15522778609557, scale1 * 0.279829618034062), 0, 0, 0);
            polyextra18.AddVertexAt(2, new Point2d(scale1 * 1.1597289883289, scale1 * 0.276588751094062), 0, 0, 0);
            polyextra18.AddVertexAt(3, new Point2d(scale1 * 1.16545779116223, scale1 * 0.277472623904062), 0, 0, 0);
            polyextra18.AddVertexAt(4, new Point2d(scale1 * 1.1666853917789, scale1 * 0.281891987904062), 0, 0, 0);
            polyextra18.AddVertexAt(5, new Point2d(scale1 * 1.1658669913789, scale1 * 0.285427479104062), 0, 0, 0);
            polyextra18.AddVertexAt(6, new Point2d(scale1 * 1.1589105879289, scale1 * 0.287195224704062), 0, 0, 0);
            polyextra18.Closed = true;
            polyextra18.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra18.Layer = "0";
            polyextra18.Color = color_gm;
            polyextra18.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra18);

            Polyline polyextra19 = new Polyline();
            polyextra19.AddVertexAt(0, new Point2d(scale1 * 1.22857292051223, scale1 * 0.290757499904062), 0, 0, 0);
            polyextra19.AddVertexAt(1, new Point2d(scale1 * 1.2310281217289, scale1 * 0.286927384504062), 0, 0, 0);
            polyextra19.AddVertexAt(2, new Point2d(scale1 * 1.23798452511223, scale1 * 0.285454263104062), 0, 0, 0);
            polyextra19.AddVertexAt(3, new Point2d(scale1 * 1.2433041277789, scale1 * 0.288105881504062), 0, 0, 0);
            polyextra19.AddVertexAt(4, new Point2d(scale1 * 1.2433041277789, scale1 * 0.293703742604062), 0, 0, 0);
            polyextra19.AddVertexAt(5, new Point2d(scale1 * 1.23266492261223, scale1 * 0.295176863904062), 0, 0, 0);
            polyextra19.Closed = true;
            polyextra19.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra19.Layer = "0";
            polyextra19.Color = color_gm;
            polyextra19.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra19);

            Polyline polyextra20 = new Polyline();
            polyextra20.AddVertexAt(0, new Point2d(scale1 * 1.19047649856761, scale1 * 0.3), 0, 0, 0);
            polyextra20.AddVertexAt(1, new Point2d(scale1 * 1.18823004232952, scale1 * 0.298498082924634), 0, 0, 0);
            polyextra20.AddVertexAt(2, new Point2d(scale1 * 1.1845472405036, scale1 * 0.287596985064447), 0, 0, 0);
            polyextra20.AddVertexAt(3, new Point2d(scale1 * 1.1845472405036, scale1 * 0.282293748259544), 0, 0, 0);
            polyextra20.AddVertexAt(4, new Point2d(scale1 * 1.18331963989496, scale1 * 0.27080340186134), 0, 0, 0);
            polyextra20.AddVertexAt(5, new Point2d(scale1 * 1.18945764293817, scale1 * 0.265794789344072), 0, 0, 0);
            polyextra20.AddVertexAt(6, new Point2d(scale1 * 1.19805084719866, scale1 * 0.264027043730021), 0, 0, 0);
            polyextra20.AddVertexAt(7, new Point2d(scale1 * 1.21810165713852, scale1 * 0.264910916537046), 0, 0, 0);
            polyextra20.AddVertexAt(8, new Point2d(scale1 * 1.22342125977467, scale1 * 0.268151783477515), 0, 0, 0);
            polyextra20.AddVertexAt(9, new Point2d(scale1 * 1.2250580605849, scale1 * 0.276106638666243), 0, 0, 0);
            polyextra20.AddVertexAt(10, new Point2d(scale1 * 1.21973845794875, scale1 * 0.280526002664119), 0, 0, 0);
            polyextra20.AddVertexAt(11, new Point2d(scale1 * 1.21482805551418, scale1 * 0.285534615200013), 0, 0, 0);
            polyextra20.AddVertexAt(12, new Point2d(scale1 * 1.21073605348667, scale1 * 0.29290022186935), 0, 0, 0);
            polyextra20.AddVertexAt(13, new Point2d(scale1 * 1.20746245186233, scale1 * 0.298203458655626), 0, 0, 0);
            polyextra20.AddVertexAt(14, new Point2d(scale1 * 1.20104623539793, scale1 * 0.3), 0, 0, 0);
            polyextra20.Closed = true;
            polyextra20.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra20.Layer = "0";
            polyextra20.Color = color_gm;
            polyextra20.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra20);

            Polyline polyextra21 = new Polyline();
            polyextra21.AddVertexAt(0, new Point2d(scale1 * 1.22023698231904, scale1 * 0.244555058125407), 0, 0, 0);
            polyextra21.AddVertexAt(1, new Point2d(scale1 * 1.22064618252063, scale1 * 0.236894827187061), 0, 0, 0);
            polyextra21.AddVertexAt(2, new Point2d(scale1 * 1.22228298333474, scale1 * 0.231002341844142), 0, 0, 0);
            polyextra21.AddVertexAt(3, new Point2d(scale1 * 1.22964858698659, scale1 * 0.229234596248716), 0, 0, 0);
            polyextra21.AddVertexAt(4, new Point2d(scale1 * 1.23701419063456, scale1 * 0.230413093324751), 0, 0, 0);
            polyextra21.AddVertexAt(5, new Point2d(scale1 * 1.24315219367777, scale1 * 0.234832457322627), 0, 0, 0);
            polyextra21.AddVertexAt(6, new Point2d(scale1 * 1.23701419063456, scale1 * 0.245144306644797), 0, 0, 0);
            polyextra21.AddVertexAt(7, new Point2d(scale1 * 1.22304423080757, scale1 * 0.251867096740753), 0, 0, 0);
            polyextra21.Closed = true;
            polyextra21.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra21.Layer = "0";
            polyextra21.Color = color_gm;
            polyextra21.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra21);

            Polyline polyextra22 = new Polyline();
            polyextra22.AddVertexAt(0, new Point2d(scale1 * 1.16729286879146, scale1 * 0.24195700775832), 0, 0, 0);
            polyextra22.AddVertexAt(1, new Point2d(scale1 * 1.15800097512081, scale1 * 0.23713588340208), 0, 0, 0);
            polyextra22.AddVertexAt(2, new Point2d(scale1 * 1.15614259638746, scale1 * 0.230707717593759), 0, 0, 0);
            polyextra22.AddVertexAt(3, new Point2d(scale1 * 1.16543449005811, scale1 * 0.228832835890353), 0, 0, 0);
            polyextra22.AddVertexAt(4, new Point2d(scale1 * 1.17695643820645, scale1 * 0.230975557826459), 0, 0, 0);
            polyextra22.AddVertexAt(5, new Point2d(scale1 * 1.18364660165117, scale1 * 0.236600202918053), 0, 0, 0);
            polyextra22.AddVertexAt(6, new Point2d(scale1 * 1.17807146544724, scale1 * 0.241689167525619), 0, 0, 0);
            polyextra22.Closed = true;
            polyextra22.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra22.Layer = "0";
            polyextra22.Color = color_gm;
            polyextra22.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra22);

            Polyline polyextra23 = new Polyline();
            polyextra23.AddVertexAt(0, new Point2d(scale1 * 1.12699297210124, scale1 * 0.240644590575248), 0, 0, 0);
            polyextra23.AddVertexAt(1, new Point2d(scale1 * 1.11880896804621, scale1 * 0.242706960439682), 0, 0, 0);
            polyextra23.AddVertexAt(2, new Point2d(scale1 * 1.11226176480142, scale1 * 0.248010197244585), 0, 0, 0);
            polyextra23.AddVertexAt(3, new Point2d(scale1 * 1.09384775567179, scale1 * 0.256554300971329), 0, 0, 0);
            polyextra23.AddVertexAt(4, new Point2d(scale1 * 1.08402695080265, scale1 * 0.264803780447692), 0, 0, 0);
            polyextra23.AddVertexAt(5, new Point2d(scale1 * 1.08648215201994, scale1 * 0.268633895907551), 0, 0, 0);
            polyextra23.AddVertexAt(6, new Point2d(scale1 * 1.10121335932364, scale1 * 0.276294126845896), 0, 0, 0);
            polyextra23.AddVertexAt(7, new Point2d(scale1 * 1.112670965003, scale1 * 0.277767248172313), 0, 0, 0);
            polyextra23.AddVertexAt(8, new Point2d(scale1 * 1.12003656865486, scale1 * 0.275115629769862), 0, 0, 0);
            polyextra23.AddVertexAt(9, new Point2d(scale1 * 1.12985737352011, scale1 * 0.27246401136741), 0, 0, 0);
            polyextra23.AddVertexAt(10, new Point2d(scale1 * 1.13763217737355, scale1 * 0.270990890040994), 0, 0, 0);
            polyextra23.AddVertexAt(11, new Point2d(scale1 * 1.14213337960653, scale1 * 0.268044647369534), 0, 0, 0);
            polyextra23.AddVertexAt(12, new Point2d(scale1 * 1.15563698630159, scale1 * 0.264509156178683), 0, 0, 0);
            polyextra23.AddVertexAt(13, new Point2d(scale1 * 1.16136578913933, scale1 * 0.254786555375904), 0, 0, 0);
            polyextra23.AddVertexAt(14, new Point2d(scale1 * 1.15563698630159, scale1 * 0.245358578842133), 0, 0, 0);
            polyextra23.AddVertexAt(15, new Point2d(scale1 * 1.14213337960653, scale1 * 0.241823087651283), 0, 0, 0);
            polyextra23.Closed = true;
            polyextra23.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra23.Layer = "0";
            polyextra23.Color = color_gm;
            polyextra23.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra23);

            Polyline polyextra24 = new Polyline();
            polyextra24.AddVertexAt(0, new Point2d(scale1 * 0.954358575745567, scale1 * 0.294560831404062), 0, 0, 0);
            polyextra24.AddVertexAt(1, new Point2d(scale1 * 0.962133379612234, scale1 * 0.292793085804062), 0, 0, 0);
            polyextra24.AddVertexAt(2, new Point2d(scale1 * 0.9641793806289, scale1 * 0.289846843104062), 0, 0, 0);
            polyextra24.AddVertexAt(3, new Point2d(scale1 * 0.962133379612234, scale1 * 0.286016727604062), 0, 0, 0);
            polyextra24.AddVertexAt(4, new Point2d(scale1 * 0.9588597779789, scale1 * 0.285722103404062), 0, 0, 0);
            polyextra24.AddVertexAt(5, new Point2d(scale1 * 0.954767775962234, scale1 * 0.287195224704062), 0, 0, 0);
            polyextra24.AddVertexAt(6, new Point2d(scale1 * 0.9514941743289, scale1 * 0.291319964404062), 0, 0, 0);
            polyextra24.Closed = true;
            polyextra24.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra24.Layer = "0";
            polyextra24.Color = color_gm;
            polyextra24.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra24);

            Polyline polyextra25 = new Polyline();
            polyextra25.AddVertexAt(0, new Point2d(scale1 * 0.975636986295567, scale1 * 0.283954357764062), 0, 0, 0);
            polyextra25.AddVertexAt(1, new Point2d(scale1 * 0.975227786095567, scale1 * 0.279829618034062), 0, 0, 0);
            polyextra25.AddVertexAt(2, new Point2d(scale1 * 0.9797289883289, scale1 * 0.276588751094062), 0, 0, 0);
            polyextra25.AddVertexAt(3, new Point2d(scale1 * 0.985457791162234, scale1 * 0.277472623904062), 0, 0, 0);
            polyextra25.AddVertexAt(4, new Point2d(scale1 * 0.9866853917789, scale1 * 0.281891987904062), 0, 0, 0);
            polyextra25.AddVertexAt(5, new Point2d(scale1 * 0.9858669913789, scale1 * 0.285427479104062), 0, 0, 0);
            polyextra25.AddVertexAt(6, new Point2d(scale1 * 0.9789105879289, scale1 * 0.287195224704062), 0, 0, 0);
            polyextra25.Closed = true;
            polyextra25.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra25.Layer = "0";
            polyextra25.Color = color_gm;
            polyextra25.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra25);

            Polyline polyextra26 = new Polyline();
            polyextra26.AddVertexAt(0, new Point2d(scale1 * 1.04857292051223, scale1 * 0.290757499904062), 0, 0, 0);
            polyextra26.AddVertexAt(1, new Point2d(scale1 * 1.0510281217289, scale1 * 0.286927384504062), 0, 0, 0);
            polyextra26.AddVertexAt(2, new Point2d(scale1 * 1.05798452511223, scale1 * 0.285454263104062), 0, 0, 0);
            polyextra26.AddVertexAt(3, new Point2d(scale1 * 1.0633041277789, scale1 * 0.288105881504062), 0, 0, 0);
            polyextra26.AddVertexAt(4, new Point2d(scale1 * 1.0633041277789, scale1 * 0.293703742604062), 0, 0, 0);
            polyextra26.AddVertexAt(5, new Point2d(scale1 * 1.05266492261223, scale1 * 0.295176863904062), 0, 0, 0);
            polyextra26.Closed = true;
            polyextra26.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra26.Layer = "0";
            polyextra26.Color = color_gm;
            polyextra26.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra26);

            Polyline polyextra27 = new Polyline();
            polyextra27.AddVertexAt(0, new Point2d(scale1 * 1.01047649856761, scale1 * 0.3), 0, 0, 0);
            polyextra27.AddVertexAt(1, new Point2d(scale1 * 1.00823004232952, scale1 * 0.298498082924634), 0, 0, 0);
            polyextra27.AddVertexAt(2, new Point2d(scale1 * 1.0045472405036, scale1 * 0.287596985064447), 0, 0, 0);
            polyextra27.AddVertexAt(3, new Point2d(scale1 * 1.0045472405036, scale1 * 0.282293748259544), 0, 0, 0);
            polyextra27.AddVertexAt(4, new Point2d(scale1 * 1.00331963989496, scale1 * 0.27080340186134), 0, 0, 0);
            polyextra27.AddVertexAt(5, new Point2d(scale1 * 1.00945764293817, scale1 * 0.265794789344072), 0, 0, 0);
            polyextra27.AddVertexAt(6, new Point2d(scale1 * 1.01805084719866, scale1 * 0.264027043730021), 0, 0, 0);
            polyextra27.AddVertexAt(7, new Point2d(scale1 * 1.03810165713852, scale1 * 0.264910916537046), 0, 0, 0);
            polyextra27.AddVertexAt(8, new Point2d(scale1 * 1.04342125977467, scale1 * 0.268151783477515), 0, 0, 0);
            polyextra27.AddVertexAt(9, new Point2d(scale1 * 1.0450580605849, scale1 * 0.276106638666243), 0, 0, 0);
            polyextra27.AddVertexAt(10, new Point2d(scale1 * 1.03973845794874, scale1 * 0.280526002664119), 0, 0, 0);
            polyextra27.AddVertexAt(11, new Point2d(scale1 * 1.03482805551418, scale1 * 0.285534615200013), 0, 0, 0);
            polyextra27.AddVertexAt(12, new Point2d(scale1 * 1.03073605348666, scale1 * 0.29290022186935), 0, 0, 0);
            polyextra27.AddVertexAt(13, new Point2d(scale1 * 1.02746245186233, scale1 * 0.298203458655626), 0, 0, 0);
            polyextra27.AddVertexAt(14, new Point2d(scale1 * 1.02104623539793, scale1 * 0.3), 0, 0, 0);
            polyextra27.Closed = true;
            polyextra27.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra27.Layer = "0";
            polyextra27.Color = color_gm;
            polyextra27.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra27);

            Polyline polyextra28 = new Polyline();
            polyextra28.AddVertexAt(0, new Point2d(scale1 * 1.04023698231904, scale1 * 0.244555058125407), 0, 0, 0);
            polyextra28.AddVertexAt(1, new Point2d(scale1 * 1.04064618252063, scale1 * 0.236894827187061), 0, 0, 0);
            polyextra28.AddVertexAt(2, new Point2d(scale1 * 1.04228298333473, scale1 * 0.231002341844142), 0, 0, 0);
            polyextra28.AddVertexAt(3, new Point2d(scale1 * 1.04964858698659, scale1 * 0.229234596248716), 0, 0, 0);
            polyextra28.AddVertexAt(4, new Point2d(scale1 * 1.05701419063456, scale1 * 0.230413093324751), 0, 0, 0);
            polyextra28.AddVertexAt(5, new Point2d(scale1 * 1.06315219367777, scale1 * 0.234832457322627), 0, 0, 0);
            polyextra28.AddVertexAt(6, new Point2d(scale1 * 1.05701419063456, scale1 * 0.245144306644797), 0, 0, 0);
            polyextra28.AddVertexAt(7, new Point2d(scale1 * 1.04304423080757, scale1 * 0.251867096740753), 0, 0, 0);
            polyextra28.Closed = true;
            polyextra28.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra28.Layer = "0";
            polyextra28.Color = color_gm;
            polyextra28.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra28);

            Polyline polyextra29 = new Polyline();
            polyextra29.AddVertexAt(0, new Point2d(scale1 * 0.987292868791458, scale1 * 0.24195700775832), 0, 0, 0);
            polyextra29.AddVertexAt(1, new Point2d(scale1 * 0.978000975120813, scale1 * 0.23713588340208), 0, 0, 0);
            polyextra29.AddVertexAt(2, new Point2d(scale1 * 0.976142596387459, scale1 * 0.230707717593759), 0, 0, 0);
            polyextra29.AddVertexAt(3, new Point2d(scale1 * 0.985434490058105, scale1 * 0.228832835890353), 0, 0, 0);
            polyextra29.AddVertexAt(4, new Point2d(scale1 * 0.996956438206447, scale1 * 0.230975557826459), 0, 0, 0);
            polyextra29.AddVertexAt(5, new Point2d(scale1 * 1.00364660165117, scale1 * 0.236600202918053), 0, 0, 0);
            polyextra29.AddVertexAt(6, new Point2d(scale1 * 0.998071465447235, scale1 * 0.241689167525619), 0, 0, 0);
            polyextra29.Closed = true;
            polyextra29.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra29.Layer = "0";
            polyextra29.Color = color_gm;
            polyextra29.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra29);

            Polyline polyextra30 = new Polyline();
            polyextra30.AddVertexAt(0, new Point2d(scale1 * 0.94699297210124, scale1 * 0.240644590575248), 0, 0, 0);
            polyextra30.AddVertexAt(1, new Point2d(scale1 * 0.938808968046214, scale1 * 0.242706960439682), 0, 0, 0);
            polyextra30.AddVertexAt(2, new Point2d(scale1 * 0.932261764801418, scale1 * 0.248010197244585), 0, 0, 0);
            polyextra30.AddVertexAt(3, new Point2d(scale1 * 0.913847755671789, scale1 * 0.256554300971329), 0, 0, 0);
            polyextra30.AddVertexAt(4, new Point2d(scale1 * 0.904026950802654, scale1 * 0.264803780447692), 0, 0, 0);
            polyextra30.AddVertexAt(5, new Point2d(scale1 * 0.906482152019938, scale1 * 0.268633895907551), 0, 0, 0);
            polyextra30.AddVertexAt(6, new Point2d(scale1 * 0.921213359323641, scale1 * 0.276294126845896), 0, 0, 0);
            polyextra30.AddVertexAt(7, new Point2d(scale1 * 0.932670965003005, scale1 * 0.277767248172313), 0, 0, 0);
            polyextra30.AddVertexAt(8, new Point2d(scale1 * 0.940036568654856, scale1 * 0.275115629769862), 0, 0, 0);
            polyextra30.AddVertexAt(9, new Point2d(scale1 * 0.949857373520111, scale1 * 0.27246401136741), 0, 0, 0);
            polyextra30.AddVertexAt(10, new Point2d(scale1 * 0.957632177373549, scale1 * 0.270990890040994), 0, 0, 0);
            polyextra30.AddVertexAt(11, new Point2d(scale1 * 0.96213337960653, scale1 * 0.268044647369534), 0, 0, 0);
            polyextra30.AddVertexAt(12, new Point2d(scale1 * 0.975636986301591, scale1 * 0.264509156178683), 0, 0, 0);
            polyextra30.AddVertexAt(13, new Point2d(scale1 * 0.981365789139333, scale1 * 0.254786555375904), 0, 0, 0);
            polyextra30.AddVertexAt(14, new Point2d(scale1 * 0.975636986301591, scale1 * 0.245358578842133), 0, 0, 0);
            polyextra30.AddVertexAt(15, new Point2d(scale1 * 0.96213337960653, scale1 * 0.241823087651283), 0, 0, 0);
            polyextra30.Closed = true;
            polyextra30.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra30.Layer = "0";
            polyextra30.Color = color_gm;
            polyextra30.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra30);

            Polyline polyextra31 = new Polyline();
            polyextra31.AddVertexAt(0, new Point2d(scale1 * 0.774358575745567, scale1 * 0.294560831404062), 0, 0, 0);
            polyextra31.AddVertexAt(1, new Point2d(scale1 * 0.782133379612234, scale1 * 0.292793085804062), 0, 0, 0);
            polyextra31.AddVertexAt(2, new Point2d(scale1 * 0.7841793806289, scale1 * 0.289846843104062), 0, 0, 0);
            polyextra31.AddVertexAt(3, new Point2d(scale1 * 0.782133379612234, scale1 * 0.286016727604062), 0, 0, 0);
            polyextra31.AddVertexAt(4, new Point2d(scale1 * 0.7788597779789, scale1 * 0.285722103404062), 0, 0, 0);
            polyextra31.AddVertexAt(5, new Point2d(scale1 * 0.774767775962234, scale1 * 0.287195224704062), 0, 0, 0);
            polyextra31.AddVertexAt(6, new Point2d(scale1 * 0.7714941743289, scale1 * 0.291319964404062), 0, 0, 0);
            polyextra31.Closed = true;
            polyextra31.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra31.Layer = "0";
            polyextra31.Color = color_gm;
            polyextra31.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra31);

            Polyline polyextra32 = new Polyline();
            polyextra32.AddVertexAt(0, new Point2d(scale1 * 0.795636986295567, scale1 * 0.283954357764062), 0, 0, 0);
            polyextra32.AddVertexAt(1, new Point2d(scale1 * 0.795227786095567, scale1 * 0.279829618034062), 0, 0, 0);
            polyextra32.AddVertexAt(2, new Point2d(scale1 * 0.7997289883289, scale1 * 0.276588751094062), 0, 0, 0);
            polyextra32.AddVertexAt(3, new Point2d(scale1 * 0.805457791162234, scale1 * 0.277472623904062), 0, 0, 0);
            polyextra32.AddVertexAt(4, new Point2d(scale1 * 0.8066853917789, scale1 * 0.281891987904062), 0, 0, 0);
            polyextra32.AddVertexAt(5, new Point2d(scale1 * 0.8058669913789, scale1 * 0.285427479104062), 0, 0, 0);
            polyextra32.AddVertexAt(6, new Point2d(scale1 * 0.7989105879289, scale1 * 0.287195224704062), 0, 0, 0);
            polyextra32.Closed = true;
            polyextra32.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra32.Layer = "0";
            polyextra32.Color = color_gm;
            polyextra32.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra32);

            Polyline polyextra33 = new Polyline();
            polyextra33.AddVertexAt(0, new Point2d(scale1 * 0.868572920512234, scale1 * 0.290757499904062), 0, 0, 0);
            polyextra33.AddVertexAt(1, new Point2d(scale1 * 0.8710281217289, scale1 * 0.286927384504062), 0, 0, 0);
            polyextra33.AddVertexAt(2, new Point2d(scale1 * 0.877984525112234, scale1 * 0.285454263104062), 0, 0, 0);
            polyextra33.AddVertexAt(3, new Point2d(scale1 * 0.8833041277789, scale1 * 0.288105881504062), 0, 0, 0);
            polyextra33.AddVertexAt(4, new Point2d(scale1 * 0.8833041277789, scale1 * 0.293703742604062), 0, 0, 0);
            polyextra33.AddVertexAt(5, new Point2d(scale1 * 0.872664922612234, scale1 * 0.295176863904062), 0, 0, 0);
            polyextra33.Closed = true;
            polyextra33.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra33.Layer = "0";
            polyextra33.Color = color_gm;
            polyextra33.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra33);

            Polyline polyextra34 = new Polyline();
            polyextra34.AddVertexAt(0, new Point2d(scale1 * 0.830476498567611, scale1 * 0.3), 0, 0, 0);
            polyextra34.AddVertexAt(1, new Point2d(scale1 * 0.828230042329524, scale1 * 0.298498082924634), 0, 0, 0);
            polyextra34.AddVertexAt(2, new Point2d(scale1 * 0.824547240503598, scale1 * 0.287596985064447), 0, 0, 0);
            polyextra34.AddVertexAt(3, new Point2d(scale1 * 0.824547240503598, scale1 * 0.282293748259544), 0, 0, 0);
            polyextra34.AddVertexAt(4, new Point2d(scale1 * 0.823319639894956, scale1 * 0.27080340186134), 0, 0, 0);
            polyextra34.AddVertexAt(5, new Point2d(scale1 * 0.829457642938166, scale1 * 0.265794789344072), 0, 0, 0);
            polyextra34.AddVertexAt(6, new Point2d(scale1 * 0.838050847198659, scale1 * 0.264027043730021), 0, 0, 0);
            polyextra34.AddVertexAt(7, new Point2d(scale1 * 0.858101657138516, scale1 * 0.264910916537046), 0, 0, 0);
            polyextra34.AddVertexAt(8, new Point2d(scale1 * 0.863421259774671, scale1 * 0.268151783477515), 0, 0, 0);
            polyextra34.AddVertexAt(9, new Point2d(scale1 * 0.8650580605849, scale1 * 0.276106638666243), 0, 0, 0);
            polyextra34.AddVertexAt(10, new Point2d(scale1 * 0.859738457948745, scale1 * 0.280526002664119), 0, 0, 0);
            polyextra34.AddVertexAt(11, new Point2d(scale1 * 0.854828055514178, scale1 * 0.285534615200013), 0, 0, 0);
            polyextra34.AddVertexAt(12, new Point2d(scale1 * 0.850736053486665, scale1 * 0.29290022186935), 0, 0, 0);
            polyextra34.AddVertexAt(13, new Point2d(scale1 * 0.847462451862327, scale1 * 0.298203458655626), 0, 0, 0);
            polyextra34.AddVertexAt(14, new Point2d(scale1 * 0.841046235397932, scale1 * 0.3), 0, 0, 0);
            polyextra34.Closed = true;
            polyextra34.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra34.Layer = "0";
            polyextra34.Color = color_gm;
            polyextra34.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra34);

            Polyline polyextra35 = new Polyline();
            polyextra35.AddVertexAt(0, new Point2d(scale1 * 0.860236982319038, scale1 * 0.244555058125407), 0, 0, 0);
            polyextra35.AddVertexAt(1, new Point2d(scale1 * 0.860646182520626, scale1 * 0.236894827187061), 0, 0, 0);
            polyextra35.AddVertexAt(2, new Point2d(scale1 * 0.862282983334735, scale1 * 0.231002341844142), 0, 0, 0);
            polyextra35.AddVertexAt(3, new Point2d(scale1 * 0.869648586986586, scale1 * 0.229234596248716), 0, 0, 0);
            polyextra35.AddVertexAt(4, new Point2d(scale1 * 0.877014190634557, scale1 * 0.230413093324751), 0, 0, 0);
            polyextra35.AddVertexAt(5, new Point2d(scale1 * 0.883152193677767, scale1 * 0.234832457322627), 0, 0, 0);
            polyextra35.AddVertexAt(6, new Point2d(scale1 * 0.877014190634557, scale1 * 0.245144306644797), 0, 0, 0);
            polyextra35.AddVertexAt(7, new Point2d(scale1 * 0.863044230807573, scale1 * 0.251867096740753), 0, 0, 0);
            polyextra35.Closed = true;
            polyextra35.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra35.Layer = "0";
            polyextra35.Color = color_gm;
            polyextra35.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra35);

            Polyline polyextra36 = new Polyline();
            polyextra36.AddVertexAt(0, new Point2d(scale1 * 0.807292868791458, scale1 * 0.24195700775832), 0, 0, 0);
            polyextra36.AddVertexAt(1, new Point2d(scale1 * 0.798000975120813, scale1 * 0.23713588340208), 0, 0, 0);
            polyextra36.AddVertexAt(2, new Point2d(scale1 * 0.79614259638746, scale1 * 0.230707717593759), 0, 0, 0);
            polyextra36.AddVertexAt(3, new Point2d(scale1 * 0.805434490058105, scale1 * 0.228832835890353), 0, 0, 0);
            polyextra36.AddVertexAt(4, new Point2d(scale1 * 0.816956438206447, scale1 * 0.230975557826459), 0, 0, 0);
            polyextra36.AddVertexAt(5, new Point2d(scale1 * 0.823646601651174, scale1 * 0.236600202918053), 0, 0, 0);
            polyextra36.AddVertexAt(6, new Point2d(scale1 * 0.818071465447235, scale1 * 0.241689167525619), 0, 0, 0);
            polyextra36.Closed = true;
            polyextra36.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra36.Layer = "0";
            polyextra36.Color = color_gm;
            polyextra36.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra36);

            Polyline polyextra37 = new Polyline();
            polyextra37.AddVertexAt(0, new Point2d(scale1 * 0.76699297210124, scale1 * 0.240644590575248), 0, 0, 0);
            polyextra37.AddVertexAt(1, new Point2d(scale1 * 0.758808968046214, scale1 * 0.242706960439682), 0, 0, 0);
            polyextra37.AddVertexAt(2, new Point2d(scale1 * 0.752261764801418, scale1 * 0.248010197244585), 0, 0, 0);
            polyextra37.AddVertexAt(3, new Point2d(scale1 * 0.733847755671789, scale1 * 0.256554300971329), 0, 0, 0);
            polyextra37.AddVertexAt(4, new Point2d(scale1 * 0.724026950802654, scale1 * 0.264803780447692), 0, 0, 0);
            polyextra37.AddVertexAt(5, new Point2d(scale1 * 0.726482152019938, scale1 * 0.268633895907551), 0, 0, 0);
            polyextra37.AddVertexAt(6, new Point2d(scale1 * 0.741213359323641, scale1 * 0.276294126845896), 0, 0, 0);
            polyextra37.AddVertexAt(7, new Point2d(scale1 * 0.752670965003005, scale1 * 0.277767248172313), 0, 0, 0);
            polyextra37.AddVertexAt(8, new Point2d(scale1 * 0.760036568654856, scale1 * 0.275115629769862), 0, 0, 0);
            polyextra37.AddVertexAt(9, new Point2d(scale1 * 0.769857373520111, scale1 * 0.27246401136741), 0, 0, 0);
            polyextra37.AddVertexAt(10, new Point2d(scale1 * 0.77763217737355, scale1 * 0.270990890040994), 0, 0, 0);
            polyextra37.AddVertexAt(11, new Point2d(scale1 * 0.78213337960653, scale1 * 0.268044647369534), 0, 0, 0);
            polyextra37.AddVertexAt(12, new Point2d(scale1 * 0.795636986301591, scale1 * 0.264509156178683), 0, 0, 0);
            polyextra37.AddVertexAt(13, new Point2d(scale1 * 0.801365789139333, scale1 * 0.254786555375904), 0, 0, 0);
            polyextra37.AddVertexAt(14, new Point2d(scale1 * 0.795636986301591, scale1 * 0.245358578842133), 0, 0, 0);
            polyextra37.AddVertexAt(15, new Point2d(scale1 * 0.78213337960653, scale1 * 0.241823087651283), 0, 0, 0);
            polyextra37.Closed = true;
            polyextra37.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra37.Layer = "0";
            polyextra37.Color = color_gm;
            polyextra37.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra37);


            Polyline polyextra38 = new Polyline();
            polyextra38.AddVertexAt(0, new Point2d(scale1 * 0.594358575745567, scale1 * 0.294560831404062), 0, 0, 0);
            polyextra38.AddVertexAt(1, new Point2d(scale1 * 0.602133379612234, scale1 * 0.292793085804062), 0, 0, 0);
            polyextra38.AddVertexAt(2, new Point2d(scale1 * 0.6041793806289, scale1 * 0.289846843104062), 0, 0, 0);
            polyextra38.AddVertexAt(3, new Point2d(scale1 * 0.602133379612234, scale1 * 0.286016727604062), 0, 0, 0);
            polyextra38.AddVertexAt(4, new Point2d(scale1 * 0.5988597779789, scale1 * 0.285722103404062), 0, 0, 0);
            polyextra38.AddVertexAt(5, new Point2d(scale1 * 0.594767775962234, scale1 * 0.287195224704062), 0, 0, 0);
            polyextra38.AddVertexAt(6, new Point2d(scale1 * 0.5914941743289, scale1 * 0.291319964404062), 0, 0, 0);
            polyextra38.Closed = true;
            polyextra38.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra38.Layer = "0";
            polyextra38.Color = color_gm;
            polyextra38.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra38);

            Polyline polyextra39 = new Polyline();
            polyextra39.AddVertexAt(0, new Point2d(scale1 * 0.615636986295567, scale1 * 0.283954357764062), 0, 0, 0);
            polyextra39.AddVertexAt(1, new Point2d(scale1 * 0.615227786095567, scale1 * 0.279829618034062), 0, 0, 0);
            polyextra39.AddVertexAt(2, new Point2d(scale1 * 0.6197289883289, scale1 * 0.276588751094062), 0, 0, 0);
            polyextra39.AddVertexAt(3, new Point2d(scale1 * 0.625457791162234, scale1 * 0.277472623904062), 0, 0, 0);
            polyextra39.AddVertexAt(4, new Point2d(scale1 * 0.6266853917789, scale1 * 0.281891987904062), 0, 0, 0);
            polyextra39.AddVertexAt(5, new Point2d(scale1 * 0.6258669913789, scale1 * 0.285427479104062), 0, 0, 0);
            polyextra39.AddVertexAt(6, new Point2d(scale1 * 0.6189105879289, scale1 * 0.287195224704062), 0, 0, 0);
            polyextra39.Closed = true;
            polyextra39.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra39.Layer = "0";
            polyextra39.Color = color_gm;
            polyextra39.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra39);

            Polyline polyextra40 = new Polyline();
            polyextra40.AddVertexAt(0, new Point2d(scale1 * 0.688572920512234, scale1 * 0.290757499904062), 0, 0, 0);
            polyextra40.AddVertexAt(1, new Point2d(scale1 * 0.6910281217289, scale1 * 0.286927384504062), 0, 0, 0);
            polyextra40.AddVertexAt(2, new Point2d(scale1 * 0.697984525112234, scale1 * 0.285454263104062), 0, 0, 0);
            polyextra40.AddVertexAt(3, new Point2d(scale1 * 0.7033041277789, scale1 * 0.288105881504062), 0, 0, 0);
            polyextra40.AddVertexAt(4, new Point2d(scale1 * 0.7033041277789, scale1 * 0.293703742604062), 0, 0, 0);
            polyextra40.AddVertexAt(5, new Point2d(scale1 * 0.692664922612234, scale1 * 0.295176863904062), 0, 0, 0);
            polyextra40.Closed = true;
            polyextra40.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra40.Layer = "0";
            polyextra40.Color = color_gm;
            polyextra40.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra40);

            Polyline polyextra41 = new Polyline();
            polyextra41.AddVertexAt(0, new Point2d(scale1 * 0.650476498567611, scale1 * 0.3), 0, 0, 0);
            polyextra41.AddVertexAt(1, new Point2d(scale1 * 0.648230042329524, scale1 * 0.298498082924634), 0, 0, 0);
            polyextra41.AddVertexAt(2, new Point2d(scale1 * 0.644547240503598, scale1 * 0.287596985064447), 0, 0, 0);
            polyextra41.AddVertexAt(3, new Point2d(scale1 * 0.644547240503598, scale1 * 0.282293748259544), 0, 0, 0);
            polyextra41.AddVertexAt(4, new Point2d(scale1 * 0.643319639894956, scale1 * 0.27080340186134), 0, 0, 0);
            polyextra41.AddVertexAt(5, new Point2d(scale1 * 0.649457642938166, scale1 * 0.265794789344072), 0, 0, 0);
            polyextra41.AddVertexAt(6, new Point2d(scale1 * 0.658050847198659, scale1 * 0.264027043730021), 0, 0, 0);
            polyextra41.AddVertexAt(7, new Point2d(scale1 * 0.678101657138516, scale1 * 0.264910916537046), 0, 0, 0);
            polyextra41.AddVertexAt(8, new Point2d(scale1 * 0.683421259774671, scale1 * 0.268151783477515), 0, 0, 0);
            polyextra41.AddVertexAt(9, new Point2d(scale1 * 0.6850580605849, scale1 * 0.276106638666243), 0, 0, 0);
            polyextra41.AddVertexAt(10, new Point2d(scale1 * 0.679738457948745, scale1 * 0.280526002664119), 0, 0, 0);
            polyextra41.AddVertexAt(11, new Point2d(scale1 * 0.674828055514178, scale1 * 0.285534615200013), 0, 0, 0);
            polyextra41.AddVertexAt(12, new Point2d(scale1 * 0.670736053486665, scale1 * 0.29290022186935), 0, 0, 0);
            polyextra41.AddVertexAt(13, new Point2d(scale1 * 0.667462451862327, scale1 * 0.298203458655626), 0, 0, 0);
            polyextra41.AddVertexAt(14, new Point2d(scale1 * 0.661046235397932, scale1 * 0.3), 0, 0, 0);
            polyextra41.Closed = true;
            polyextra41.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra41.Layer = "0";
            polyextra41.Color = color_gm;
            polyextra41.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra41);

            Polyline polyextra42 = new Polyline();
            polyextra42.AddVertexAt(0, new Point2d(scale1 * 0.680236982319038, scale1 * 0.244555058125407), 0, 0, 0);
            polyextra42.AddVertexAt(1, new Point2d(scale1 * 0.680646182520626, scale1 * 0.236894827187061), 0, 0, 0);
            polyextra42.AddVertexAt(2, new Point2d(scale1 * 0.682282983334735, scale1 * 0.231002341844142), 0, 0, 0);
            polyextra42.AddVertexAt(3, new Point2d(scale1 * 0.689648586986586, scale1 * 0.229234596248716), 0, 0, 0);
            polyextra42.AddVertexAt(4, new Point2d(scale1 * 0.697014190634557, scale1 * 0.230413093324751), 0, 0, 0);
            polyextra42.AddVertexAt(5, new Point2d(scale1 * 0.703152193677767, scale1 * 0.234832457322627), 0, 0, 0);
            polyextra42.AddVertexAt(6, new Point2d(scale1 * 0.697014190634557, scale1 * 0.245144306644797), 0, 0, 0);
            polyextra42.AddVertexAt(7, new Point2d(scale1 * 0.683044230807573, scale1 * 0.251867096740753), 0, 0, 0);
            polyextra42.Closed = true;
            polyextra42.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra42.Layer = "0";
            polyextra42.Color = color_gm;
            polyextra42.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra42);

            Polyline polyextra43 = new Polyline();
            polyextra43.AddVertexAt(0, new Point2d(scale1 * 0.627292868791458, scale1 * 0.24195700775832), 0, 0, 0);
            polyextra43.AddVertexAt(1, new Point2d(scale1 * 0.618000975120813, scale1 * 0.23713588340208), 0, 0, 0);
            polyextra43.AddVertexAt(2, new Point2d(scale1 * 0.61614259638746, scale1 * 0.230707717593759), 0, 0, 0);
            polyextra43.AddVertexAt(3, new Point2d(scale1 * 0.625434490058105, scale1 * 0.228832835890353), 0, 0, 0);
            polyextra43.AddVertexAt(4, new Point2d(scale1 * 0.636956438206447, scale1 * 0.230975557826459), 0, 0, 0);
            polyextra43.AddVertexAt(5, new Point2d(scale1 * 0.643646601651174, scale1 * 0.236600202918053), 0, 0, 0);
            polyextra43.AddVertexAt(6, new Point2d(scale1 * 0.638071465447235, scale1 * 0.241689167525619), 0, 0, 0);
            polyextra43.Closed = true;
            polyextra43.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra43.Layer = "0";
            polyextra43.Color = color_gm;
            polyextra43.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra43);

            Polyline polyextra44 = new Polyline();
            polyextra44.AddVertexAt(0, new Point2d(scale1 * 0.58699297210124, scale1 * 0.240644590575248), 0, 0, 0);
            polyextra44.AddVertexAt(1, new Point2d(scale1 * 0.578808968046215, scale1 * 0.242706960439682), 0, 0, 0);
            polyextra44.AddVertexAt(2, new Point2d(scale1 * 0.572261764801418, scale1 * 0.248010197244585), 0, 0, 0);
            polyextra44.AddVertexAt(3, new Point2d(scale1 * 0.553847755671789, scale1 * 0.256554300971329), 0, 0, 0);
            polyextra44.AddVertexAt(4, new Point2d(scale1 * 0.544026950802654, scale1 * 0.264803780447692), 0, 0, 0);
            polyextra44.AddVertexAt(5, new Point2d(scale1 * 0.546482152019938, scale1 * 0.268633895907551), 0, 0, 0);
            polyextra44.AddVertexAt(6, new Point2d(scale1 * 0.561213359323641, scale1 * 0.276294126845896), 0, 0, 0);
            polyextra44.AddVertexAt(7, new Point2d(scale1 * 0.572670965003005, scale1 * 0.277767248172313), 0, 0, 0);
            polyextra44.AddVertexAt(8, new Point2d(scale1 * 0.580036568654856, scale1 * 0.275115629769862), 0, 0, 0);
            polyextra44.AddVertexAt(9, new Point2d(scale1 * 0.589857373520111, scale1 * 0.27246401136741), 0, 0, 0);
            polyextra44.AddVertexAt(10, new Point2d(scale1 * 0.59763217737355, scale1 * 0.270990890040994), 0, 0, 0);
            polyextra44.AddVertexAt(11, new Point2d(scale1 * 0.60213337960653, scale1 * 0.268044647369534), 0, 0, 0);
            polyextra44.AddVertexAt(12, new Point2d(scale1 * 0.615636986301591, scale1 * 0.264509156178683), 0, 0, 0);
            polyextra44.AddVertexAt(13, new Point2d(scale1 * 0.621365789139333, scale1 * 0.254786555375904), 0, 0, 0);
            polyextra44.AddVertexAt(14, new Point2d(scale1 * 0.615636986301591, scale1 * 0.245358578842133), 0, 0, 0);
            polyextra44.AddVertexAt(15, new Point2d(scale1 * 0.60213337960653, scale1 * 0.241823087651283), 0, 0, 0);
            polyextra44.Closed = true;
            polyextra44.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra44.Layer = "0";
            polyextra44.Color = color_gm;
            polyextra44.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra44);

            Polyline polyextra45 = new Polyline();
            polyextra45.AddVertexAt(0, new Point2d(scale1 * 0.414358575745567, scale1 * 0.294560831404062), 0, 0, 0);
            polyextra45.AddVertexAt(1, new Point2d(scale1 * 0.422133379612234, scale1 * 0.292793085804062), 0, 0, 0);
            polyextra45.AddVertexAt(2, new Point2d(scale1 * 0.4241793806289, scale1 * 0.289846843104062), 0, 0, 0);
            polyextra45.AddVertexAt(3, new Point2d(scale1 * 0.422133379612234, scale1 * 0.286016727604062), 0, 0, 0);
            polyextra45.AddVertexAt(4, new Point2d(scale1 * 0.4188597779789, scale1 * 0.285722103404062), 0, 0, 0);
            polyextra45.AddVertexAt(5, new Point2d(scale1 * 0.414767775962234, scale1 * 0.287195224704062), 0, 0, 0);
            polyextra45.AddVertexAt(6, new Point2d(scale1 * 0.4114941743289, scale1 * 0.291319964404062), 0, 0, 0);
            polyextra45.Closed = true;
            polyextra45.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra45.Layer = "0";
            polyextra45.Color = color_gm;
            polyextra45.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra45);

            Polyline polyextra46 = new Polyline();
            polyextra46.AddVertexAt(0, new Point2d(scale1 * 0.435636986295567, scale1 * 0.283954357764062), 0, 0, 0);
            polyextra46.AddVertexAt(1, new Point2d(scale1 * 0.435227786095567, scale1 * 0.279829618034062), 0, 0, 0);
            polyextra46.AddVertexAt(2, new Point2d(scale1 * 0.4397289883289, scale1 * 0.276588751094062), 0, 0, 0);
            polyextra46.AddVertexAt(3, new Point2d(scale1 * 0.445457791162234, scale1 * 0.277472623904062), 0, 0, 0);
            polyextra46.AddVertexAt(4, new Point2d(scale1 * 0.4466853917789, scale1 * 0.281891987904062), 0, 0, 0);
            polyextra46.AddVertexAt(5, new Point2d(scale1 * 0.4458669913789, scale1 * 0.285427479104062), 0, 0, 0);
            polyextra46.AddVertexAt(6, new Point2d(scale1 * 0.4389105879289, scale1 * 0.287195224704062), 0, 0, 0);
            polyextra46.Closed = true;
            polyextra46.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra46.Layer = "0";
            polyextra46.Color = color_gm;
            polyextra46.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra46);

            Polyline polyextra47 = new Polyline();
            polyextra47.AddVertexAt(0, new Point2d(scale1 * 0.508572920512234, scale1 * 0.290757499904062), 0, 0, 0);
            polyextra47.AddVertexAt(1, new Point2d(scale1 * 0.5110281217289, scale1 * 0.286927384504062), 0, 0, 0);
            polyextra47.AddVertexAt(2, new Point2d(scale1 * 0.517984525112234, scale1 * 0.285454263104062), 0, 0, 0);
            polyextra47.AddVertexAt(3, new Point2d(scale1 * 0.5233041277789, scale1 * 0.288105881504062), 0, 0, 0);
            polyextra47.AddVertexAt(4, new Point2d(scale1 * 0.5233041277789, scale1 * 0.293703742604062), 0, 0, 0);
            polyextra47.AddVertexAt(5, new Point2d(scale1 * 0.512664922612234, scale1 * 0.295176863904062), 0, 0, 0);
            polyextra47.Closed = true;
            polyextra47.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra47.Layer = "0";
            polyextra47.Color = color_gm;
            polyextra47.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra47);

            Polyline polyextra48 = new Polyline();
            polyextra48.AddVertexAt(0, new Point2d(scale1 * 0.470476498567611, scale1 * 0.3), 0, 0, 0);
            polyextra48.AddVertexAt(1, new Point2d(scale1 * 0.468230042329524, scale1 * 0.298498082924634), 0, 0, 0);
            polyextra48.AddVertexAt(2, new Point2d(scale1 * 0.464547240503598, scale1 * 0.287596985064447), 0, 0, 0);
            polyextra48.AddVertexAt(3, new Point2d(scale1 * 0.464547240503598, scale1 * 0.282293748259544), 0, 0, 0);
            polyextra48.AddVertexAt(4, new Point2d(scale1 * 0.463319639894956, scale1 * 0.27080340186134), 0, 0, 0);
            polyextra48.AddVertexAt(5, new Point2d(scale1 * 0.469457642938166, scale1 * 0.265794789344072), 0, 0, 0);
            polyextra48.AddVertexAt(6, new Point2d(scale1 * 0.478050847198659, scale1 * 0.264027043730021), 0, 0, 0);
            polyextra48.AddVertexAt(7, new Point2d(scale1 * 0.498101657138516, scale1 * 0.264910916537046), 0, 0, 0);
            polyextra48.AddVertexAt(8, new Point2d(scale1 * 0.503421259774671, scale1 * 0.268151783477515), 0, 0, 0);
            polyextra48.AddVertexAt(9, new Point2d(scale1 * 0.5050580605849, scale1 * 0.276106638666243), 0, 0, 0);
            polyextra48.AddVertexAt(10, new Point2d(scale1 * 0.499738457948745, scale1 * 0.280526002664119), 0, 0, 0);
            polyextra48.AddVertexAt(11, new Point2d(scale1 * 0.494828055514178, scale1 * 0.285534615200013), 0, 0, 0);
            polyextra48.AddVertexAt(12, new Point2d(scale1 * 0.490736053486665, scale1 * 0.29290022186935), 0, 0, 0);
            polyextra48.AddVertexAt(13, new Point2d(scale1 * 0.487462451862327, scale1 * 0.298203458655626), 0, 0, 0);
            polyextra48.AddVertexAt(14, new Point2d(scale1 * 0.481046235397932, scale1 * 0.3), 0, 0, 0);
            polyextra48.Closed = true;
            polyextra48.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra48.Layer = "0";
            polyextra48.Color = color_gm;
            polyextra48.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra48);

            Polyline polyextra49 = new Polyline();
            polyextra49.AddVertexAt(0, new Point2d(scale1 * 0.500236982319038, scale1 * 0.244555058125407), 0, 0, 0);
            polyextra49.AddVertexAt(1, new Point2d(scale1 * 0.500646182520626, scale1 * 0.236894827187061), 0, 0, 0);
            polyextra49.AddVertexAt(2, new Point2d(scale1 * 0.502282983334735, scale1 * 0.231002341844142), 0, 0, 0);
            polyextra49.AddVertexAt(3, new Point2d(scale1 * 0.509648586986586, scale1 * 0.229234596248716), 0, 0, 0);
            polyextra49.AddVertexAt(4, new Point2d(scale1 * 0.517014190634557, scale1 * 0.230413093324751), 0, 0, 0);
            polyextra49.AddVertexAt(5, new Point2d(scale1 * 0.523152193677767, scale1 * 0.234832457322627), 0, 0, 0);
            polyextra49.AddVertexAt(6, new Point2d(scale1 * 0.517014190634557, scale1 * 0.245144306644797), 0, 0, 0);
            polyextra49.AddVertexAt(7, new Point2d(scale1 * 0.503044230807573, scale1 * 0.251867096740753), 0, 0, 0);
            polyextra49.Closed = true;
            polyextra49.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra49.Layer = "0";
            polyextra49.Color = color_gm;
            polyextra49.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra49);

            Polyline polyextra50 = new Polyline();
            polyextra50.AddVertexAt(0, new Point2d(scale1 * 0.447292868791458, scale1 * 0.24195700775832), 0, 0, 0);
            polyextra50.AddVertexAt(1, new Point2d(scale1 * 0.438000975120813, scale1 * 0.23713588340208), 0, 0, 0);
            polyextra50.AddVertexAt(2, new Point2d(scale1 * 0.436142596387459, scale1 * 0.230707717593759), 0, 0, 0);
            polyextra50.AddVertexAt(3, new Point2d(scale1 * 0.445434490058105, scale1 * 0.228832835890353), 0, 0, 0);
            polyextra50.AddVertexAt(4, new Point2d(scale1 * 0.456956438206447, scale1 * 0.230975557826459), 0, 0, 0);
            polyextra50.AddVertexAt(5, new Point2d(scale1 * 0.463646601651174, scale1 * 0.236600202918053), 0, 0, 0);
            polyextra50.AddVertexAt(6, new Point2d(scale1 * 0.458071465447235, scale1 * 0.241689167525619), 0, 0, 0);
            polyextra50.Closed = true;
            polyextra50.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra50.Layer = "0";
            polyextra50.Color = color_gm;
            polyextra50.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra50);

            Polyline polyextra51 = new Polyline();
            polyextra51.AddVertexAt(0, new Point2d(scale1 * 0.40699297210124, scale1 * 0.240644590575248), 0, 0, 0);
            polyextra51.AddVertexAt(1, new Point2d(scale1 * 0.398808968046215, scale1 * 0.242706960439682), 0, 0, 0);
            polyextra51.AddVertexAt(2, new Point2d(scale1 * 0.392261764801418, scale1 * 0.248010197244585), 0, 0, 0);
            polyextra51.AddVertexAt(3, new Point2d(scale1 * 0.373847755671789, scale1 * 0.256554300971329), 0, 0, 0);
            polyextra51.AddVertexAt(4, new Point2d(scale1 * 0.364026950802654, scale1 * 0.264803780447692), 0, 0, 0);
            polyextra51.AddVertexAt(5, new Point2d(scale1 * 0.366482152019938, scale1 * 0.268633895907551), 0, 0, 0);
            polyextra51.AddVertexAt(6, new Point2d(scale1 * 0.381213359323641, scale1 * 0.276294126845896), 0, 0, 0);
            polyextra51.AddVertexAt(7, new Point2d(scale1 * 0.392670965003005, scale1 * 0.277767248172313), 0, 0, 0);
            polyextra51.AddVertexAt(8, new Point2d(scale1 * 0.400036568654856, scale1 * 0.275115629769862), 0, 0, 0);
            polyextra51.AddVertexAt(9, new Point2d(scale1 * 0.409857373520111, scale1 * 0.27246401136741), 0, 0, 0);
            polyextra51.AddVertexAt(10, new Point2d(scale1 * 0.417632177373549, scale1 * 0.270990890040994), 0, 0, 0);
            polyextra51.AddVertexAt(11, new Point2d(scale1 * 0.42213337960653, scale1 * 0.268044647369534), 0, 0, 0);
            polyextra51.AddVertexAt(12, new Point2d(scale1 * 0.435636986301591, scale1 * 0.264509156178683), 0, 0, 0);
            polyextra51.AddVertexAt(13, new Point2d(scale1 * 0.441365789139333, scale1 * 0.254786555375904), 0, 0, 0);
            polyextra51.AddVertexAt(14, new Point2d(scale1 * 0.435636986301591, scale1 * 0.245358578842133), 0, 0, 0);
            polyextra51.AddVertexAt(15, new Point2d(scale1 * 0.42213337960653, scale1 * 0.241823087651283), 0, 0, 0);
            polyextra51.Closed = true;
            polyextra51.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra51.Layer = "0";
            polyextra51.Color = color_gm;
            polyextra51.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra51);

            Polyline polyextra52 = new Polyline();
            polyextra52.AddVertexAt(0, new Point2d(scale1 * 0.234358575745567, scale1 * 0.294560831404062), 0, 0, 0);
            polyextra52.AddVertexAt(1, new Point2d(scale1 * 0.242133379612234, scale1 * 0.292793085804062), 0, 0, 0);
            polyextra52.AddVertexAt(2, new Point2d(scale1 * 0.2441793806289, scale1 * 0.289846843104062), 0, 0, 0);
            polyextra52.AddVertexAt(3, new Point2d(scale1 * 0.242133379612234, scale1 * 0.286016727604062), 0, 0, 0);
            polyextra52.AddVertexAt(4, new Point2d(scale1 * 0.2388597779789, scale1 * 0.285722103404062), 0, 0, 0);
            polyextra52.AddVertexAt(5, new Point2d(scale1 * 0.234767775962234, scale1 * 0.287195224704062), 0, 0, 0);
            polyextra52.AddVertexAt(6, new Point2d(scale1 * 0.2314941743289, scale1 * 0.291319964404062), 0, 0, 0);
            polyextra52.Closed = true;
            polyextra52.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra52.Layer = "0";
            polyextra52.Color = color_gm;
            polyextra52.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra52);

            Polyline polyextra53 = new Polyline();
            polyextra53.AddVertexAt(0, new Point2d(scale1 * 0.255636986295567, scale1 * 0.283954357764062), 0, 0, 0);
            polyextra53.AddVertexAt(1, new Point2d(scale1 * 0.255227786095567, scale1 * 0.279829618034062), 0, 0, 0);
            polyextra53.AddVertexAt(2, new Point2d(scale1 * 0.2597289883289, scale1 * 0.276588751094062), 0, 0, 0);
            polyextra53.AddVertexAt(3, new Point2d(scale1 * 0.265457791162234, scale1 * 0.277472623904062), 0, 0, 0);
            polyextra53.AddVertexAt(4, new Point2d(scale1 * 0.2666853917789, scale1 * 0.281891987904062), 0, 0, 0);
            polyextra53.AddVertexAt(5, new Point2d(scale1 * 0.2658669913789, scale1 * 0.285427479104062), 0, 0, 0);
            polyextra53.AddVertexAt(6, new Point2d(scale1 * 0.2589105879289, scale1 * 0.287195224704062), 0, 0, 0);
            polyextra53.Closed = true;
            polyextra53.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra53.Layer = "0";
            polyextra53.Color = color_gm;
            polyextra53.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra53);

            Polyline polyextra54 = new Polyline();
            polyextra54.AddVertexAt(0, new Point2d(scale1 * 0.328572920512234, scale1 * 0.290757499904062), 0, 0, 0);
            polyextra54.AddVertexAt(1, new Point2d(scale1 * 0.3310281217289, scale1 * 0.286927384504062), 0, 0, 0);
            polyextra54.AddVertexAt(2, new Point2d(scale1 * 0.337984525112234, scale1 * 0.285454263104062), 0, 0, 0);
            polyextra54.AddVertexAt(3, new Point2d(scale1 * 0.3433041277789, scale1 * 0.288105881504062), 0, 0, 0);
            polyextra54.AddVertexAt(4, new Point2d(scale1 * 0.3433041277789, scale1 * 0.293703742604062), 0, 0, 0);
            polyextra54.AddVertexAt(5, new Point2d(scale1 * 0.332664922612234, scale1 * 0.295176863904062), 0, 0, 0);
            polyextra54.Closed = true;
            polyextra54.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra54.Layer = "0";
            polyextra54.Color = color_gm;
            polyextra54.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra54);

            Polyline polyextra55 = new Polyline();
            polyextra55.AddVertexAt(0, new Point2d(scale1 * 0.290476498567611, scale1 * 0.3), 0, 0, 0);
            polyextra55.AddVertexAt(1, new Point2d(scale1 * 0.288230042329524, scale1 * 0.298498082924634), 0, 0, 0);
            polyextra55.AddVertexAt(2, new Point2d(scale1 * 0.284547240503598, scale1 * 0.287596985064447), 0, 0, 0);
            polyextra55.AddVertexAt(3, new Point2d(scale1 * 0.284547240503598, scale1 * 0.282293748259544), 0, 0, 0);
            polyextra55.AddVertexAt(4, new Point2d(scale1 * 0.283319639894956, scale1 * 0.27080340186134), 0, 0, 0);
            polyextra55.AddVertexAt(5, new Point2d(scale1 * 0.289457642938166, scale1 * 0.265794789344072), 0, 0, 0);
            polyextra55.AddVertexAt(6, new Point2d(scale1 * 0.298050847198659, scale1 * 0.264027043730021), 0, 0, 0);
            polyextra55.AddVertexAt(7, new Point2d(scale1 * 0.318101657138516, scale1 * 0.264910916537046), 0, 0, 0);
            polyextra55.AddVertexAt(8, new Point2d(scale1 * 0.323421259774671, scale1 * 0.268151783477515), 0, 0, 0);
            polyextra55.AddVertexAt(9, new Point2d(scale1 * 0.3250580605849, scale1 * 0.276106638666243), 0, 0, 0);
            polyextra55.AddVertexAt(10, new Point2d(scale1 * 0.319738457948745, scale1 * 0.280526002664119), 0, 0, 0);
            polyextra55.AddVertexAt(11, new Point2d(scale1 * 0.314828055514178, scale1 * 0.285534615200013), 0, 0, 0);
            polyextra55.AddVertexAt(12, new Point2d(scale1 * 0.310736053486665, scale1 * 0.29290022186935), 0, 0, 0);
            polyextra55.AddVertexAt(13, new Point2d(scale1 * 0.307462451862327, scale1 * 0.298203458655626), 0, 0, 0);
            polyextra55.AddVertexAt(14, new Point2d(scale1 * 0.301046235397932, scale1 * 0.3), 0, 0, 0);
            polyextra55.Closed = true;
            polyextra55.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra55.Layer = "0";
            polyextra55.Color = color_gm;
            polyextra55.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra55);

            Polyline polyextra56 = new Polyline();
            polyextra56.AddVertexAt(0, new Point2d(scale1 * 0.320236982319038, scale1 * 0.244555058125407), 0, 0, 0);
            polyextra56.AddVertexAt(1, new Point2d(scale1 * 0.320646182520626, scale1 * 0.236894827187061), 0, 0, 0);
            polyextra56.AddVertexAt(2, new Point2d(scale1 * 0.322282983334735, scale1 * 0.231002341844142), 0, 0, 0);
            polyextra56.AddVertexAt(3, new Point2d(scale1 * 0.329648586986586, scale1 * 0.229234596248716), 0, 0, 0);
            polyextra56.AddVertexAt(4, new Point2d(scale1 * 0.337014190634557, scale1 * 0.230413093324751), 0, 0, 0);
            polyextra56.AddVertexAt(5, new Point2d(scale1 * 0.343152193677767, scale1 * 0.234832457322627), 0, 0, 0);
            polyextra56.AddVertexAt(6, new Point2d(scale1 * 0.337014190634557, scale1 * 0.245144306644797), 0, 0, 0);
            polyextra56.AddVertexAt(7, new Point2d(scale1 * 0.323044230807573, scale1 * 0.251867096740753), 0, 0, 0);
            polyextra56.Closed = true;
            polyextra56.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra56.Layer = "0";
            polyextra56.Color = color_gm;
            polyextra56.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra56);

            Polyline polyextra57 = new Polyline();
            polyextra57.AddVertexAt(0, new Point2d(scale1 * 0.267292868791458, scale1 * 0.24195700775832), 0, 0, 0);
            polyextra57.AddVertexAt(1, new Point2d(scale1 * 0.258000975120813, scale1 * 0.23713588340208), 0, 0, 0);
            polyextra57.AddVertexAt(2, new Point2d(scale1 * 0.25614259638746, scale1 * 0.230707717593759), 0, 0, 0);
            polyextra57.AddVertexAt(3, new Point2d(scale1 * 0.265434490058105, scale1 * 0.228832835890353), 0, 0, 0);
            polyextra57.AddVertexAt(4, new Point2d(scale1 * 0.276956438206447, scale1 * 0.230975557826459), 0, 0, 0);
            polyextra57.AddVertexAt(5, new Point2d(scale1 * 0.283646601651174, scale1 * 0.236600202918053), 0, 0, 0);
            polyextra57.AddVertexAt(6, new Point2d(scale1 * 0.278071465447235, scale1 * 0.241689167525619), 0, 0, 0);
            polyextra57.Closed = true;
            polyextra57.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra57.Layer = "0";
            polyextra57.Color = color_gm;
            polyextra57.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra57);

            Polyline polyextra58 = new Polyline();
            polyextra58.AddVertexAt(0, new Point2d(scale1 * 0.22699297210124, scale1 * 0.240644590575248), 0, 0, 0);
            polyextra58.AddVertexAt(1, new Point2d(scale1 * 0.218808968046214, scale1 * 0.242706960439682), 0, 0, 0);
            polyextra58.AddVertexAt(2, new Point2d(scale1 * 0.212261764801418, scale1 * 0.248010197244585), 0, 0, 0);
            polyextra58.AddVertexAt(3, new Point2d(scale1 * 0.193847755671789, scale1 * 0.256554300971329), 0, 0, 0);
            polyextra58.AddVertexAt(4, new Point2d(scale1 * 0.184026950802654, scale1 * 0.264803780447692), 0, 0, 0);
            polyextra58.AddVertexAt(5, new Point2d(scale1 * 0.186482152019938, scale1 * 0.268633895907551), 0, 0, 0);
            polyextra58.AddVertexAt(6, new Point2d(scale1 * 0.201213359323641, scale1 * 0.276294126845896), 0, 0, 0);
            polyextra58.AddVertexAt(7, new Point2d(scale1 * 0.212670965003005, scale1 * 0.277767248172313), 0, 0, 0);
            polyextra58.AddVertexAt(8, new Point2d(scale1 * 0.220036568654856, scale1 * 0.275115629769862), 0, 0, 0);
            polyextra58.AddVertexAt(9, new Point2d(scale1 * 0.229857373520111, scale1 * 0.27246401136741), 0, 0, 0);
            polyextra58.AddVertexAt(10, new Point2d(scale1 * 0.237632177373549, scale1 * 0.270990890040994), 0, 0, 0);
            polyextra58.AddVertexAt(11, new Point2d(scale1 * 0.24213337960653, scale1 * 0.268044647369534), 0, 0, 0);
            polyextra58.AddVertexAt(12, new Point2d(scale1 * 0.255636986301591, scale1 * 0.264509156178683), 0, 0, 0);
            polyextra58.AddVertexAt(13, new Point2d(scale1 * 0.261365789139333, scale1 * 0.254786555375904), 0, 0, 0);
            polyextra58.AddVertexAt(14, new Point2d(scale1 * 0.255636986301591, scale1 * 0.245358578842133), 0, 0, 0);
            polyextra58.AddVertexAt(15, new Point2d(scale1 * 0.24213337960653, scale1 * 0.241823087651283), 0, 0, 0);
            polyextra58.Closed = true;
            polyextra58.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra58.Layer = "0";
            polyextra58.Color = color_gm;
            polyextra58.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra58);

            Polyline polyextra59 = new Polyline();
            polyextra59.AddVertexAt(0, new Point2d(scale1 * 0.0543585757455669, scale1 * 0.294560831404062), 0, 0, 0);
            polyextra59.AddVertexAt(1, new Point2d(scale1 * 0.0621333796122336, scale1 * 0.292793085804062), 0, 0, 0);
            polyextra59.AddVertexAt(2, new Point2d(scale1 * 0.0641793806289004, scale1 * 0.289846843104062), 0, 0, 0);
            polyextra59.AddVertexAt(3, new Point2d(scale1 * 0.0621333796122336, scale1 * 0.286016727604062), 0, 0, 0);
            polyextra59.AddVertexAt(4, new Point2d(scale1 * 0.0588597779789004, scale1 * 0.285722103404062), 0, 0, 0);
            polyextra59.AddVertexAt(5, new Point2d(scale1 * 0.0547677759622338, scale1 * 0.287195224704062), 0, 0, 0);
            polyextra59.AddVertexAt(6, new Point2d(scale1 * 0.0514941743289004, scale1 * 0.291319964404062), 0, 0, 0);
            polyextra59.Closed = true;
            polyextra59.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra59.Layer = "0";
            polyextra59.Color = color_gm;
            polyextra59.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra59);

            Polyline polyextra60 = new Polyline();
            polyextra60.AddVertexAt(0, new Point2d(scale1 * 0.075636986295567, scale1 * 0.283954357764062), 0, 0, 0);
            polyextra60.AddVertexAt(1, new Point2d(scale1 * 0.075227786095567, scale1 * 0.279829618034062), 0, 0, 0);
            polyextra60.AddVertexAt(2, new Point2d(scale1 * 0.0797289883289003, scale1 * 0.276588751094062), 0, 0, 0);
            polyextra60.AddVertexAt(3, new Point2d(scale1 * 0.0854577911622341, scale1 * 0.277472623904062), 0, 0, 0);
            polyextra60.AddVertexAt(4, new Point2d(scale1 * 0.0866853917789001, scale1 * 0.281891987904062), 0, 0, 0);
            polyextra60.AddVertexAt(5, new Point2d(scale1 * 0.0858669913789001, scale1 * 0.285427479104062), 0, 0, 0);
            polyextra60.AddVertexAt(6, new Point2d(scale1 * 0.0789105879289003, scale1 * 0.287195224704062), 0, 0, 0);
            polyextra60.Closed = true;
            polyextra60.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra60.Layer = "0";
            polyextra60.Color = color_gm;
            polyextra60.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra60);

            Polyline polyextra61 = new Polyline();
            polyextra61.AddVertexAt(0, new Point2d(scale1 * 0.148572920512234, scale1 * 0.290757499904062), 0, 0, 0);
            polyextra61.AddVertexAt(1, new Point2d(scale1 * 0.1510281217289, scale1 * 0.286927384504062), 0, 0, 0);
            polyextra61.AddVertexAt(2, new Point2d(scale1 * 0.157984525112234, scale1 * 0.285454263104062), 0, 0, 0);
            polyextra61.AddVertexAt(3, new Point2d(scale1 * 0.1633041277789, scale1 * 0.288105881504062), 0, 0, 0);
            polyextra61.AddVertexAt(4, new Point2d(scale1 * 0.1633041277789, scale1 * 0.293703742604062), 0, 0, 0);
            polyextra61.AddVertexAt(5, new Point2d(scale1 * 0.152664922612234, scale1 * 0.295176863904062), 0, 0, 0);
            polyextra61.Closed = true;
            polyextra61.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra61.Layer = "0";
            polyextra61.Color = color_gm;
            polyextra61.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra61);

            Polyline polyextra62 = new Polyline();
            polyextra62.AddVertexAt(0, new Point2d(scale1 * 0.110476498567611, scale1 * 0.3), 0, 0, 0);
            polyextra62.AddVertexAt(1, new Point2d(scale1 * 0.108230042329524, scale1 * 0.298498082924634), 0, 0, 0);
            polyextra62.AddVertexAt(2, new Point2d(scale1 * 0.104547240503598, scale1 * 0.287596985064447), 0, 0, 0);
            polyextra62.AddVertexAt(3, new Point2d(scale1 * 0.104547240503598, scale1 * 0.282293748259544), 0, 0, 0);
            polyextra62.AddVertexAt(4, new Point2d(scale1 * 0.103319639894956, scale1 * 0.27080340186134), 0, 0, 0);
            polyextra62.AddVertexAt(5, new Point2d(scale1 * 0.109457642938166, scale1 * 0.265794789344072), 0, 0, 0);
            polyextra62.AddVertexAt(6, new Point2d(scale1 * 0.118050847198659, scale1 * 0.264027043730021), 0, 0, 0);
            polyextra62.AddVertexAt(7, new Point2d(scale1 * 0.138101657138516, scale1 * 0.264910916537046), 0, 0, 0);
            polyextra62.AddVertexAt(8, new Point2d(scale1 * 0.143421259774671, scale1 * 0.268151783477515), 0, 0, 0);
            polyextra62.AddVertexAt(9, new Point2d(scale1 * 0.1450580605849, scale1 * 0.276106638666243), 0, 0, 0);
            polyextra62.AddVertexAt(10, new Point2d(scale1 * 0.139738457948745, scale1 * 0.280526002664119), 0, 0, 0);
            polyextra62.AddVertexAt(11, new Point2d(scale1 * 0.134828055514178, scale1 * 0.285534615200013), 0, 0, 0);
            polyextra62.AddVertexAt(12, new Point2d(scale1 * 0.130736053486665, scale1 * 0.29290022186935), 0, 0, 0);
            polyextra62.AddVertexAt(13, new Point2d(scale1 * 0.127462451862327, scale1 * 0.298203458655626), 0, 0, 0);
            polyextra62.AddVertexAt(14, new Point2d(scale1 * 0.121046235397932, scale1 * 0.3), 0, 0, 0);
            polyextra62.Closed = true;
            polyextra62.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra62.Layer = "0";
            polyextra62.Color = color_gm;
            polyextra62.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra62);

            Polyline polyextra63 = new Polyline();
            polyextra63.AddVertexAt(0, new Point2d(scale1 * 0.140236982319038, scale1 * 0.244555058125407), 0, 0, 0);
            polyextra63.AddVertexAt(1, new Point2d(scale1 * 0.140646182520626, scale1 * 0.236894827187061), 0, 0, 0);
            polyextra63.AddVertexAt(2, new Point2d(scale1 * 0.142282983334735, scale1 * 0.231002341844142), 0, 0, 0);
            polyextra63.AddVertexAt(3, new Point2d(scale1 * 0.149648586986586, scale1 * 0.229234596248716), 0, 0, 0);
            polyextra63.AddVertexAt(4, new Point2d(scale1 * 0.157014190634557, scale1 * 0.230413093324751), 0, 0, 0);
            polyextra63.AddVertexAt(5, new Point2d(scale1 * 0.163152193677767, scale1 * 0.234832457322627), 0, 0, 0);
            polyextra63.AddVertexAt(6, new Point2d(scale1 * 0.157014190634557, scale1 * 0.245144306644797), 0, 0, 0);
            polyextra63.AddVertexAt(7, new Point2d(scale1 * 0.143044230807573, scale1 * 0.251867096740753), 0, 0, 0);
            polyextra63.Closed = true;
            polyextra63.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra63.Layer = "0";
            polyextra63.Color = color_gm;
            polyextra63.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra63);

            Polyline polyextra64 = new Polyline();
            polyextra64.AddVertexAt(0, new Point2d(scale1 * 0.087292868791458, scale1 * 0.24195700775832), 0, 0, 0);
            polyextra64.AddVertexAt(1, new Point2d(scale1 * 0.0780009751208126, scale1 * 0.23713588340208), 0, 0, 0);
            polyextra64.AddVertexAt(2, new Point2d(scale1 * 0.0761425963874596, scale1 * 0.230707717593759), 0, 0, 0);
            polyextra64.AddVertexAt(3, new Point2d(scale1 * 0.085434490058105, scale1 * 0.228832835890353), 0, 0, 0);
            polyextra64.AddVertexAt(4, new Point2d(scale1 * 0.0969564382064469, scale1 * 0.230975557826459), 0, 0, 0);
            polyextra64.AddVertexAt(5, new Point2d(scale1 * 0.103646601651174, scale1 * 0.236600202918053), 0, 0, 0);
            polyextra64.AddVertexAt(6, new Point2d(scale1 * 0.098071465447235, scale1 * 0.241689167525619), 0, 0, 0);
            polyextra64.Closed = true;
            polyextra64.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra64.Layer = "0";
            polyextra64.Color = color_gm;
            polyextra64.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra64);

            Polyline polyextra65 = new Polyline();
            polyextra65.AddVertexAt(0, new Point2d(scale1 * 0.0469929721012403, scale1 * 0.240644590575248), 0, 0, 0);
            polyextra65.AddVertexAt(1, new Point2d(scale1 * 0.0388089680462145, scale1 * 0.242706960439682), 0, 0, 0);
            polyextra65.AddVertexAt(2, new Point2d(scale1 * 0.0322617648014178, scale1 * 0.248010197244585), 0, 0, 0);
            polyextra65.AddVertexAt(3, new Point2d(scale1 * 0.0138477556717893, scale1 * 0.256554300971329), 0, 0, 0);
            polyextra65.AddVertexAt(4, new Point2d(scale1 * 0.00402695080265403, scale1 * 0.264803780447692), 0, 0, 0);
            polyextra65.AddVertexAt(5, new Point2d(scale1 * 0.00648215201993785, scale1 * 0.268633895907551), 0, 0, 0);
            polyextra65.AddVertexAt(6, new Point2d(scale1 * 0.0212133593236408, scale1 * 0.276294126845896), 0, 0, 0);
            polyextra65.AddVertexAt(7, new Point2d(scale1 * 0.0326709650030048, scale1 * 0.277767248172313), 0, 0, 0);
            polyextra65.AddVertexAt(8, new Point2d(scale1 * 0.0400365686548563, scale1 * 0.275115629769862), 0, 0, 0);
            polyextra65.AddVertexAt(9, new Point2d(scale1 * 0.0498573735201111, scale1 * 0.27246401136741), 0, 0, 0);
            polyextra65.AddVertexAt(10, new Point2d(scale1 * 0.0576321773735495, scale1 * 0.270990890040994), 0, 0, 0);
            polyextra65.AddVertexAt(11, new Point2d(scale1 * 0.06213337960653, scale1 * 0.268044647369534), 0, 0, 0);
            polyextra65.AddVertexAt(12, new Point2d(scale1 * 0.0756369863015911, scale1 * 0.264509156178683), 0, 0, 0);
            polyextra65.AddVertexAt(13, new Point2d(scale1 * 0.0813657891393329, scale1 * 0.254786555375904), 0, 0, 0);
            polyextra65.AddVertexAt(14, new Point2d(scale1 * 0.0756369863015911, scale1 * 0.245358578842133), 0, 0, 0);
            polyextra65.AddVertexAt(15, new Point2d(scale1 * 0.06213337960653, scale1 * 0.241823087651283), 0, 0, 0);
            polyextra65.Closed = true;
            polyextra65.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));

            polyextra65.Layer = "0";
            polyextra65.Color = color_gm;
            polyextra65.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polyextra65);

            #endregion


        }


        private void add_pattern_topsoil(BlockTableRecord bltrec1, double scale1, double graph_vexag, double stick_vexag, Polyline poly1, BlockTableRecord BTrecord, Autodesk.AutoCAD.DatabaseServices.Transaction Trans1)
        {

            #region TOPSOIL

            Autodesk.AutoCAD.Colors.Color color_ts = Autodesk.AutoCAD.Colors.Color.FromRgb(127, 95, 93);



            double len_line1 = scale1 / 60;
            double pattern_height = scale1 / 85.7;

            double spc_h_edge = scale1 / 600;
            double spc_v_edge = scale1 / 600;
            double spc_hor = scale1 / 120;
            double spc_ver = scale1 / 300;

            int nr_col = 0;
            int nr_rows = 0;

            double x1 = poly1.GetPoint2dAt(3).X + spc_h_edge;
            double y1 = poly1.GetPoint2dAt(3).Y + spc_v_edge;


            double stick_width = poly1.GetPoint2dAt(1).X - poly1.GetPoint2dAt(0).X;
            double rectangle_height = poly1.GetPoint2dAt(1).Y - poly1.GetPoint2dAt(2).Y;

            if (rectangle_height >= pattern_height + 2 * spc_v_edge)
            {
                double nr1 = Math.Floor((rectangle_height - 2 * spc_v_edge) / (pattern_height + spc_ver));
                nr_rows = Convert.ToInt32(nr1);

                if (stick_width - 2 * spc_h_edge < len_line1 + 2 * spc_hor)
                {
                    nr_col = 1;
                }
                else
                {
                    double nr2 = Math.Floor((stick_width - 2 * spc_h_edge) / (len_line1 + spc_hor));

                    nr_col = Convert.ToInt32(nr2);
                }

                double dif_len = stick_width - (nr_col * (len_line1 + spc_hor) - spc_hor);
                double dif_hght = rectangle_height - (nr_rows * (pattern_height + spc_ver) - spc_ver);

                if (nr_rows > 0 && nr_col > 0)
                {
                    for (int m = 0; m < nr_col; ++m)
                    {
                        for (int n = 0; n < nr_rows; ++n)
                        {
                            double x2 = x1 + dif_len / 2 + m * (len_line1 + spc_hor);
                            double y2 = y1 + dif_hght / 2 + n * (pattern_height + spc_ver);

                            double x3 = x2 + len_line1;
                            double y3 = y2;

                            Polyline poly_down = new Polyline();
                            poly_down.AddVertexAt(0, new Point2d(x2, y2), 0, 0, 0);
                            poly_down.AddVertexAt(1, new Point2d(x3, y3), 0, 0, 0);
                            poly_down.Layer = "0";
                            poly_down.Color = color_ts;
                            poly_down.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(poly_down);

                            Polyline poly_left1 = new Polyline();
                            poly_left1.AddVertexAt(0, new Point2d(x2 + scale1 / 300, y2 + scale1 / 300), 0, 0, 0);
                            poly_left1.AddVertexAt(1, new Point2d(x2, y2 + scale1 / 120), 0, 0, 0);
                            poly_left1.Layer = "0";
                            poly_left1.Color = color_ts;
                            poly_left1.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(poly_left1);

                            Polyline poly_left2 = new Polyline();
                            poly_left2.AddVertexAt(0, new Point2d(x2 + scale1 / 162.162, y2 + scale1 / 180.18), 0, 0, 0);
                            poly_left2.AddVertexAt(1, new Point2d(x2 + scale1 / 214.28, y2 + scale1 / 85.71), 0, 0, 0);
                            poly_left2.Layer = "0";
                            poly_left2.Color = color_ts;
                            poly_left2.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(poly_left2);

                            Polyline poly_right1 = new Polyline();
                            poly_right1.AddVertexAt(0, new Point2d(x2 + scale1 / 75, y2 + scale1 / 300), 0, 0, 0);
                            poly_right1.AddVertexAt(1, new Point2d(x2 + scale1 / 60, y2 + scale1 / 120), 0, 0, 0);
                            poly_right1.Layer = "0";
                            poly_right1.Color = color_ts;
                            poly_right1.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(poly_right1);

                            Polyline poly_right2 = new Polyline();
                            poly_right2.AddVertexAt(0, new Point2d(x2 + scale1 / 107.14, y2 + scale1 / 181.181), 0, 0, 0);
                            poly_right2.AddVertexAt(1, new Point2d(x2 + scale1 / 92.3, y2 + scale1 / 85.7), 0, 0, 0);
                            poly_right2.Layer = "0";
                            poly_right2.Color = color_ts;
                            poly_right2.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(poly_right2);


                        }
                    }

                    Polyline poly2 = new Polyline();
                    poly2 = poly1.Clone() as Polyline;
                    BTrecord.AppendEntity(poly2);
                    Trans1.AddNewlyCreatedDBObject(poly2, true);

                    Hatch hatch1 = CreateHatch(poly2, "DOTS", scale1 / 5, 22);
                    hatch1.Layer = "0";
                    hatch1.Color = color_ts;
                    hatch1.LineWeight = LineWeight.LineWeight000;
                    bltrec1.AppendEntity(hatch1);
                    poly2.Erase();

                }
            }
            else
            {
            }
            #endregion

        }


        private void add_pattern_sandstone(BlockTableRecord bltrec1, double scale1, double graph_vexag, double stick_vexag, Polyline poly1, BlockTableRecord BTrecord, Autodesk.AutoCAD.DatabaseServices.Transaction Trans1)

        {
            #region SANDSTONE

            Autodesk.AutoCAD.Colors.Color color1 = Autodesk.AutoCAD.Colors.Color.FromRgb(121, 178, 6);
            double rec_hght = poly1.GetPoint2dAt(1).Y - poly1.GetPoint2dAt(2).Y;
            double spc = scale1 / 12;
            int nr_col = 0;
            double rad1 = scale1 / 120;
            int nr_rows = 0;
            double spc_h_edge = scale1 / 600;
            double spc_v_edge = scale1 / 600;


            double x1 = poly1.GetPoint2dAt(3).X + spc_h_edge;
            double y1 = poly1.GetPoint2dAt(3).Y + spc_v_edge;
            double stick_width = poly1.GetPoint2dAt(1).X - poly1.GetPoint2dAt(0).X;

            if (rec_hght >= rad1 && stick_width >= rad1)
            {
                if (rec_hght < spc + rad1)
                {
                    nr_rows = 1;
                }
                else
                {
                    nr_rows = Convert.ToInt32(Math.Round(rec_hght / spc, 0, MidpointRounding.AwayFromZero) + 1);
                }
                if (stick_width < spc + rad1)
                {
                    nr_col = 1;
                }
                else
                {
                    nr_col = Convert.ToInt32(Math.Round(stick_width / spc, 0, MidpointRounding.AwayFromZero) + 1);
                }
                double dif1 = stick_width - (nr_col - 1) * spc;
                double dif2 = rec_hght - (nr_rows - 1) * spc;
                if (dif1 < 2 * rad1)
                {
                    nr_col = nr_col - 1;
                    dif1 = stick_width - (nr_col - 1) * spc;
                }
                if (dif2 < 2 * rad1)
                {
                    nr_rows = nr_rows - 1;
                    dif2 = rec_hght - (nr_rows - 1) * spc;
                }
                if (nr_rows > 0 && nr_col > 0)
                {
                    for (int m = 0; m < nr_col; ++m)
                    {
                        for (int n = 0; n < nr_rows; ++n)
                        {
                            double x2 = x1 + dif1 / 2 + m * spc;
                            double y2 = y1 + dif2 / 2 + n * spc;
                            Circle c1 = new Circle(new Point3d(x2, y2, 0), Vector3d.ZAxis, rad1);
                            BTrecord.AppendEntity(c1);
                            Trans1.AddNewlyCreatedDBObject(c1, true);
                            Hatch hatch1 = CreateHatch(c1, "SOLID", 1, 0);
                            hatch1.Layer = "0";
                            hatch1.Color = color1;
                            hatch1.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(hatch1);
                            c1.Erase();
                        }
                    }
                }
            }
            else
            {
            }
            #endregion
        }

        private void add_pattern_SM(BlockTableRecord bltrec1, double scale1, double graph_vexag, double stick_vexag, Polyline poly1, BlockTableRecord BTrecord, Autodesk.AutoCAD.DatabaseServices.Transaction Trans1)
        {


            Autodesk.AutoCAD.Colors.Color color1 = Autodesk.AutoCAD.Colors.Color.FromRgb(0, 255, 0);


            double spc_linie = scale1 / 14.2;
            int nr_col = 0;
            int nr_rows = 1;
            double spc_h_edge = 0;
            double spc_v_edge = 0;

            double x1 = poly1.GetPoint2dAt(3).X + spc_h_edge;
            double y1 = poly1.GetPoint2dAt(3).Y + spc_v_edge;
            double stick_width = poly1.GetPoint2dAt(1).X - poly1.GetPoint2dAt(0).X;
            double stick_height = poly1.GetPoint2dAt(1).Y - poly1.GetPoint2dAt(2).Y;

            if (stick_width < spc_linie)
            {
                nr_col = 1;
            }
            else
            {
                double nr2 = Math.Ceiling(stick_width / spc_linie);

                nr_col = Convert.ToInt32(nr2);
            }



            double dif_len = stick_width - ((nr_col - 1) * spc_linie);


            if (nr_rows > 0 && nr_col > 0)
            {
                for (int m = 0; m < nr_col; ++m)
                {
                    for (int n = 0; n < nr_rows; ++n)
                    {
                        double x2 = x1 + m * spc_linie + dif_len / 2;
                        double y2 = y1;

                        double x3 = x2;
                        double y3 = y2 + stick_height;

                        Polyline poly_up_down = new Polyline();
                        poly_up_down.AddVertexAt(0, new Point2d(x2, y2), 0, 0, 0);
                        poly_up_down.AddVertexAt(1, new Point2d(x3, y3), 0, 0, 0);
                        poly_up_down.Layer = "0";
                        poly_up_down.Color = color1;
                        poly_up_down.LineWeight = LineWeight.LineWeight000;
                        bltrec1.AppendEntity(poly_up_down);
                    }
                }
            }

            double spc_ver = scale1 * 0.01;
            double spc_hor = scale1 * 0.01;
            double r1 = 0.0015 * scale1;

            nr_col = 0;
            nr_rows = 0;


            x1 = poly1.GetPoint2dAt(3).X;
            y1 = poly1.GetPoint2dAt(3).Y;


            if (stick_width < spc_hor)
            {
                nr_col = 1;
            }
            else
            {
                double nr2 = Math.Ceiling(stick_width / spc_hor);

                nr_col = Convert.ToInt32(nr2);
            }


            if (stick_height < spc_ver)
            {
                nr_rows = 1;
            }
            else
            {
                double nr2 = Math.Ceiling(stick_height / spc_ver);
                nr_rows = Convert.ToInt32(nr2);
            }

            double dif_len_h = stick_width - ((nr_col - 1) * spc_hor);
            double dif_len_v = stick_height - ((nr_rows - 1) * spc_ver);

            if (nr_rows > 0 && nr_col > 0)
            {
                for (int m = 0; m < nr_col; ++m)
                {
                    for (int n = 0; n < nr_rows; ++n)
                    {
                        double x2 = x1 + r1 + m * spc_hor;
                        double y2 = y1 + r1 + n * spc_ver;

                        Circle cerc1 = new Circle(new Point3d(x2, y2, 0), Vector3d.ZAxis, r1);
                        cerc1.TransformBy(Matrix3d.Displacement(new Point3d(x2, y2, 0).GetVectorTo(new Point3d(x2 + dif_len_h / 2, y2 + dif_len_v / 2, 0))));

                        BTrecord.AppendEntity(cerc1);
                        Trans1.AddNewlyCreatedDBObject(cerc1, true);
                        Hatch hatch1 = CreateHatch(cerc1, "SOLID", 1, 0);
                        hatch1.Layer = "0";
                        hatch1.LineWeight = LineWeight.LineWeight000;
                        hatch1.Color = color1;
                        bltrec1.AppendEntity(hatch1);
                        cerc1.Erase();
                    }
                }
            }
        }

        private void add_pattern_ML(BlockTableRecord bltrec1, double scale1, Polyline poly1, BlockTableRecord BTrecord, Autodesk.AutoCAD.DatabaseServices.Transaction Trans1)
        {
            string nume_hatch = "ANSI31";
            double hatch_scale = scale1 / 3;
            double hatch_angle = 45;
            Autodesk.AutoCAD.Colors.Color color1 = Autodesk.AutoCAD.Colors.Color.FromRgb(176, 71, 159);

            Polyline poly2 = new Polyline();
            poly2 = poly1.Clone() as Polyline;
            BTrecord.AppendEntity(poly2);
            Trans1.AddNewlyCreatedDBObject(poly2, true);

            Hatch hatch1 = CreateHatch(poly2, nume_hatch, hatch_scale, hatch_angle * Math.PI / 180);
            hatch1.Layer = "0";
            hatch1.LineWeight = LineWeight.LineWeight000;
            hatch1.Color = color1;
            bltrec1.AppendEntity(hatch1);
            poly2.Erase();

        }

        private void add_pattern_CLML(BlockTableRecord bltrec1, double scale1, Polyline poly1, BlockTableRecord BTrecord, Autodesk.AutoCAD.DatabaseServices.Transaction Trans1)
        {
            string nume_hatch = "ANSI31";
            double hatch_scale1 = 0.568 * scale1;
            double hatch_angle1 = 0;
            double hatch_scale2 = 0.6 * scale1;
            double hatch_angle2 = 45;

            Autodesk.AutoCAD.Colors.Color color1 = Autodesk.AutoCAD.Colors.Color.FromRgb(204, 0, 204);

            Polyline poly2 = new Polyline();
            poly2 = poly1.Clone() as Polyline;
            BTrecord.AppendEntity(poly2);
            Trans1.AddNewlyCreatedDBObject(poly2, true);

            Hatch hatch1 = CreateHatch(poly2, nume_hatch, hatch_scale1, hatch_angle1 * Math.PI / 180);
            hatch1.Layer = "0";
            hatch1.LineWeight = LineWeight.LineWeight000;
            hatch1.Color = color1;
            bltrec1.AppendEntity(hatch1);

            Hatch hatch2 = CreateHatch(poly2, nume_hatch, hatch_scale2, hatch_angle2 * Math.PI / 180);
            hatch2.Layer = "0";
            hatch2.LineWeight = LineWeight.LineWeight000;
            hatch2.Color = color1;
            bltrec1.AppendEntity(hatch2);

            poly2.Erase();

        }



        private void add_pattern_GPGC(BlockTableRecord bltrec1, double scale1, double graph_vexag, double stick_vexag, Polyline poly1, BlockTableRecord BTrecord, Autodesk.AutoCAD.DatabaseServices.Transaction Trans1)
        {



            Autodesk.AutoCAD.Colors.Color color_GP = Autodesk.AutoCAD.Colors.Color.FromRgb(127, 95, 93);



            double pattern_width = 0.1 * scale1;
            double pattern_height = 0.06 * scale1;

            double spc_h_edge = 0;
            double spc_v_edge = 0;
            double spc_hor = 0;
            double spc_ver = 0;

            int nr_col = 0;
            int nr_rows = 1;

            double x1 = poly1.GetPoint2dAt(3).X + spc_h_edge;
            double y1 = poly1.GetPoint2dAt(3).Y + spc_v_edge;


            double stick_width = poly1.GetPoint2dAt(1).X - poly1.GetPoint2dAt(0).X;
            double rectangle_height = poly1.GetPoint2dAt(1).Y - poly1.GetPoint2dAt(2).Y;

            if (rectangle_height >= pattern_height + 2 * spc_v_edge)
            {
                double nr1 = Math.Floor((rectangle_height - 2 * spc_v_edge) / (pattern_height + spc_ver));
                nr_rows = Convert.ToInt32(nr1);

                if (stick_width - 2 * spc_h_edge < pattern_width + 2 * spc_hor)
                {
                    nr_col = 1;
                }
                else
                {
                    double nr2 = Math.Floor((stick_width - 2 * spc_h_edge) / (pattern_width + spc_hor));

                    nr_col = Convert.ToInt32(nr2);
                }

                double dif_len = stick_width - (nr_col * (pattern_width + spc_hor) - spc_hor);
                double dif_hght = rectangle_height - (nr_rows * (pattern_height + spc_ver) - spc_ver);

                Point3d pt_ins = poly1.GetPoint3dAt(3);

                if (nr_rows > 0 && nr_col > 0)
                {
                    for (int m = 0; m < nr_col; ++m)
                    {
                        for (int n = 0; n < nr_rows; ++n)
                        {
                            double x2 = x1 + dif_len / 2 + m * (pattern_width + spc_hor);
                            double y2 = y1 + dif_hght / 2 + n * (pattern_height + spc_ver);

                            double x3 = x2 + pattern_width;
                            double y3 = y2;


                            Polyline polygp1 = get_poly_gp1(scale1);
                            polygp1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));
                            polygp1.Layer = "0";
                            polygp1.Color = color_GP;
                            polygp1.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp1);

                            Polyline polygp2 = get_poly_gp2(scale1);
                            polygp2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp2.Layer = "0";
                            polygp2.Color = color_GP;
                            polygp2.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp2);

                            Polyline polygp3 = get_poly_gp3(scale1);
                            polygp3.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp3.Layer = "0";
                            polygp3.Color = color_GP;
                            polygp3.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp3);

                            Polyline polygp4 = get_poly_gp4(scale1);
                            polygp4.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp4.Layer = "0";
                            polygp4.Color = color_GP;
                            polygp4.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp4);

                            Polyline polygp5 = get_poly_gp5(scale1);
                            polygp5.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp5.Layer = "0";
                            polygp5.Color = color_GP;
                            polygp5.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp5);

                            Polyline polygp6 = get_poly_gp6(scale1);
                            polygp6.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp6.Layer = "0";
                            polygp6.Color = color_GP;
                            polygp6.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp6);

                            Polyline polygp7 = get_poly_gp7(scale1);
                            polygp7.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp7.Layer = "0";
                            polygp7.Color = color_GP;
                            polygp7.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp7);

                            Polyline polygp8 = get_poly_gp8(scale1);
                            polygp8.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp8.Layer = "0";
                            polygp8.Color = color_GP;
                            polygp8.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp8);

                            Polyline polygp9 = get_poly_gp9(scale1);
                            polygp9.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp9.Layer = "0";
                            polygp9.Color = color_GP;
                            polygp9.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp9);

                            Polyline polygp10 = get_poly_gp10(scale1);
                            polygp10.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp10.Layer = "0";
                            polygp10.Color = color_GP;
                            polygp10.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp10);

                        }
                    }



                }
            }
            else
            {
            }
            string nume_hatch = "ANSI31";
            double hatch_scale1 = 0.045 * scale1;
            double hatch_angle1 = 0;


            Polyline poly2 = new Polyline();
            poly2.AddVertexAt(0, new Point2d((poly1.GetPoint2dAt(0).X + poly1.GetPoint2dAt(1).X) / 2, (poly1.GetPoint2dAt(0).Y + poly1.GetPoint2dAt(1).Y) / 2), 0, 0, 0);
            poly2.AddVertexAt(1, poly1.GetPoint2dAt(1), 0, 0, 0);
            poly2.AddVertexAt(2, poly1.GetPoint2dAt(2), 0, 0, 0);
            poly2.AddVertexAt(3, new Point2d((poly1.GetPoint2dAt(2).X + poly1.GetPoint2dAt(3).X) / 2, (poly1.GetPoint2dAt(2).Y + poly1.GetPoint2dAt(3).Y) / 2), 0, 0, 0);
            poly2.Closed = true;


            BTrecord.AppendEntity(poly2);
            Trans1.AddNewlyCreatedDBObject(poly2, true);

            Hatch hatch1 = CreateHatch(poly2, nume_hatch, hatch_scale1, hatch_angle1 * Math.PI / 180);
            hatch1.Layer = "0";
            hatch1.LineWeight = LineWeight.LineWeight000;
            hatch1.Color = color_GP;
            bltrec1.AppendEntity(hatch1);



            poly2.Erase();

        }


        private void add_pattern_GPGC_legend(BlockTableRecord bltrec1, double scale1, double graph_vexag, double stick_vexag, Polyline poly1, BlockTableRecord BTrecord, Autodesk.AutoCAD.DatabaseServices.Transaction Trans1)
        {



            Autodesk.AutoCAD.Colors.Color color_GP = Autodesk.AutoCAD.Colors.Color.FromRgb(127, 95, 93);



            double pattern_width = 0.1 * scale1;
            double pattern_height = 0.06 * scale1;

            double spc_h_edge = 0;
            double spc_v_edge = 0;
            double spc_hor = 0;
            double spc_ver = 0;

            int nr_col = 0;
            int nr_rows = 1;

            double x1 = poly1.GetPoint2dAt(3).X + spc_h_edge;
            double y1 = poly1.GetPoint2dAt(3).Y + spc_v_edge;


            double stick_width = poly1.GetPoint2dAt(1).X - poly1.GetPoint2dAt(0).X;
            double rectangle_height = poly1.GetPoint2dAt(1).Y - poly1.GetPoint2dAt(2).Y;

            if (rectangle_height >= pattern_height + 2 * spc_v_edge)
            {
                double nr1 = Math.Floor((rectangle_height - 2 * spc_v_edge) / (pattern_height + spc_ver));
                nr_rows = Convert.ToInt32(nr1);

                if (stick_width - 2 * spc_h_edge < pattern_width + 2 * spc_hor)
                {
                    nr_col = 1;
                }
                else
                {
                    double nr2 = Math.Floor((stick_width - 2 * spc_h_edge) / (pattern_width + spc_hor));

                    nr_col = Convert.ToInt32(nr2);
                }

                double dif_len = stick_width - (nr_col * (pattern_width + spc_hor) - spc_hor);
                double dif_hght = rectangle_height - (nr_rows * (pattern_height + spc_ver) - spc_ver);

                Point3d pt_ins = poly1.GetPoint3dAt(3);

                if (nr_rows > 0 && nr_col > 0)
                {
                    for (int m = 0; m < nr_col; ++m)
                    {
                        for (int n = 0; n < nr_rows; ++n)
                        {
                            double x2 = x1 + dif_len / 2 + m * (pattern_width + spc_hor);
                            double y2 = y1 + dif_hght / 2 + n * (pattern_height + spc_ver);

                            double x3 = x2 + pattern_width;
                            double y3 = y2;


                            Polyline polygp1 = get_poly_gp1(scale1);
                            polygp1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));
                            polygp1.Layer = "0";
                            polygp1.Color = color_GP;
                            polygp1.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp1);

                            Polyline polygp2 = get_poly_gp2(scale1);
                            polygp2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp2.Layer = "0";
                            polygp2.Color = color_GP;
                            polygp2.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp2);

                            Polyline polygp3 = get_poly_gp3(scale1);
                            polygp3.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp3.Layer = "0";
                            polygp3.Color = color_GP;
                            polygp3.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp3);

                            Polyline polygp4 = get_poly_gp4(scale1);
                            polygp4.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp4.Layer = "0";
                            polygp4.Color = color_GP;
                            polygp4.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp4);

                            Polyline polygp5 = get_poly_gp5(scale1);
                            polygp5.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp5.Layer = "0";
                            polygp5.Color = color_GP;
                            polygp5.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp5);

                            Polyline polygp6 = get_poly_gp6(scale1);
                            polygp6.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp6.Layer = "0";
                            polygp6.Color = color_GP;
                            polygp6.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp6);

                            Polyline polygp7 = get_poly_gp7(scale1);
                            polygp7.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp7.Layer = "0";
                            polygp7.Color = color_GP;
                            polygp7.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp7);

                            Polyline polygp8 = get_poly_gp8(scale1);
                            polygp8.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp8.Layer = "0";
                            polygp8.Color = color_GP;
                            polygp8.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp8);

                            Polyline polygp9 = get_poly_gp9(scale1);
                            polygp9.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp9.Layer = "0";
                            polygp9.Color = color_GP;
                            polygp9.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp9);

                            Polyline polygp10 = get_poly_gp10(scale1);
                            polygp10.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(pt_ins.X + m * pattern_width, pt_ins.Y + n * pattern_height, 0))));

                            polygp10.Layer = "0";
                            polygp10.Color = color_GP;
                            polygp10.LineWeight = LineWeight.LineWeight000;
                            bltrec1.AppendEntity(polygp10);

                        }
                    }



                }
            }
            else
            {
            }

            #region poly extra for filling the gap
            Polyline polye1 = new Polyline();
            polye1.AddVertexAt(0, new Point2d(scale1 * 1.46028574100171, scale1 * 0.3), 0, 0, 0);
            polye1.AddVertexAt(1, new Point2d(scale1 * 1.46164016907527, scale1 * 0.297968357889651), 0, 0, 0);
            polye1.AddVertexAt(2, new Point2d(scale1 * 1.46467928719744, scale1 * 0.297687823909144), 0, 0, 0);
            polye1.AddVertexAt(3, new Point2d(scale1 * 1.46818596195379, scale1 * 0.29937102779219), 0, 0, 0);
            polye1.AddVertexAt(4, new Point2d(scale1 * 1.46818596195379, scale1 * 0.3), 0, 0, 0);
            polye1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            polye1.Layer = "0";
            polye1.Color = color_GP;
            polye1.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polye1);
            Polyline polye2 = new Polyline();
            polye2.AddVertexAt(0, new Point2d(scale1 * 1.44565464779079, scale1 * 0.3), 0, 0, 0);
            polye2.AddVertexAt(1, new Point2d(scale1 * 1.44457435192772, scale1 * 0.297407289928637), 0, 0, 0);
            polye2.AddVertexAt(2, new Point2d(scale1 * 1.43943122895175, scale1 * 0.295583819055336), 0, 0, 0);
            polye2.AddVertexAt(3, new Point2d(scale1 * 1.42937876131689, scale1 * 0.295303285074829), 0, 0, 0);
            polye2.AddVertexAt(4, new Point2d(scale1 * 1.424469416658, scale1 * 0.296144887016352), 0, 0, 0);
            polye2.AddVertexAt(5, new Point2d(scale1 * 1.42096274190166, scale1 * 0.297828090899398), 0, 0, 0);
            polye2.AddVertexAt(6, new Point2d(scale1 * 1.41875935585757, scale1 * 0.3), 0, 0, 0);
            polye2.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            polye2.Layer = "0";
            polye2.Color = color_GP;
            polye2.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(polye2);
            Polyline poly3 = new Polyline();
            poly3.AddVertexAt(0, new Point2d(scale1 * 1.45836727263602, scale1 * 0.282994061761134), 0, 0, 0);
            poly3.AddVertexAt(1, new Point2d(scale1 * 1.4564970460993, scale1 * 0.282853794770882), 0, 0, 0);
            poly3.AddVertexAt(2, new Point2d(scale1 * 1.4541592629284, scale1 * 0.28355512972215), 0, 0, 0);
            poly3.AddVertexAt(3, new Point2d(scale1 * 1.45228903639168, scale1 * 0.285518867585705), 0, 0, 0);
            poly3.AddVertexAt(4, new Point2d(scale1 * 1.45392548461131, scale1 * 0.287061804478495), 0, 0, 0);
            poly3.AddVertexAt(5, new Point2d(scale1 * 1.45836727263602, scale1 * 0.286220202536974), 0, 0, 0);
            poly3.AddVertexAt(6, new Point2d(scale1 * 1.45953616422147, scale1 * 0.284817532634435), 0, 0, 0);
            poly3.Closed = true;
            poly3.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly3.Layer = "0";
            poly3.Color = color_GP;
            poly3.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly3);
            Polyline poly4 = new Polyline();
            poly4.AddVertexAt(0, new Point2d(scale1 * 1.4679521836367, scale1 * 0.28355512972215), 0, 0, 0);
            poly4.AddVertexAt(1, new Point2d(scale1 * 1.46608195709998, scale1 * 0.282012192829359), 0, 0, 0);
            poly4.AddVertexAt(2, new Point2d(scale1 * 1.46584817878289, scale1 * 0.280048454965803), 0, 0, 0);
            poly4.AddVertexAt(3, new Point2d(scale1 * 1.46841974027088, scale1 * 0.278505518073012), 0, 0, 0);
            poly4.AddVertexAt(4, new Point2d(scale1 * 1.47169263671013, scale1 * 0.278926319043773), 0, 0, 0);
            poly4.AddVertexAt(5, new Point2d(scale1 * 1.4723939716614, scale1 * 0.281030323897581), 0, 0, 0);
            poly4.AddVertexAt(6, new Point2d(scale1 * 1.47192641502722, scale1 * 0.282713527780627), 0, 0, 0);
            poly4.AddVertexAt(7, new Point2d(scale1 * 1.4679521836367, scale1 * 0.28355512972215), 0, 0, 0);
            poly4.Closed = true;
            poly4.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly4.Layer = "0";
            poly4.Color = color_GP;
            poly4.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly4);
            Polyline poly5 = new Polyline();
            poly5.AddVertexAt(0, new Point2d(scale1 * 1.41657711212705, scale1 * 0.287355090003573), 0, 0, 0);
            poly5.AddVertexAt(1, new Point2d(scale1 * 1.41423932895615, scale1 * 0.285251085149766), 0, 0, 0);
            poly5.AddVertexAt(2, new Point2d(scale1 * 1.41564199885869, scale1 * 0.283427614276465), 0, 0, 0);
            poly5.AddVertexAt(3, new Point2d(scale1 * 1.41961623024921, scale1 * 0.282726279325197), 0, 0, 0);
            poly5.AddVertexAt(4, new Point2d(scale1 * 1.42265534837138, scale1 * 0.283988682237481), 0, 0, 0);
            poly5.AddVertexAt(5, new Point2d(scale1 * 1.42265534837138, scale1 * 0.286653755052303), 0, 0, 0);
            poly5.AddVertexAt(6, new Point2d(scale1 * 1.41657711212705, scale1 * 0.287355090003573), 0, 0, 0);
            poly5.Closed = true;
            poly5.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly5.Layer = "0";
            poly5.Color = color_GP;
            poly5.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly5);
            Polyline poly6 = new Polyline();
            poly6.AddVertexAt(0, new Point2d(scale1 * 1.41108075824353, scale1 * 0.266735842436258), 0, 0, 0);
            poly6.AddVertexAt(1, new Point2d(scale1 * 1.40947696180731, scale1 * 0.263254670769046), 0, 0, 0);
            poly6.AddVertexAt(2, new Point2d(scale1 * 1.4097107401244, scale1 * 0.259607729022446), 0, 0, 0);
            poly6.AddVertexAt(3, new Point2d(scale1 * 1.41064585339276, scale1 * 0.256802389217368), 0, 0, 0);
            poly6.AddVertexAt(4, new Point2d(scale1 * 1.41485386310037, scale1 * 0.255960787275844), 0, 0, 0);
            poly6.AddVertexAt(5, new Point2d(scale1 * 1.41906187280799, scale1 * 0.256521855236861), 0, 0, 0);
            poly6.AddVertexAt(6, new Point2d(scale1 * 1.42256854756433, scale1 * 0.258625860090668), 0, 0, 0);
            poly6.AddVertexAt(7, new Point2d(scale1 * 1.41906187280799, scale1 * 0.263535204749552), 0, 0, 0);
            poly6.AddVertexAt(8, new Point2d(scale1 * 1.41108075824353, scale1 * 0.266735842436258), 0, 0, 0);
            poly6.Closed = true;
            poly6.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly6.Layer = "0";
            poly6.Color = color_GP;
            poly6.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly6);
            Polyline poly7 = new Polyline();
            poly7.AddVertexAt(0, new Point2d(scale1 * 1.5, scale1 * 0.282677952223438), 0, 0, 0);
            poly7.AddVertexAt(1, new Point2d(scale1 * 1.49989813913476, scale1 * 0.282764533958894), 0, 0, 0);
            poly7.AddVertexAt(2, new Point2d(scale1 * 1.49756035596386, scale1 * 0.286271208715241), 0, 0, 0);
            poly7.AddVertexAt(3, new Point2d(scale1 * 1.49569012942714, scale1 * 0.288796014539811), 0, 0, 0);
            poly7.AddVertexAt(4, new Point2d(scale1 * 1.49148211971953, scale1 * 0.289777883471588), 0, 0, 0);
            poly7.AddVertexAt(5, new Point2d(scale1 * 1.48797544496318, scale1 * 0.290759752403364), 0, 0, 0);
            poly7.AddVertexAt(6, new Point2d(scale1 * 1.48470254852392, scale1 * 0.288936281530065), 0, 0, 0);
            poly7.AddVertexAt(7, new Point2d(scale1 * 1.48259854367012, scale1 * 0.283746402890672), 0, 0, 0);
            poly7.AddVertexAt(8, new Point2d(scale1 * 1.48259854367012, scale1 * 0.281221597066103), 0, 0, 0);
            poly7.AddVertexAt(9, new Point2d(scale1 * 1.48189720871885, scale1 * 0.275751184446203), 0, 0, 0);
            poly7.AddVertexAt(10, new Point2d(scale1 * 1.48540388347519, scale1 * 0.273366645611888), 0, 0, 0);
            poly7.AddVertexAt(11, new Point2d(scale1 * 1.49031322813408, scale1 * 0.272525043670365), 0, 0, 0);
            poly7.AddVertexAt(12, new Point2d(scale1 * 1.5, scale1 * 0.272880884269521), 0, 0, 0);
            poly7.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly7.Layer = "0";
            poly7.Color = color_GP;
            poly7.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly7);
            Polyline poly8 = new Polyline();
            poly8.AddVertexAt(0, new Point2d(scale1 * 1.45836727263602, scale1 * 0.274437775355651), 0, 0, 0);
            poly8.AddVertexAt(1, new Point2d(scale1 * 1.46608195709998, scale1 * 0.272754571472604), 0, 0, 0);
            poly8.AddVertexAt(2, new Point2d(scale1 * 1.46935485353923, scale1 * 0.268125760794228), 0, 0, 0);
            poly8.AddVertexAt(3, new Point2d(scale1 * 1.46608195709998, scale1 * 0.263637217106105), 0, 0, 0);
            poly8.AddVertexAt(4, new Point2d(scale1 * 1.45836727263602, scale1 * 0.261954013223059), 0, 0, 0);
            poly8.AddVertexAt(5, new Point2d(scale1 * 1.4497174749037, scale1 * 0.261392945262043), 0, 0, 0);
            poly8.AddVertexAt(6, new Point2d(scale1 * 1.4450419085619, scale1 * 0.26237481419382), 0, 0, 0);
            poly8.AddVertexAt(7, new Point2d(scale1 * 1.44130145548847, scale1 * 0.26489962001839), 0, 0, 0);
            poly8.AddVertexAt(8, new Point2d(scale1 * 1.43078143121943, scale1 * 0.268967362735751), 0, 0, 0);
            poly8.AddVertexAt(9, new Point2d(scale1 * 1.42517075160927, scale1 * 0.272894838462859), 0, 0, 0);
            poly8.AddVertexAt(10, new Point2d(scale1 * 1.42657342151181, scale1 * 0.27471830933616), 0, 0, 0);
            poly8.AddVertexAt(11, new Point2d(scale1 * 1.43498944092704, scale1 * 0.278365251082757), 0, 0, 0);
            poly8.AddVertexAt(12, new Point2d(scale1 * 1.44153523380556, scale1 * 0.279066586034028), 0, 0, 0);
            poly8.AddVertexAt(13, new Point2d(scale1 * 1.44574324351317, scale1 * 0.277804183121743), 0, 0, 0);
            poly8.AddVertexAt(14, new Point2d(scale1 * 1.45135392312332, scale1 * 0.276541780209458), 0, 0, 0);
            poly8.AddVertexAt(15, new Point2d(scale1 * 1.45579571114803, scale1 * 0.275840445258188), 0, 0, 0);
            poly8.Closed = true;
            poly8.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly8.Layer = "0";
            poly8.Color = color_GP;
            poly8.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly8);
            Polyline poly9 = new Polyline();
            poly9.AddVertexAt(0, new Point2d(scale1 * 1.47889889791182, scale1 * 0.261890255500208), 0, 0, 0);
            poly9.AddVertexAt(1, new Point2d(scale1 * 1.47274102660718, scale1 * 0.262017770945895), 0, 0, 0);
            poly9.AddVertexAt(2, new Point2d(scale1 * 1.46743251686181, scale1 * 0.259722492923566), 0, 0, 0);
            poly9.AddVertexAt(3, new Point2d(scale1 * 1.46637081491274, scale1 * 0.256662122227115), 0, 0, 0);
            poly9.AddVertexAt(4, new Point2d(scale1 * 1.47167932465811, scale1 * 0.255769514107321), 0, 0, 0);
            poly9.AddVertexAt(5, new Point2d(scale1 * 1.47826187674237, scale1 * 0.2567896376728), 0, 0, 0);
            poly9.AddVertexAt(6, new Point2d(scale1 * 1.48208400375904, scale1 * 0.259467462032192), 0, 0, 0);
            poly9.AddVertexAt(7, new Point2d(scale1 * 1.47889889791182, scale1 * 0.261890255500208), 0, 0, 0);
            poly9.Closed = true;
            poly9.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly9.Layer = "0";
            poly9.Color = color_GP;
            poly9.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly9);
            Polyline poly10 = new Polyline();
            poly10.AddVertexAt(0, new Point2d(scale1 * 1.49647313832164, scale1 * 0.248628649148944), 0, 0, 0);
            poly10.AddVertexAt(1, new Point2d(scale1 * 1.49670691663873, scale1 * 0.243298503519296), 0, 0, 0);
            poly10.AddVertexAt(2, new Point2d(scale1 * 1.49507046841911, scale1 * 0.242036100607014), 0, 0, 0);
            poly10.AddVertexAt(3, new Point2d(scale1 * 1.49086245871149, scale1 * 0.240633430704475), 0, 0, 0);
            poly10.AddVertexAt(4, new Point2d(scale1 * 1.48688822732096, scale1 * 0.240773697694727), 0, 0, 0);
            poly10.AddVertexAt(5, new Point2d(scale1 * 1.4838491091988, scale1 * 0.242036100607014), 0, 0, 0);
            poly10.AddVertexAt(6, new Point2d(scale1 * 1.48057621275954, scale1 * 0.243298503519296), 0, 0, 0);
            poly10.AddVertexAt(7, new Point2d(scale1 * 1.47753709463737, scale1 * 0.244560906431583), 0, 0, 0);
            poly10.AddVertexAt(8, new Point2d(scale1 * 1.47566686810066, scale1 * 0.246945445265898), 0, 0, 0);
            poly10.AddVertexAt(9, new Point2d(scale1 * 1.47566686810066, scale1 * 0.248628649148944), 0, 0, 0);
            poly10.AddVertexAt(10, new Point2d(scale1 * 1.479173542857, scale1 * 0.252275590895543), 0, 0, 0);
            poly10.AddVertexAt(11, new Point2d(scale1 * 1.48595311405261, scale1 * 0.253818527788335), 0, 0, 0);
            poly10.AddVertexAt(12, new Point2d(scale1 * 1.49273268524821, scale1 * 0.252275590895543), 0, 0, 0);
            poly10.Closed = true;
            poly10.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly10.Layer = "0";
            poly10.Color = color_GP;
            poly10.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly10);
            Polyline poly11 = new Polyline();
            poly11.AddVertexAt(0, new Point2d(scale1 * 1.46210772570945, scale1 * 0.243438770509551), 0, 0, 0);
            poly11.AddVertexAt(1, new Point2d(scale1 * 1.45930238590438, scale1 * 0.241895833616759), 0, 0, 0);
            poly11.AddVertexAt(2, new Point2d(scale1 * 1.46023749917273, scale1 * 0.240072362743458), 0, 0, 0);
            poly11.AddVertexAt(3, new Point2d(scale1 * 1.46164016907527, scale1 * 0.237968357889651), 0, 0, 0);
            poly11.AddVertexAt(4, new Point2d(scale1 * 1.46467928719744, scale1 * 0.237687823909144), 0, 0, 0);
            poly11.AddVertexAt(5, new Point2d(scale1 * 1.46818596195379, scale1 * 0.23937102779219), 0, 0, 0);
            poly11.AddVertexAt(6, new Point2d(scale1 * 1.46818596195379, scale1 * 0.241615299636252), 0, 0, 0);
            poly11.AddVertexAt(7, new Point2d(scale1 * 1.46584817878289, scale1 * 0.244140105460821), 0, 0, 0);
            poly11.AddVertexAt(8, new Point2d(scale1 * 1.46210772570945, scale1 * 0.243438770509551), 0, 0, 0);
            poly11.Closed = true;
            poly11.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly11.Layer = "0";
            poly11.Color = color_GP;
            poly11.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly11);
            Polyline poly12 = new Polyline();
            poly12.AddVertexAt(0, new Point2d(scale1 * 1.41652095387695, scale1 * 0.248628649148944), 0, 0, 0);
            poly12.AddVertexAt(1, new Point2d(scale1 * 1.42306674675547, scale1 * 0.250452120022245), 0, 0, 0);
            poly12.AddVertexAt(2, new Point2d(scale1 * 1.4359245541954, scale1 * 0.249750785070974), 0, 0, 0);
            poly12.AddVertexAt(3, new Point2d(scale1 * 1.44247034707391, scale1 * 0.248628649148944), 0, 0, 0);
            poly12.AddVertexAt(4, new Point2d(scale1 * 1.44363923865936, scale1 * 0.246805178275643), 0, 0, 0);
            poly12.AddVertexAt(5, new Point2d(scale1 * 1.44387301697645, scale1 * 0.244981707402343), 0, 0, 0);
            poly12.AddVertexAt(6, new Point2d(scale1 * 1.44410679529354, scale1 * 0.242737435558282), 0, 0, 0);
            poly12.AddVertexAt(7, new Point2d(scale1 * 1.44597702183026, scale1 * 0.240773697694727), 0, 0, 0);
            poly12.AddVertexAt(8, new Point2d(scale1 * 1.44457435192772, scale1 * 0.237407289928637), 0, 0, 0);
            poly12.AddVertexAt(9, new Point2d(scale1 * 1.43943122895175, scale1 * 0.235583819055336), 0, 0, 0);
            poly12.AddVertexAt(10, new Point2d(scale1 * 1.42937876131689, scale1 * 0.235303285074829), 0, 0, 0);
            poly12.AddVertexAt(11, new Point2d(scale1 * 1.424469416658, scale1 * 0.236144887016352), 0, 0, 0);
            poly12.AddVertexAt(12, new Point2d(scale1 * 1.42096274190166, scale1 * 0.237828090899398), 0, 0, 0);
            poly12.AddVertexAt(13, new Point2d(scale1 * 1.4176898454624, scale1 * 0.241054231675236), 0, 0, 0);
            poly12.AddVertexAt(14, new Point2d(scale1 * 1.41581961892568, scale1 * 0.244420639441328), 0, 0, 0);
            poly12.Closed = true;
            poly12.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly12.Layer = "0";
            poly12.Color = color_GP;
            poly12.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly12);
            Polyline poly13 = new Polyline();
            poly13.AddVertexAt(0, new Point2d(scale1 * 1.36028574100171, scale1 * 0.3), 0, 0, 0);
            poly13.AddVertexAt(1, new Point2d(scale1 * 1.36164016907527, scale1 * 0.297968357889651), 0, 0, 0);
            poly13.AddVertexAt(2, new Point2d(scale1 * 1.36467928719744, scale1 * 0.297687823909144), 0, 0, 0);
            poly13.AddVertexAt(3, new Point2d(scale1 * 1.36818596195379, scale1 * 0.29937102779219), 0, 0, 0);
            poly13.AddVertexAt(4, new Point2d(scale1 * 1.36818596195379, scale1 * 0.3), 0, 0, 0);
            poly13.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly13.Layer = "0";
            poly13.Color = color_GP;
            poly13.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly13);
            Polyline poly14 = new Polyline();
            poly14.AddVertexAt(0, new Point2d(scale1 * 1.34565464779079, scale1 * 0.3), 0, 0, 0);
            poly14.AddVertexAt(1, new Point2d(scale1 * 1.34457435192772, scale1 * 0.297407289928637), 0, 0, 0);
            poly14.AddVertexAt(2, new Point2d(scale1 * 1.33943122895175, scale1 * 0.295583819055336), 0, 0, 0);
            poly14.AddVertexAt(3, new Point2d(scale1 * 1.32937876131689, scale1 * 0.295303285074829), 0, 0, 0);
            poly14.AddVertexAt(4, new Point2d(scale1 * 1.324469416658, scale1 * 0.296144887016352), 0, 0, 0);
            poly14.AddVertexAt(5, new Point2d(scale1 * 1.32096274190166, scale1 * 0.297828090899398), 0, 0, 0);
            poly14.AddVertexAt(6, new Point2d(scale1 * 1.31875935585757, scale1 * 0.3), 0, 0, 0);
            poly14.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly14.Layer = "0";
            poly14.Color = color_GP;
            poly14.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly14);
            Polyline poly15 = new Polyline();
            poly15.AddVertexAt(0, new Point2d(scale1 * 1.35836727263602, scale1 * 0.282994061761134), 0, 0, 0);
            poly15.AddVertexAt(1, new Point2d(scale1 * 1.3564970460993, scale1 * 0.282853794770882), 0, 0, 0);
            poly15.AddVertexAt(2, new Point2d(scale1 * 1.3541592629284, scale1 * 0.28355512972215), 0, 0, 0);
            poly15.AddVertexAt(3, new Point2d(scale1 * 1.35228903639168, scale1 * 0.285518867585705), 0, 0, 0);
            poly15.AddVertexAt(4, new Point2d(scale1 * 1.35392548461131, scale1 * 0.287061804478495), 0, 0, 0);
            poly15.AddVertexAt(5, new Point2d(scale1 * 1.35836727263602, scale1 * 0.286220202536974), 0, 0, 0);
            poly15.AddVertexAt(6, new Point2d(scale1 * 1.35953616422147, scale1 * 0.284817532634435), 0, 0, 0);
            poly15.Closed = true;
            poly15.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly15.Layer = "0";
            poly15.Color = color_GP;
            poly15.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly15);
            Polyline poly16 = new Polyline();
            poly16.AddVertexAt(0, new Point2d(scale1 * 1.3679521836367, scale1 * 0.28355512972215), 0, 0, 0);
            poly16.AddVertexAt(1, new Point2d(scale1 * 1.36608195709998, scale1 * 0.282012192829359), 0, 0, 0);
            poly16.AddVertexAt(2, new Point2d(scale1 * 1.36584817878289, scale1 * 0.280048454965803), 0, 0, 0);
            poly16.AddVertexAt(3, new Point2d(scale1 * 1.36841974027088, scale1 * 0.278505518073012), 0, 0, 0);
            poly16.AddVertexAt(4, new Point2d(scale1 * 1.37169263671013, scale1 * 0.278926319043773), 0, 0, 0);
            poly16.AddVertexAt(5, new Point2d(scale1 * 1.3723939716614, scale1 * 0.281030323897581), 0, 0, 0);
            poly16.AddVertexAt(6, new Point2d(scale1 * 1.37192641502722, scale1 * 0.282713527780627), 0, 0, 0);
            poly16.AddVertexAt(7, new Point2d(scale1 * 1.3679521836367, scale1 * 0.28355512972215), 0, 0, 0);
            poly16.Closed = true;
            poly16.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly16.Layer = "0";
            poly16.Color = color_GP;
            poly16.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly16);
            Polyline poly17 = new Polyline();
            poly17.AddVertexAt(0, new Point2d(scale1 * 1.31657711212705, scale1 * 0.287355090003573), 0, 0, 0);
            poly17.AddVertexAt(1, new Point2d(scale1 * 1.31423932895615, scale1 * 0.285251085149766), 0, 0, 0);
            poly17.AddVertexAt(2, new Point2d(scale1 * 1.31564199885869, scale1 * 0.283427614276465), 0, 0, 0);
            poly17.AddVertexAt(3, new Point2d(scale1 * 1.31961623024921, scale1 * 0.282726279325197), 0, 0, 0);
            poly17.AddVertexAt(4, new Point2d(scale1 * 1.32265534837138, scale1 * 0.283988682237481), 0, 0, 0);
            poly17.AddVertexAt(5, new Point2d(scale1 * 1.32265534837138, scale1 * 0.286653755052303), 0, 0, 0);
            poly17.AddVertexAt(6, new Point2d(scale1 * 1.31657711212705, scale1 * 0.287355090003573), 0, 0, 0);
            poly17.Closed = true;
            poly17.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly17.Layer = "0";
            poly17.Color = color_GP;
            poly17.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly17);
            Polyline poly18 = new Polyline();
            poly18.AddVertexAt(0, new Point2d(scale1 * 1.31108075824353, scale1 * 0.266735842436258), 0, 0, 0);
            poly18.AddVertexAt(1, new Point2d(scale1 * 1.30947696180731, scale1 * 0.263254670769046), 0, 0, 0);
            poly18.AddVertexAt(2, new Point2d(scale1 * 1.3097107401244, scale1 * 0.259607729022446), 0, 0, 0);
            poly18.AddVertexAt(3, new Point2d(scale1 * 1.31064585339276, scale1 * 0.256802389217368), 0, 0, 0);
            poly18.AddVertexAt(4, new Point2d(scale1 * 1.31485386310037, scale1 * 0.255960787275844), 0, 0, 0);
            poly18.AddVertexAt(5, new Point2d(scale1 * 1.31906187280799, scale1 * 0.256521855236861), 0, 0, 0);
            poly18.AddVertexAt(6, new Point2d(scale1 * 1.32256854756433, scale1 * 0.258625860090668), 0, 0, 0);
            poly18.AddVertexAt(7, new Point2d(scale1 * 1.31906187280799, scale1 * 0.263535204749552), 0, 0, 0);
            poly18.AddVertexAt(8, new Point2d(scale1 * 1.31108075824353, scale1 * 0.266735842436258), 0, 0, 0);
            poly18.Closed = true;
            poly18.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly18.Layer = "0";
            poly18.Color = color_GP;
            poly18.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly18);
            Polyline poly19 = new Polyline();
            poly19.AddVertexAt(0, new Point2d(scale1 * 1.38797544496318, scale1 * 0.290759752403364), 0, 0, 0);
            poly19.AddVertexAt(1, new Point2d(scale1 * 1.38470254852392, scale1 * 0.288936281530065), 0, 0, 0);
            poly19.AddVertexAt(2, new Point2d(scale1 * 1.38259854367012, scale1 * 0.283746402890672), 0, 0, 0);
            poly19.AddVertexAt(3, new Point2d(scale1 * 1.38259854367012, scale1 * 0.281221597066103), 0, 0, 0);
            poly19.AddVertexAt(4, new Point2d(scale1 * 1.38189720871885, scale1 * 0.275751184446203), 0, 0, 0);
            poly19.AddVertexAt(5, new Point2d(scale1 * 1.38540388347519, scale1 * 0.273366645611888), 0, 0, 0);
            poly19.AddVertexAt(6, new Point2d(scale1 * 1.39031322813408, scale1 * 0.272525043670365), 0, 0, 0);
            poly19.AddVertexAt(7, new Point2d(scale1 * 1.40176836567148, scale1 * 0.272945844641126), 0, 0, 0);
            poly19.AddVertexAt(8, new Point2d(scale1 * 1.40480748379364, scale1 * 0.274488781533918), 0, 0, 0);
            poly19.AddVertexAt(9, new Point2d(scale1 * 1.405742597062, scale1 * 0.278275990270772), 0, 0, 0);
            poly19.AddVertexAt(10, new Point2d(scale1 * 1.40270347893983, scale1 * 0.28037999512458), 0, 0, 0);
            poly19.AddVertexAt(11, new Point2d(scale1 * 1.39989813913476, scale1 * 0.282764533958894), 0, 0, 0);
            poly19.AddVertexAt(12, new Point2d(scale1 * 1.39756035596386, scale1 * 0.286271208715241), 0, 0, 0);
            poly19.AddVertexAt(13, new Point2d(scale1 * 1.39569012942714, scale1 * 0.288796014539811), 0, 0, 0);
            poly19.AddVertexAt(14, new Point2d(scale1 * 1.39148211971953, scale1 * 0.289777883471588), 0, 0, 0);
            poly19.AddVertexAt(15, new Point2d(scale1 * 1.38797544496318, scale1 * 0.290759752403364), 0, 0, 0);
            poly19.Closed = true;
            poly19.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly19.Layer = "0";
            poly19.Color = color_GP;
            poly19.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly19);
            Polyline poly20 = new Polyline();
            poly20.AddVertexAt(0, new Point2d(scale1 * 1.35836727263602, scale1 * 0.274437775355651), 0, 0, 0);
            poly20.AddVertexAt(1, new Point2d(scale1 * 1.36608195709998, scale1 * 0.272754571472604), 0, 0, 0);
            poly20.AddVertexAt(2, new Point2d(scale1 * 1.36935485353923, scale1 * 0.268125760794228), 0, 0, 0);
            poly20.AddVertexAt(3, new Point2d(scale1 * 1.36608195709998, scale1 * 0.263637217106105), 0, 0, 0);
            poly20.AddVertexAt(4, new Point2d(scale1 * 1.35836727263602, scale1 * 0.261954013223059), 0, 0, 0);
            poly20.AddVertexAt(5, new Point2d(scale1 * 1.3497174749037, scale1 * 0.261392945262043), 0, 0, 0);
            poly20.AddVertexAt(6, new Point2d(scale1 * 1.3450419085619, scale1 * 0.26237481419382), 0, 0, 0);
            poly20.AddVertexAt(7, new Point2d(scale1 * 1.34130145548847, scale1 * 0.26489962001839), 0, 0, 0);
            poly20.AddVertexAt(8, new Point2d(scale1 * 1.33078143121943, scale1 * 0.268967362735751), 0, 0, 0);
            poly20.AddVertexAt(9, new Point2d(scale1 * 1.32517075160927, scale1 * 0.272894838462859), 0, 0, 0);
            poly20.AddVertexAt(10, new Point2d(scale1 * 1.32657342151181, scale1 * 0.27471830933616), 0, 0, 0);
            poly20.AddVertexAt(11, new Point2d(scale1 * 1.33498944092704, scale1 * 0.278365251082757), 0, 0, 0);
            poly20.AddVertexAt(12, new Point2d(scale1 * 1.34153523380556, scale1 * 0.279066586034028), 0, 0, 0);
            poly20.AddVertexAt(13, new Point2d(scale1 * 1.34574324351317, scale1 * 0.277804183121743), 0, 0, 0);
            poly20.AddVertexAt(14, new Point2d(scale1 * 1.35135392312332, scale1 * 0.276541780209458), 0, 0, 0);
            poly20.AddVertexAt(15, new Point2d(scale1 * 1.35579571114803, scale1 * 0.275840445258188), 0, 0, 0);
            poly20.Closed = true;
            poly20.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly20.Layer = "0";
            poly20.Color = color_GP;
            poly20.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly20);
            Polyline poly21 = new Polyline();
            poly21.AddVertexAt(0, new Point2d(scale1 * 1.37889889791182, scale1 * 0.261890255500208), 0, 0, 0);
            poly21.AddVertexAt(1, new Point2d(scale1 * 1.37274102660718, scale1 * 0.262017770945895), 0, 0, 0);
            poly21.AddVertexAt(2, new Point2d(scale1 * 1.36743251686181, scale1 * 0.259722492923566), 0, 0, 0);
            poly21.AddVertexAt(3, new Point2d(scale1 * 1.36637081491274, scale1 * 0.256662122227115), 0, 0, 0);
            poly21.AddVertexAt(4, new Point2d(scale1 * 1.37167932465811, scale1 * 0.255769514107321), 0, 0, 0);
            poly21.AddVertexAt(5, new Point2d(scale1 * 1.37826187674237, scale1 * 0.2567896376728), 0, 0, 0);
            poly21.AddVertexAt(6, new Point2d(scale1 * 1.38208400375904, scale1 * 0.259467462032192), 0, 0, 0);
            poly21.AddVertexAt(7, new Point2d(scale1 * 1.37889889791182, scale1 * 0.261890255500208), 0, 0, 0);
            poly21.Closed = true;
            poly21.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly21.Layer = "0";
            poly21.Color = color_GP;
            poly21.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly21);
            Polyline poly22 = new Polyline();
            poly22.AddVertexAt(0, new Point2d(scale1 * 1.39647313832164, scale1 * 0.248628649148944), 0, 0, 0);
            poly22.AddVertexAt(1, new Point2d(scale1 * 1.39670691663873, scale1 * 0.243298503519296), 0, 0, 0);
            poly22.AddVertexAt(2, new Point2d(scale1 * 1.39507046841911, scale1 * 0.242036100607014), 0, 0, 0);
            poly22.AddVertexAt(3, new Point2d(scale1 * 1.39086245871149, scale1 * 0.240633430704475), 0, 0, 0);
            poly22.AddVertexAt(4, new Point2d(scale1 * 1.38688822732096, scale1 * 0.240773697694727), 0, 0, 0);
            poly22.AddVertexAt(5, new Point2d(scale1 * 1.3838491091988, scale1 * 0.242036100607014), 0, 0, 0);
            poly22.AddVertexAt(6, new Point2d(scale1 * 1.38057621275954, scale1 * 0.243298503519296), 0, 0, 0);
            poly22.AddVertexAt(7, new Point2d(scale1 * 1.37753709463737, scale1 * 0.244560906431583), 0, 0, 0);
            poly22.AddVertexAt(8, new Point2d(scale1 * 1.37566686810066, scale1 * 0.246945445265898), 0, 0, 0);
            poly22.AddVertexAt(9, new Point2d(scale1 * 1.37566686810066, scale1 * 0.248628649148944), 0, 0, 0);
            poly22.AddVertexAt(10, new Point2d(scale1 * 1.379173542857, scale1 * 0.252275590895543), 0, 0, 0);
            poly22.AddVertexAt(11, new Point2d(scale1 * 1.38595311405261, scale1 * 0.253818527788335), 0, 0, 0);
            poly22.AddVertexAt(12, new Point2d(scale1 * 1.39273268524821, scale1 * 0.252275590895543), 0, 0, 0);
            poly22.Closed = true;
            poly22.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly22.Layer = "0";
            poly22.Color = color_GP;
            poly22.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly22);
            Polyline poly23 = new Polyline();
            poly23.AddVertexAt(0, new Point2d(scale1 * 1.36210772570945, scale1 * 0.243438770509551), 0, 0, 0);
            poly23.AddVertexAt(1, new Point2d(scale1 * 1.35930238590438, scale1 * 0.241895833616759), 0, 0, 0);
            poly23.AddVertexAt(2, new Point2d(scale1 * 1.36023749917273, scale1 * 0.240072362743458), 0, 0, 0);
            poly23.AddVertexAt(3, new Point2d(scale1 * 1.36164016907527, scale1 * 0.237968357889651), 0, 0, 0);
            poly23.AddVertexAt(4, new Point2d(scale1 * 1.36467928719744, scale1 * 0.237687823909144), 0, 0, 0);
            poly23.AddVertexAt(5, new Point2d(scale1 * 1.36818596195379, scale1 * 0.23937102779219), 0, 0, 0);
            poly23.AddVertexAt(6, new Point2d(scale1 * 1.36818596195379, scale1 * 0.241615299636252), 0, 0, 0);
            poly23.AddVertexAt(7, new Point2d(scale1 * 1.36584817878289, scale1 * 0.244140105460821), 0, 0, 0);
            poly23.AddVertexAt(8, new Point2d(scale1 * 1.36210772570945, scale1 * 0.243438770509551), 0, 0, 0);
            poly23.Closed = true;
            poly23.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly23.Layer = "0";
            poly23.Color = color_GP;
            poly23.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly23);
            Polyline poly24 = new Polyline();
            poly24.AddVertexAt(0, new Point2d(scale1 * 1.31652095387695, scale1 * 0.248628649148944), 0, 0, 0);
            poly24.AddVertexAt(1, new Point2d(scale1 * 1.32306674675547, scale1 * 0.250452120022245), 0, 0, 0);
            poly24.AddVertexAt(2, new Point2d(scale1 * 1.3359245541954, scale1 * 0.249750785070974), 0, 0, 0);
            poly24.AddVertexAt(3, new Point2d(scale1 * 1.34247034707391, scale1 * 0.248628649148944), 0, 0, 0);
            poly24.AddVertexAt(4, new Point2d(scale1 * 1.34363923865936, scale1 * 0.246805178275643), 0, 0, 0);
            poly24.AddVertexAt(5, new Point2d(scale1 * 1.34387301697645, scale1 * 0.244981707402343), 0, 0, 0);
            poly24.AddVertexAt(6, new Point2d(scale1 * 1.34410679529354, scale1 * 0.242737435558282), 0, 0, 0);
            poly24.AddVertexAt(7, new Point2d(scale1 * 1.34597702183026, scale1 * 0.240773697694727), 0, 0, 0);
            poly24.AddVertexAt(8, new Point2d(scale1 * 1.34457435192772, scale1 * 0.237407289928637), 0, 0, 0);
            poly24.AddVertexAt(9, new Point2d(scale1 * 1.33943122895175, scale1 * 0.235583819055336), 0, 0, 0);
            poly24.AddVertexAt(10, new Point2d(scale1 * 1.32937876131689, scale1 * 0.235303285074829), 0, 0, 0);
            poly24.AddVertexAt(11, new Point2d(scale1 * 1.324469416658, scale1 * 0.236144887016352), 0, 0, 0);
            poly24.AddVertexAt(12, new Point2d(scale1 * 1.32096274190166, scale1 * 0.237828090899398), 0, 0, 0);
            poly24.AddVertexAt(13, new Point2d(scale1 * 1.3176898454624, scale1 * 0.241054231675236), 0, 0, 0);
            poly24.AddVertexAt(14, new Point2d(scale1 * 1.31581961892568, scale1 * 0.244420639441328), 0, 0, 0);
            poly24.Closed = true;
            poly24.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly24.Layer = "0";
            poly24.Color = color_GP;
            poly24.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly24);
            Polyline poly25 = new Polyline();
            poly25.AddVertexAt(0, new Point2d(scale1 * 1.26028574100171, scale1 * 0.3), 0, 0, 0);
            poly25.AddVertexAt(1, new Point2d(scale1 * 1.26164016907527, scale1 * 0.297968357889651), 0, 0, 0);
            poly25.AddVertexAt(2, new Point2d(scale1 * 1.26467928719744, scale1 * 0.297687823909144), 0, 0, 0);
            poly25.AddVertexAt(3, new Point2d(scale1 * 1.26818596195379, scale1 * 0.29937102779219), 0, 0, 0);
            poly25.AddVertexAt(4, new Point2d(scale1 * 1.26818596195379, scale1 * 0.3), 0, 0, 0);
            poly25.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly25.Layer = "0";
            poly25.Color = color_GP;
            poly25.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly25);
            Polyline poly26 = new Polyline();
            poly26.AddVertexAt(0, new Point2d(scale1 * 1.24565464779079, scale1 * 0.3), 0, 0, 0);
            poly26.AddVertexAt(1, new Point2d(scale1 * 1.24457435192772, scale1 * 0.297407289928637), 0, 0, 0);
            poly26.AddVertexAt(2, new Point2d(scale1 * 1.23943122895175, scale1 * 0.295583819055336), 0, 0, 0);
            poly26.AddVertexAt(3, new Point2d(scale1 * 1.22937876131689, scale1 * 0.295303285074829), 0, 0, 0);
            poly26.AddVertexAt(4, new Point2d(scale1 * 1.224469416658, scale1 * 0.296144887016352), 0, 0, 0);
            poly26.AddVertexAt(5, new Point2d(scale1 * 1.22096274190166, scale1 * 0.297828090899398), 0, 0, 0);
            poly26.AddVertexAt(6, new Point2d(scale1 * 1.21875935585757, scale1 * 0.3), 0, 0, 0);
            poly26.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly26.Layer = "0";
            poly26.Color = color_GP;
            poly26.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly26);
            Polyline poly27 = new Polyline();
            poly27.AddVertexAt(0, new Point2d(scale1 * 1.25836727263602, scale1 * 0.282994061761134), 0, 0, 0);
            poly27.AddVertexAt(1, new Point2d(scale1 * 1.2564970460993, scale1 * 0.282853794770882), 0, 0, 0);
            poly27.AddVertexAt(2, new Point2d(scale1 * 1.2541592629284, scale1 * 0.28355512972215), 0, 0, 0);
            poly27.AddVertexAt(3, new Point2d(scale1 * 1.25228903639168, scale1 * 0.285518867585705), 0, 0, 0);
            poly27.AddVertexAt(4, new Point2d(scale1 * 1.25392548461131, scale1 * 0.287061804478495), 0, 0, 0);
            poly27.AddVertexAt(5, new Point2d(scale1 * 1.25836727263602, scale1 * 0.286220202536974), 0, 0, 0);
            poly27.AddVertexAt(6, new Point2d(scale1 * 1.25953616422147, scale1 * 0.284817532634435), 0, 0, 0);
            poly27.Closed = true;
            poly27.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly27.Layer = "0";
            poly27.Color = color_GP;
            poly27.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly27);
            Polyline poly28 = new Polyline();
            poly28.AddVertexAt(0, new Point2d(scale1 * 1.2679521836367, scale1 * 0.28355512972215), 0, 0, 0);
            poly28.AddVertexAt(1, new Point2d(scale1 * 1.26608195709998, scale1 * 0.282012192829359), 0, 0, 0);
            poly28.AddVertexAt(2, new Point2d(scale1 * 1.26584817878289, scale1 * 0.280048454965803), 0, 0, 0);
            poly28.AddVertexAt(3, new Point2d(scale1 * 1.26841974027088, scale1 * 0.278505518073012), 0, 0, 0);
            poly28.AddVertexAt(4, new Point2d(scale1 * 1.27169263671013, scale1 * 0.278926319043773), 0, 0, 0);
            poly28.AddVertexAt(5, new Point2d(scale1 * 1.2723939716614, scale1 * 0.281030323897581), 0, 0, 0);
            poly28.AddVertexAt(6, new Point2d(scale1 * 1.27192641502722, scale1 * 0.282713527780627), 0, 0, 0);
            poly28.AddVertexAt(7, new Point2d(scale1 * 1.2679521836367, scale1 * 0.28355512972215), 0, 0, 0);
            poly28.Closed = true;
            poly28.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly28.Layer = "0";
            poly28.Color = color_GP;
            poly28.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly28);
            Polyline poly29 = new Polyline();
            poly29.AddVertexAt(0, new Point2d(scale1 * 1.21657711212705, scale1 * 0.287355090003573), 0, 0, 0);
            poly29.AddVertexAt(1, new Point2d(scale1 * 1.21423932895615, scale1 * 0.285251085149766), 0, 0, 0);
            poly29.AddVertexAt(2, new Point2d(scale1 * 1.21564199885869, scale1 * 0.283427614276465), 0, 0, 0);
            poly29.AddVertexAt(3, new Point2d(scale1 * 1.21961623024921, scale1 * 0.282726279325197), 0, 0, 0);
            poly29.AddVertexAt(4, new Point2d(scale1 * 1.22265534837138, scale1 * 0.283988682237481), 0, 0, 0);
            poly29.AddVertexAt(5, new Point2d(scale1 * 1.22265534837138, scale1 * 0.286653755052303), 0, 0, 0);
            poly29.AddVertexAt(6, new Point2d(scale1 * 1.21657711212705, scale1 * 0.287355090003573), 0, 0, 0);
            poly29.Closed = true;
            poly29.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly29.Layer = "0";
            poly29.Color = color_GP;
            poly29.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly29);
            Polyline poly30 = new Polyline();
            poly30.AddVertexAt(0, new Point2d(scale1 * 1.21108075824353, scale1 * 0.266735842436258), 0, 0, 0);
            poly30.AddVertexAt(1, new Point2d(scale1 * 1.20947696180731, scale1 * 0.263254670769046), 0, 0, 0);
            poly30.AddVertexAt(2, new Point2d(scale1 * 1.2097107401244, scale1 * 0.259607729022446), 0, 0, 0);
            poly30.AddVertexAt(3, new Point2d(scale1 * 1.21064585339276, scale1 * 0.256802389217368), 0, 0, 0);
            poly30.AddVertexAt(4, new Point2d(scale1 * 1.21485386310037, scale1 * 0.255960787275844), 0, 0, 0);
            poly30.AddVertexAt(5, new Point2d(scale1 * 1.21906187280799, scale1 * 0.256521855236861), 0, 0, 0);
            poly30.AddVertexAt(6, new Point2d(scale1 * 1.22256854756433, scale1 * 0.258625860090668), 0, 0, 0);
            poly30.AddVertexAt(7, new Point2d(scale1 * 1.21906187280799, scale1 * 0.263535204749552), 0, 0, 0);
            poly30.AddVertexAt(8, new Point2d(scale1 * 1.21108075824353, scale1 * 0.266735842436258), 0, 0, 0);
            poly30.Closed = true;
            poly30.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly30.Layer = "0";
            poly30.Color = color_GP;
            poly30.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly30);
            Polyline poly31 = new Polyline();
            poly31.AddVertexAt(0, new Point2d(scale1 * 1.28797544496318, scale1 * 0.290759752403364), 0, 0, 0);
            poly31.AddVertexAt(1, new Point2d(scale1 * 1.28470254852392, scale1 * 0.288936281530065), 0, 0, 0);
            poly31.AddVertexAt(2, new Point2d(scale1 * 1.28259854367012, scale1 * 0.283746402890672), 0, 0, 0);
            poly31.AddVertexAt(3, new Point2d(scale1 * 1.28259854367012, scale1 * 0.281221597066103), 0, 0, 0);
            poly31.AddVertexAt(4, new Point2d(scale1 * 1.28189720871885, scale1 * 0.275751184446203), 0, 0, 0);
            poly31.AddVertexAt(5, new Point2d(scale1 * 1.28540388347519, scale1 * 0.273366645611888), 0, 0, 0);
            poly31.AddVertexAt(6, new Point2d(scale1 * 1.29031322813408, scale1 * 0.272525043670365), 0, 0, 0);
            poly31.AddVertexAt(7, new Point2d(scale1 * 1.30176836567148, scale1 * 0.272945844641126), 0, 0, 0);
            poly31.AddVertexAt(8, new Point2d(scale1 * 1.30480748379364, scale1 * 0.274488781533918), 0, 0, 0);
            poly31.AddVertexAt(9, new Point2d(scale1 * 1.305742597062, scale1 * 0.278275990270772), 0, 0, 0);
            poly31.AddVertexAt(10, new Point2d(scale1 * 1.30270347893983, scale1 * 0.28037999512458), 0, 0, 0);
            poly31.AddVertexAt(11, new Point2d(scale1 * 1.29989813913476, scale1 * 0.282764533958894), 0, 0, 0);
            poly31.AddVertexAt(12, new Point2d(scale1 * 1.29756035596386, scale1 * 0.286271208715241), 0, 0, 0);
            poly31.AddVertexAt(13, new Point2d(scale1 * 1.29569012942714, scale1 * 0.288796014539811), 0, 0, 0);
            poly31.AddVertexAt(14, new Point2d(scale1 * 1.29148211971953, scale1 * 0.289777883471588), 0, 0, 0);
            poly31.AddVertexAt(15, new Point2d(scale1 * 1.28797544496318, scale1 * 0.290759752403364), 0, 0, 0);
            poly31.Closed = true;
            poly31.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly31.Layer = "0";
            poly31.Color = color_GP;
            poly31.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly31);
            Polyline poly32 = new Polyline();
            poly32.AddVertexAt(0, new Point2d(scale1 * 1.25836727263602, scale1 * 0.274437775355651), 0, 0, 0);
            poly32.AddVertexAt(1, new Point2d(scale1 * 1.26608195709998, scale1 * 0.272754571472604), 0, 0, 0);
            poly32.AddVertexAt(2, new Point2d(scale1 * 1.26935485353923, scale1 * 0.268125760794228), 0, 0, 0);
            poly32.AddVertexAt(3, new Point2d(scale1 * 1.26608195709998, scale1 * 0.263637217106105), 0, 0, 0);
            poly32.AddVertexAt(4, new Point2d(scale1 * 1.25836727263602, scale1 * 0.261954013223059), 0, 0, 0);
            poly32.AddVertexAt(5, new Point2d(scale1 * 1.2497174749037, scale1 * 0.261392945262043), 0, 0, 0);
            poly32.AddVertexAt(6, new Point2d(scale1 * 1.2450419085619, scale1 * 0.26237481419382), 0, 0, 0);
            poly32.AddVertexAt(7, new Point2d(scale1 * 1.24130145548847, scale1 * 0.26489962001839), 0, 0, 0);
            poly32.AddVertexAt(8, new Point2d(scale1 * 1.23078143121943, scale1 * 0.268967362735751), 0, 0, 0);
            poly32.AddVertexAt(9, new Point2d(scale1 * 1.22517075160927, scale1 * 0.272894838462859), 0, 0, 0);
            poly32.AddVertexAt(10, new Point2d(scale1 * 1.22657342151181, scale1 * 0.27471830933616), 0, 0, 0);
            poly32.AddVertexAt(11, new Point2d(scale1 * 1.23498944092704, scale1 * 0.278365251082757), 0, 0, 0);
            poly32.AddVertexAt(12, new Point2d(scale1 * 1.24153523380556, scale1 * 0.279066586034028), 0, 0, 0);
            poly32.AddVertexAt(13, new Point2d(scale1 * 1.24574324351317, scale1 * 0.277804183121743), 0, 0, 0);
            poly32.AddVertexAt(14, new Point2d(scale1 * 1.25135392312332, scale1 * 0.276541780209458), 0, 0, 0);
            poly32.AddVertexAt(15, new Point2d(scale1 * 1.25579571114803, scale1 * 0.275840445258188), 0, 0, 0);
            poly32.Closed = true;
            poly32.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly32.Layer = "0";
            poly32.Color = color_GP;
            poly32.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly32);
            Polyline poly33 = new Polyline();
            poly33.AddVertexAt(0, new Point2d(scale1 * 1.27889889791182, scale1 * 0.261890255500208), 0, 0, 0);
            poly33.AddVertexAt(1, new Point2d(scale1 * 1.27274102660718, scale1 * 0.262017770945895), 0, 0, 0);
            poly33.AddVertexAt(2, new Point2d(scale1 * 1.26743251686181, scale1 * 0.259722492923566), 0, 0, 0);
            poly33.AddVertexAt(3, new Point2d(scale1 * 1.26637081491274, scale1 * 0.256662122227115), 0, 0, 0);
            poly33.AddVertexAt(4, new Point2d(scale1 * 1.27167932465811, scale1 * 0.255769514107321), 0, 0, 0);
            poly33.AddVertexAt(5, new Point2d(scale1 * 1.27826187674237, scale1 * 0.2567896376728), 0, 0, 0);
            poly33.AddVertexAt(6, new Point2d(scale1 * 1.28208400375904, scale1 * 0.259467462032192), 0, 0, 0);
            poly33.AddVertexAt(7, new Point2d(scale1 * 1.27889889791182, scale1 * 0.261890255500208), 0, 0, 0);
            poly33.Closed = true;
            poly33.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly33.Layer = "0";
            poly33.Color = color_GP;
            poly33.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly33);
            Polyline poly34 = new Polyline();
            poly34.AddVertexAt(0, new Point2d(scale1 * 1.29647313832164, scale1 * 0.248628649148944), 0, 0, 0);
            poly34.AddVertexAt(1, new Point2d(scale1 * 1.29670691663873, scale1 * 0.243298503519296), 0, 0, 0);
            poly34.AddVertexAt(2, new Point2d(scale1 * 1.29507046841911, scale1 * 0.242036100607014), 0, 0, 0);
            poly34.AddVertexAt(3, new Point2d(scale1 * 1.29086245871149, scale1 * 0.240633430704475), 0, 0, 0);
            poly34.AddVertexAt(4, new Point2d(scale1 * 1.28688822732096, scale1 * 0.240773697694727), 0, 0, 0);
            poly34.AddVertexAt(5, new Point2d(scale1 * 1.2838491091988, scale1 * 0.242036100607014), 0, 0, 0);
            poly34.AddVertexAt(6, new Point2d(scale1 * 1.28057621275954, scale1 * 0.243298503519296), 0, 0, 0);
            poly34.AddVertexAt(7, new Point2d(scale1 * 1.27753709463737, scale1 * 0.244560906431583), 0, 0, 0);
            poly34.AddVertexAt(8, new Point2d(scale1 * 1.27566686810066, scale1 * 0.246945445265898), 0, 0, 0);
            poly34.AddVertexAt(9, new Point2d(scale1 * 1.27566686810066, scale1 * 0.248628649148944), 0, 0, 0);
            poly34.AddVertexAt(10, new Point2d(scale1 * 1.279173542857, scale1 * 0.252275590895543), 0, 0, 0);
            poly34.AddVertexAt(11, new Point2d(scale1 * 1.28595311405261, scale1 * 0.253818527788335), 0, 0, 0);
            poly34.AddVertexAt(12, new Point2d(scale1 * 1.29273268524821, scale1 * 0.252275590895543), 0, 0, 0);
            poly34.Closed = true;
            poly34.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly34.Layer = "0";
            poly34.Color = color_GP;
            poly34.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly34);
            Polyline poly35 = new Polyline();
            poly35.AddVertexAt(0, new Point2d(scale1 * 1.26210772570945, scale1 * 0.243438770509551), 0, 0, 0);
            poly35.AddVertexAt(1, new Point2d(scale1 * 1.25930238590438, scale1 * 0.241895833616759), 0, 0, 0);
            poly35.AddVertexAt(2, new Point2d(scale1 * 1.26023749917273, scale1 * 0.240072362743458), 0, 0, 0);
            poly35.AddVertexAt(3, new Point2d(scale1 * 1.26164016907527, scale1 * 0.237968357889651), 0, 0, 0);
            poly35.AddVertexAt(4, new Point2d(scale1 * 1.26467928719744, scale1 * 0.237687823909144), 0, 0, 0);
            poly35.AddVertexAt(5, new Point2d(scale1 * 1.26818596195379, scale1 * 0.23937102779219), 0, 0, 0);
            poly35.AddVertexAt(6, new Point2d(scale1 * 1.26818596195379, scale1 * 0.241615299636252), 0, 0, 0);
            poly35.AddVertexAt(7, new Point2d(scale1 * 1.26584817878289, scale1 * 0.244140105460821), 0, 0, 0);
            poly35.AddVertexAt(8, new Point2d(scale1 * 1.26210772570945, scale1 * 0.243438770509551), 0, 0, 0);
            poly35.Closed = true;
            poly35.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly35.Layer = "0";
            poly35.Color = color_GP;
            poly35.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly35);
            Polyline poly36 = new Polyline();
            poly36.AddVertexAt(0, new Point2d(scale1 * 1.21652095387695, scale1 * 0.248628649148944), 0, 0, 0);
            poly36.AddVertexAt(1, new Point2d(scale1 * 1.22306674675547, scale1 * 0.250452120022245), 0, 0, 0);
            poly36.AddVertexAt(2, new Point2d(scale1 * 1.2359245541954, scale1 * 0.249750785070974), 0, 0, 0);
            poly36.AddVertexAt(3, new Point2d(scale1 * 1.24247034707391, scale1 * 0.248628649148944), 0, 0, 0);
            poly36.AddVertexAt(4, new Point2d(scale1 * 1.24363923865936, scale1 * 0.246805178275643), 0, 0, 0);
            poly36.AddVertexAt(5, new Point2d(scale1 * 1.24387301697645, scale1 * 0.244981707402343), 0, 0, 0);
            poly36.AddVertexAt(6, new Point2d(scale1 * 1.24410679529354, scale1 * 0.242737435558282), 0, 0, 0);
            poly36.AddVertexAt(7, new Point2d(scale1 * 1.24597702183026, scale1 * 0.240773697694727), 0, 0, 0);
            poly36.AddVertexAt(8, new Point2d(scale1 * 1.24457435192772, scale1 * 0.237407289928637), 0, 0, 0);
            poly36.AddVertexAt(9, new Point2d(scale1 * 1.23943122895175, scale1 * 0.235583819055336), 0, 0, 0);
            poly36.AddVertexAt(10, new Point2d(scale1 * 1.22937876131689, scale1 * 0.235303285074829), 0, 0, 0);
            poly36.AddVertexAt(11, new Point2d(scale1 * 1.224469416658, scale1 * 0.236144887016352), 0, 0, 0);
            poly36.AddVertexAt(12, new Point2d(scale1 * 1.22096274190166, scale1 * 0.237828090899398), 0, 0, 0);
            poly36.AddVertexAt(13, new Point2d(scale1 * 1.2176898454624, scale1 * 0.241054231675236), 0, 0, 0);
            poly36.AddVertexAt(14, new Point2d(scale1 * 1.21581961892568, scale1 * 0.244420639441328), 0, 0, 0);
            poly36.Closed = true;
            poly36.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly36.Layer = "0";
            poly36.Color = color_GP;
            poly36.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly36);
            Polyline poly37 = new Polyline();
            poly37.AddVertexAt(0, new Point2d(scale1 * 1.16028574100171, scale1 * 0.3), 0, 0, 0);
            poly37.AddVertexAt(1, new Point2d(scale1 * 1.16164016907527, scale1 * 0.297968357889651), 0, 0, 0);
            poly37.AddVertexAt(2, new Point2d(scale1 * 1.16467928719744, scale1 * 0.297687823909144), 0, 0, 0);
            poly37.AddVertexAt(3, new Point2d(scale1 * 1.16818596195379, scale1 * 0.29937102779219), 0, 0, 0);
            poly37.AddVertexAt(4, new Point2d(scale1 * 1.16818596195379, scale1 * 0.3), 0, 0, 0);
            poly37.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly37.Layer = "0";
            poly37.Color = color_GP;
            poly37.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly37);
            Polyline poly38 = new Polyline();
            poly38.AddVertexAt(0, new Point2d(scale1 * 1.14565464779079, scale1 * 0.3), 0, 0, 0);
            poly38.AddVertexAt(1, new Point2d(scale1 * 1.14457435192772, scale1 * 0.297407289928637), 0, 0, 0);
            poly38.AddVertexAt(2, new Point2d(scale1 * 1.13943122895175, scale1 * 0.295583819055336), 0, 0, 0);
            poly38.AddVertexAt(3, new Point2d(scale1 * 1.12937876131689, scale1 * 0.295303285074829), 0, 0, 0);
            poly38.AddVertexAt(4, new Point2d(scale1 * 1.124469416658, scale1 * 0.296144887016352), 0, 0, 0);
            poly38.AddVertexAt(5, new Point2d(scale1 * 1.12096274190166, scale1 * 0.297828090899398), 0, 0, 0);
            poly38.AddVertexAt(6, new Point2d(scale1 * 1.11875935585757, scale1 * 0.3), 0, 0, 0);
            poly38.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly38.Layer = "0";
            poly38.Color = color_GP;
            poly38.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly38);
            Polyline poly39 = new Polyline();
            poly39.AddVertexAt(0, new Point2d(scale1 * 1.15836727263602, scale1 * 0.282994061761134), 0, 0, 0);
            poly39.AddVertexAt(1, new Point2d(scale1 * 1.1564970460993, scale1 * 0.282853794770882), 0, 0, 0);
            poly39.AddVertexAt(2, new Point2d(scale1 * 1.1541592629284, scale1 * 0.28355512972215), 0, 0, 0);
            poly39.AddVertexAt(3, new Point2d(scale1 * 1.15228903639168, scale1 * 0.285518867585705), 0, 0, 0);
            poly39.AddVertexAt(4, new Point2d(scale1 * 1.15392548461131, scale1 * 0.287061804478495), 0, 0, 0);
            poly39.AddVertexAt(5, new Point2d(scale1 * 1.15836727263602, scale1 * 0.286220202536974), 0, 0, 0);
            poly39.AddVertexAt(6, new Point2d(scale1 * 1.15953616422147, scale1 * 0.284817532634435), 0, 0, 0);
            poly39.Closed = true;
            poly39.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly39.Layer = "0";
            poly39.Color = color_GP;
            poly39.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly39);
            Polyline poly40 = new Polyline();
            poly40.AddVertexAt(0, new Point2d(scale1 * 1.1679521836367, scale1 * 0.28355512972215), 0, 0, 0);
            poly40.AddVertexAt(1, new Point2d(scale1 * 1.16608195709998, scale1 * 0.282012192829359), 0, 0, 0);
            poly40.AddVertexAt(2, new Point2d(scale1 * 1.16584817878289, scale1 * 0.280048454965803), 0, 0, 0);
            poly40.AddVertexAt(3, new Point2d(scale1 * 1.16841974027088, scale1 * 0.278505518073012), 0, 0, 0);
            poly40.AddVertexAt(4, new Point2d(scale1 * 1.17169263671013, scale1 * 0.278926319043773), 0, 0, 0);
            poly40.AddVertexAt(5, new Point2d(scale1 * 1.1723939716614, scale1 * 0.281030323897581), 0, 0, 0);
            poly40.AddVertexAt(6, new Point2d(scale1 * 1.17192641502722, scale1 * 0.282713527780627), 0, 0, 0);
            poly40.AddVertexAt(7, new Point2d(scale1 * 1.1679521836367, scale1 * 0.28355512972215), 0, 0, 0);
            poly40.Closed = true;
            poly40.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly40.Layer = "0";
            poly40.Color = color_GP;
            poly40.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly40);
            Polyline poly41 = new Polyline();
            poly41.AddVertexAt(0, new Point2d(scale1 * 1.11657711212705, scale1 * 0.287355090003573), 0, 0, 0);
            poly41.AddVertexAt(1, new Point2d(scale1 * 1.11423932895615, scale1 * 0.285251085149766), 0, 0, 0);
            poly41.AddVertexAt(2, new Point2d(scale1 * 1.11564199885869, scale1 * 0.283427614276465), 0, 0, 0);
            poly41.AddVertexAt(3, new Point2d(scale1 * 1.11961623024921, scale1 * 0.282726279325197), 0, 0, 0);
            poly41.AddVertexAt(4, new Point2d(scale1 * 1.12265534837138, scale1 * 0.283988682237481), 0, 0, 0);
            poly41.AddVertexAt(5, new Point2d(scale1 * 1.12265534837138, scale1 * 0.286653755052303), 0, 0, 0);
            poly41.AddVertexAt(6, new Point2d(scale1 * 1.11657711212705, scale1 * 0.287355090003573), 0, 0, 0);
            poly41.Closed = true;
            poly41.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly41.Layer = "0";
            poly41.Color = color_GP;
            poly41.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly41);
            Polyline poly42 = new Polyline();
            poly42.AddVertexAt(0, new Point2d(scale1 * 1.11108075824353, scale1 * 0.266735842436258), 0, 0, 0);
            poly42.AddVertexAt(1, new Point2d(scale1 * 1.10947696180731, scale1 * 0.263254670769046), 0, 0, 0);
            poly42.AddVertexAt(2, new Point2d(scale1 * 1.1097107401244, scale1 * 0.259607729022446), 0, 0, 0);
            poly42.AddVertexAt(3, new Point2d(scale1 * 1.11064585339276, scale1 * 0.256802389217368), 0, 0, 0);
            poly42.AddVertexAt(4, new Point2d(scale1 * 1.11485386310037, scale1 * 0.255960787275844), 0, 0, 0);
            poly42.AddVertexAt(5, new Point2d(scale1 * 1.11906187280799, scale1 * 0.256521855236861), 0, 0, 0);
            poly42.AddVertexAt(6, new Point2d(scale1 * 1.12256854756433, scale1 * 0.258625860090668), 0, 0, 0);
            poly42.AddVertexAt(7, new Point2d(scale1 * 1.11906187280799, scale1 * 0.263535204749552), 0, 0, 0);
            poly42.AddVertexAt(8, new Point2d(scale1 * 1.11108075824353, scale1 * 0.266735842436258), 0, 0, 0);
            poly42.Closed = true;
            poly42.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly42.Layer = "0";
            poly42.Color = color_GP;
            poly42.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly42);
            Polyline poly43 = new Polyline();
            poly43.AddVertexAt(0, new Point2d(scale1 * 1.18797544496318, scale1 * 0.290759752403364), 0, 0, 0);
            poly43.AddVertexAt(1, new Point2d(scale1 * 1.18470254852392, scale1 * 0.288936281530065), 0, 0, 0);
            poly43.AddVertexAt(2, new Point2d(scale1 * 1.18259854367012, scale1 * 0.283746402890672), 0, 0, 0);
            poly43.AddVertexAt(3, new Point2d(scale1 * 1.18259854367012, scale1 * 0.281221597066103), 0, 0, 0);
            poly43.AddVertexAt(4, new Point2d(scale1 * 1.18189720871885, scale1 * 0.275751184446203), 0, 0, 0);
            poly43.AddVertexAt(5, new Point2d(scale1 * 1.18540388347519, scale1 * 0.273366645611888), 0, 0, 0);
            poly43.AddVertexAt(6, new Point2d(scale1 * 1.19031322813408, scale1 * 0.272525043670365), 0, 0, 0);
            poly43.AddVertexAt(7, new Point2d(scale1 * 1.20176836567148, scale1 * 0.272945844641126), 0, 0, 0);
            poly43.AddVertexAt(8, new Point2d(scale1 * 1.20480748379364, scale1 * 0.274488781533918), 0, 0, 0);
            poly43.AddVertexAt(9, new Point2d(scale1 * 1.205742597062, scale1 * 0.278275990270772), 0, 0, 0);
            poly43.AddVertexAt(10, new Point2d(scale1 * 1.20270347893983, scale1 * 0.28037999512458), 0, 0, 0);
            poly43.AddVertexAt(11, new Point2d(scale1 * 1.19989813913476, scale1 * 0.282764533958894), 0, 0, 0);
            poly43.AddVertexAt(12, new Point2d(scale1 * 1.19756035596386, scale1 * 0.286271208715241), 0, 0, 0);
            poly43.AddVertexAt(13, new Point2d(scale1 * 1.19569012942714, scale1 * 0.288796014539811), 0, 0, 0);
            poly43.AddVertexAt(14, new Point2d(scale1 * 1.19148211971953, scale1 * 0.289777883471588), 0, 0, 0);
            poly43.AddVertexAt(15, new Point2d(scale1 * 1.18797544496318, scale1 * 0.290759752403364), 0, 0, 0);
            poly43.Closed = true;
            poly43.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly43.Layer = "0";
            poly43.Color = color_GP;
            poly43.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly43);
            Polyline poly44 = new Polyline();
            poly44.AddVertexAt(0, new Point2d(scale1 * 1.15836727263602, scale1 * 0.274437775355651), 0, 0, 0);
            poly44.AddVertexAt(1, new Point2d(scale1 * 1.16608195709998, scale1 * 0.272754571472604), 0, 0, 0);
            poly44.AddVertexAt(2, new Point2d(scale1 * 1.16935485353923, scale1 * 0.268125760794228), 0, 0, 0);
            poly44.AddVertexAt(3, new Point2d(scale1 * 1.16608195709998, scale1 * 0.263637217106105), 0, 0, 0);
            poly44.AddVertexAt(4, new Point2d(scale1 * 1.15836727263602, scale1 * 0.261954013223059), 0, 0, 0);
            poly44.AddVertexAt(5, new Point2d(scale1 * 1.1497174749037, scale1 * 0.261392945262043), 0, 0, 0);
            poly44.AddVertexAt(6, new Point2d(scale1 * 1.1450419085619, scale1 * 0.26237481419382), 0, 0, 0);
            poly44.AddVertexAt(7, new Point2d(scale1 * 1.14130145548847, scale1 * 0.26489962001839), 0, 0, 0);
            poly44.AddVertexAt(8, new Point2d(scale1 * 1.13078143121943, scale1 * 0.268967362735751), 0, 0, 0);
            poly44.AddVertexAt(9, new Point2d(scale1 * 1.12517075160927, scale1 * 0.272894838462859), 0, 0, 0);
            poly44.AddVertexAt(10, new Point2d(scale1 * 1.12657342151181, scale1 * 0.27471830933616), 0, 0, 0);
            poly44.AddVertexAt(11, new Point2d(scale1 * 1.13498944092704, scale1 * 0.278365251082757), 0, 0, 0);
            poly44.AddVertexAt(12, new Point2d(scale1 * 1.14153523380556, scale1 * 0.279066586034028), 0, 0, 0);
            poly44.AddVertexAt(13, new Point2d(scale1 * 1.14574324351317, scale1 * 0.277804183121743), 0, 0, 0);
            poly44.AddVertexAt(14, new Point2d(scale1 * 1.15135392312332, scale1 * 0.276541780209458), 0, 0, 0);
            poly44.AddVertexAt(15, new Point2d(scale1 * 1.15579571114803, scale1 * 0.275840445258188), 0, 0, 0);
            poly44.Closed = true;
            poly44.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly44.Layer = "0";
            poly44.Color = color_GP;
            poly44.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly44);
            Polyline poly45 = new Polyline();
            poly45.AddVertexAt(0, new Point2d(scale1 * 1.17889889791182, scale1 * 0.261890255500208), 0, 0, 0);
            poly45.AddVertexAt(1, new Point2d(scale1 * 1.17274102660718, scale1 * 0.262017770945895), 0, 0, 0);
            poly45.AddVertexAt(2, new Point2d(scale1 * 1.16743251686181, scale1 * 0.259722492923566), 0, 0, 0);
            poly45.AddVertexAt(3, new Point2d(scale1 * 1.16637081491274, scale1 * 0.256662122227115), 0, 0, 0);
            poly45.AddVertexAt(4, new Point2d(scale1 * 1.17167932465811, scale1 * 0.255769514107321), 0, 0, 0);
            poly45.AddVertexAt(5, new Point2d(scale1 * 1.17826187674237, scale1 * 0.2567896376728), 0, 0, 0);
            poly45.AddVertexAt(6, new Point2d(scale1 * 1.18208400375904, scale1 * 0.259467462032192), 0, 0, 0);
            poly45.AddVertexAt(7, new Point2d(scale1 * 1.17889889791182, scale1 * 0.261890255500208), 0, 0, 0);
            poly45.Closed = true;
            poly45.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly45.Layer = "0";
            poly45.Color = color_GP;
            poly45.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly45);
            Polyline poly46 = new Polyline();
            poly46.AddVertexAt(0, new Point2d(scale1 * 1.19647313832164, scale1 * 0.248628649148944), 0, 0, 0);
            poly46.AddVertexAt(1, new Point2d(scale1 * 1.19670691663873, scale1 * 0.243298503519296), 0, 0, 0);
            poly46.AddVertexAt(2, new Point2d(scale1 * 1.19507046841911, scale1 * 0.242036100607014), 0, 0, 0);
            poly46.AddVertexAt(3, new Point2d(scale1 * 1.19086245871149, scale1 * 0.240633430704475), 0, 0, 0);
            poly46.AddVertexAt(4, new Point2d(scale1 * 1.18688822732096, scale1 * 0.240773697694727), 0, 0, 0);
            poly46.AddVertexAt(5, new Point2d(scale1 * 1.1838491091988, scale1 * 0.242036100607014), 0, 0, 0);
            poly46.AddVertexAt(6, new Point2d(scale1 * 1.18057621275954, scale1 * 0.243298503519296), 0, 0, 0);
            poly46.AddVertexAt(7, new Point2d(scale1 * 1.17753709463737, scale1 * 0.244560906431583), 0, 0, 0);
            poly46.AddVertexAt(8, new Point2d(scale1 * 1.17566686810066, scale1 * 0.246945445265898), 0, 0, 0);
            poly46.AddVertexAt(9, new Point2d(scale1 * 1.17566686810066, scale1 * 0.248628649148944), 0, 0, 0);
            poly46.AddVertexAt(10, new Point2d(scale1 * 1.179173542857, scale1 * 0.252275590895543), 0, 0, 0);
            poly46.AddVertexAt(11, new Point2d(scale1 * 1.18595311405261, scale1 * 0.253818527788335), 0, 0, 0);
            poly46.AddVertexAt(12, new Point2d(scale1 * 1.19273268524821, scale1 * 0.252275590895543), 0, 0, 0);
            poly46.Closed = true;
            poly46.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly46.Layer = "0";
            poly46.Color = color_GP;
            poly46.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly46);
            Polyline poly47 = new Polyline();
            poly47.AddVertexAt(0, new Point2d(scale1 * 1.16210772570945, scale1 * 0.243438770509551), 0, 0, 0);
            poly47.AddVertexAt(1, new Point2d(scale1 * 1.15930238590438, scale1 * 0.241895833616759), 0, 0, 0);
            poly47.AddVertexAt(2, new Point2d(scale1 * 1.16023749917273, scale1 * 0.240072362743458), 0, 0, 0);
            poly47.AddVertexAt(3, new Point2d(scale1 * 1.16164016907527, scale1 * 0.237968357889651), 0, 0, 0);
            poly47.AddVertexAt(4, new Point2d(scale1 * 1.16467928719744, scale1 * 0.237687823909144), 0, 0, 0);
            poly47.AddVertexAt(5, new Point2d(scale1 * 1.16818596195379, scale1 * 0.23937102779219), 0, 0, 0);
            poly47.AddVertexAt(6, new Point2d(scale1 * 1.16818596195379, scale1 * 0.241615299636252), 0, 0, 0);
            poly47.AddVertexAt(7, new Point2d(scale1 * 1.16584817878289, scale1 * 0.244140105460821), 0, 0, 0);
            poly47.AddVertexAt(8, new Point2d(scale1 * 1.16210772570945, scale1 * 0.243438770509551), 0, 0, 0);
            poly47.Closed = true;
            poly47.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly47.Layer = "0";
            poly47.Color = color_GP;
            poly47.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly47);
            Polyline poly48 = new Polyline();
            poly48.AddVertexAt(0, new Point2d(scale1 * 1.11652095387695, scale1 * 0.248628649148944), 0, 0, 0);
            poly48.AddVertexAt(1, new Point2d(scale1 * 1.12306674675547, scale1 * 0.250452120022245), 0, 0, 0);
            poly48.AddVertexAt(2, new Point2d(scale1 * 1.1359245541954, scale1 * 0.249750785070974), 0, 0, 0);
            poly48.AddVertexAt(3, new Point2d(scale1 * 1.14247034707391, scale1 * 0.248628649148944), 0, 0, 0);
            poly48.AddVertexAt(4, new Point2d(scale1 * 1.14363923865936, scale1 * 0.246805178275643), 0, 0, 0);
            poly48.AddVertexAt(5, new Point2d(scale1 * 1.14387301697645, scale1 * 0.244981707402343), 0, 0, 0);
            poly48.AddVertexAt(6, new Point2d(scale1 * 1.14410679529354, scale1 * 0.242737435558282), 0, 0, 0);
            poly48.AddVertexAt(7, new Point2d(scale1 * 1.14597702183026, scale1 * 0.240773697694727), 0, 0, 0);
            poly48.AddVertexAt(8, new Point2d(scale1 * 1.14457435192772, scale1 * 0.237407289928637), 0, 0, 0);
            poly48.AddVertexAt(9, new Point2d(scale1 * 1.13943122895175, scale1 * 0.235583819055336), 0, 0, 0);
            poly48.AddVertexAt(10, new Point2d(scale1 * 1.12937876131689, scale1 * 0.235303285074829), 0, 0, 0);
            poly48.AddVertexAt(11, new Point2d(scale1 * 1.124469416658, scale1 * 0.236144887016352), 0, 0, 0);
            poly48.AddVertexAt(12, new Point2d(scale1 * 1.12096274190166, scale1 * 0.237828090899398), 0, 0, 0);
            poly48.AddVertexAt(13, new Point2d(scale1 * 1.1176898454624, scale1 * 0.241054231675236), 0, 0, 0);
            poly48.AddVertexAt(14, new Point2d(scale1 * 1.11581961892568, scale1 * 0.244420639441328), 0, 0, 0);
            poly48.Closed = true;
            poly48.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly48.Layer = "0";
            poly48.Color = color_GP;
            poly48.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly48);
            Polyline poly49 = new Polyline();
            poly49.AddVertexAt(0, new Point2d(scale1 * 1.06028574100171, scale1 * 0.3), 0, 0, 0);
            poly49.AddVertexAt(1, new Point2d(scale1 * 1.06164016907527, scale1 * 0.297968357889651), 0, 0, 0);
            poly49.AddVertexAt(2, new Point2d(scale1 * 1.06467928719744, scale1 * 0.297687823909144), 0, 0, 0);
            poly49.AddVertexAt(3, new Point2d(scale1 * 1.06818596195379, scale1 * 0.29937102779219), 0, 0, 0);
            poly49.AddVertexAt(4, new Point2d(scale1 * 1.06818596195379, scale1 * 0.3), 0, 0, 0);
            poly49.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly49.Layer = "0";
            poly49.Color = color_GP;
            poly49.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly49);
            Polyline poly50 = new Polyline();
            poly50.AddVertexAt(0, new Point2d(scale1 * 1.04565464779079, scale1 * 0.3), 0, 0, 0);
            poly50.AddVertexAt(1, new Point2d(scale1 * 1.04457435192772, scale1 * 0.297407289928637), 0, 0, 0);
            poly50.AddVertexAt(2, new Point2d(scale1 * 1.03943122895175, scale1 * 0.295583819055336), 0, 0, 0);
            poly50.AddVertexAt(3, new Point2d(scale1 * 1.02937876131689, scale1 * 0.295303285074829), 0, 0, 0);
            poly50.AddVertexAt(4, new Point2d(scale1 * 1.024469416658, scale1 * 0.296144887016352), 0, 0, 0);
            poly50.AddVertexAt(5, new Point2d(scale1 * 1.02096274190166, scale1 * 0.297828090899398), 0, 0, 0);
            poly50.AddVertexAt(6, new Point2d(scale1 * 1.01875935585757, scale1 * 0.3), 0, 0, 0);
            poly50.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly50.Layer = "0";
            poly50.Color = color_GP;
            poly50.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly50);
            Polyline poly51 = new Polyline();
            poly51.AddVertexAt(0, new Point2d(scale1 * 1.05836727263602, scale1 * 0.282994061761134), 0, 0, 0);
            poly51.AddVertexAt(1, new Point2d(scale1 * 1.0564970460993, scale1 * 0.282853794770882), 0, 0, 0);
            poly51.AddVertexAt(2, new Point2d(scale1 * 1.0541592629284, scale1 * 0.28355512972215), 0, 0, 0);
            poly51.AddVertexAt(3, new Point2d(scale1 * 1.05228903639168, scale1 * 0.285518867585705), 0, 0, 0);
            poly51.AddVertexAt(4, new Point2d(scale1 * 1.05392548461131, scale1 * 0.287061804478495), 0, 0, 0);
            poly51.AddVertexAt(5, new Point2d(scale1 * 1.05836727263602, scale1 * 0.286220202536974), 0, 0, 0);
            poly51.AddVertexAt(6, new Point2d(scale1 * 1.05953616422147, scale1 * 0.284817532634435), 0, 0, 0);
            poly51.Closed = true;
            poly51.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly51.Layer = "0";
            poly51.Color = color_GP;
            poly51.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly51);
            Polyline poly52 = new Polyline();
            poly52.AddVertexAt(0, new Point2d(scale1 * 1.0679521836367, scale1 * 0.28355512972215), 0, 0, 0);
            poly52.AddVertexAt(1, new Point2d(scale1 * 1.06608195709998, scale1 * 0.282012192829359), 0, 0, 0);
            poly52.AddVertexAt(2, new Point2d(scale1 * 1.06584817878289, scale1 * 0.280048454965803), 0, 0, 0);
            poly52.AddVertexAt(3, new Point2d(scale1 * 1.06841974027088, scale1 * 0.278505518073012), 0, 0, 0);
            poly52.AddVertexAt(4, new Point2d(scale1 * 1.07169263671013, scale1 * 0.278926319043773), 0, 0, 0);
            poly52.AddVertexAt(5, new Point2d(scale1 * 1.0723939716614, scale1 * 0.281030323897581), 0, 0, 0);
            poly52.AddVertexAt(6, new Point2d(scale1 * 1.07192641502722, scale1 * 0.282713527780627), 0, 0, 0);
            poly52.AddVertexAt(7, new Point2d(scale1 * 1.0679521836367, scale1 * 0.28355512972215), 0, 0, 0);
            poly52.Closed = true;
            poly52.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly52.Layer = "0";
            poly52.Color = color_GP;
            poly52.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly52);
            Polyline poly53 = new Polyline();
            poly53.AddVertexAt(0, new Point2d(scale1 * 1.01657711212705, scale1 * 0.287355090003573), 0, 0, 0);
            poly53.AddVertexAt(1, new Point2d(scale1 * 1.01423932895615, scale1 * 0.285251085149766), 0, 0, 0);
            poly53.AddVertexAt(2, new Point2d(scale1 * 1.01564199885869, scale1 * 0.283427614276465), 0, 0, 0);
            poly53.AddVertexAt(3, new Point2d(scale1 * 1.01961623024921, scale1 * 0.282726279325197), 0, 0, 0);
            poly53.AddVertexAt(4, new Point2d(scale1 * 1.02265534837138, scale1 * 0.283988682237481), 0, 0, 0);
            poly53.AddVertexAt(5, new Point2d(scale1 * 1.02265534837138, scale1 * 0.286653755052303), 0, 0, 0);
            poly53.AddVertexAt(6, new Point2d(scale1 * 1.01657711212705, scale1 * 0.287355090003573), 0, 0, 0);
            poly53.Closed = true;
            poly53.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly53.Layer = "0";
            poly53.Color = color_GP;
            poly53.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly53);
            Polyline poly54 = new Polyline();
            poly54.AddVertexAt(0, new Point2d(scale1 * 1.01108075824353, scale1 * 0.266735842436258), 0, 0, 0);
            poly54.AddVertexAt(1, new Point2d(scale1 * 1.00947696180731, scale1 * 0.263254670769046), 0, 0, 0);
            poly54.AddVertexAt(2, new Point2d(scale1 * 1.0097107401244, scale1 * 0.259607729022446), 0, 0, 0);
            poly54.AddVertexAt(3, new Point2d(scale1 * 1.01064585339276, scale1 * 0.256802389217368), 0, 0, 0);
            poly54.AddVertexAt(4, new Point2d(scale1 * 1.01485386310037, scale1 * 0.255960787275844), 0, 0, 0);
            poly54.AddVertexAt(5, new Point2d(scale1 * 1.01906187280799, scale1 * 0.256521855236861), 0, 0, 0);
            poly54.AddVertexAt(6, new Point2d(scale1 * 1.02256854756433, scale1 * 0.258625860090668), 0, 0, 0);
            poly54.AddVertexAt(7, new Point2d(scale1 * 1.01906187280799, scale1 * 0.263535204749552), 0, 0, 0);
            poly54.AddVertexAt(8, new Point2d(scale1 * 1.01108075824353, scale1 * 0.266735842436258), 0, 0, 0);
            poly54.Closed = true;
            poly54.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly54.Layer = "0";
            poly54.Color = color_GP;
            poly54.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly54);
            Polyline poly55 = new Polyline();
            poly55.AddVertexAt(0, new Point2d(scale1 * 1.08797544496318, scale1 * 0.290759752403364), 0, 0, 0);
            poly55.AddVertexAt(1, new Point2d(scale1 * 1.08470254852392, scale1 * 0.288936281530065), 0, 0, 0);
            poly55.AddVertexAt(2, new Point2d(scale1 * 1.08259854367012, scale1 * 0.283746402890672), 0, 0, 0);
            poly55.AddVertexAt(3, new Point2d(scale1 * 1.08259854367012, scale1 * 0.281221597066103), 0, 0, 0);
            poly55.AddVertexAt(4, new Point2d(scale1 * 1.08189720871885, scale1 * 0.275751184446203), 0, 0, 0);
            poly55.AddVertexAt(5, new Point2d(scale1 * 1.08540388347519, scale1 * 0.273366645611888), 0, 0, 0);
            poly55.AddVertexAt(6, new Point2d(scale1 * 1.09031322813408, scale1 * 0.272525043670365), 0, 0, 0);
            poly55.AddVertexAt(7, new Point2d(scale1 * 1.10176836567147, scale1 * 0.272945844641126), 0, 0, 0);
            poly55.AddVertexAt(8, new Point2d(scale1 * 1.10480748379364, scale1 * 0.274488781533918), 0, 0, 0);
            poly55.AddVertexAt(9, new Point2d(scale1 * 1.105742597062, scale1 * 0.278275990270772), 0, 0, 0);
            poly55.AddVertexAt(10, new Point2d(scale1 * 1.10270347893983, scale1 * 0.28037999512458), 0, 0, 0);
            poly55.AddVertexAt(11, new Point2d(scale1 * 1.09989813913476, scale1 * 0.282764533958894), 0, 0, 0);
            poly55.AddVertexAt(12, new Point2d(scale1 * 1.09756035596386, scale1 * 0.286271208715241), 0, 0, 0);
            poly55.AddVertexAt(13, new Point2d(scale1 * 1.09569012942714, scale1 * 0.288796014539811), 0, 0, 0);
            poly55.AddVertexAt(14, new Point2d(scale1 * 1.09148211971953, scale1 * 0.289777883471588), 0, 0, 0);
            poly55.AddVertexAt(15, new Point2d(scale1 * 1.08797544496318, scale1 * 0.290759752403364), 0, 0, 0);
            poly55.Closed = true;
            poly55.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly55.Layer = "0";
            poly55.Color = color_GP;
            poly55.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly55);
            Polyline poly56 = new Polyline();
            poly56.AddVertexAt(0, new Point2d(scale1 * 1.05836727263602, scale1 * 0.274437775355651), 0, 0, 0);
            poly56.AddVertexAt(1, new Point2d(scale1 * 1.06608195709998, scale1 * 0.272754571472604), 0, 0, 0);
            poly56.AddVertexAt(2, new Point2d(scale1 * 1.06935485353923, scale1 * 0.268125760794228), 0, 0, 0);
            poly56.AddVertexAt(3, new Point2d(scale1 * 1.06608195709998, scale1 * 0.263637217106105), 0, 0, 0);
            poly56.AddVertexAt(4, new Point2d(scale1 * 1.05836727263602, scale1 * 0.261954013223059), 0, 0, 0);
            poly56.AddVertexAt(5, new Point2d(scale1 * 1.0497174749037, scale1 * 0.261392945262043), 0, 0, 0);
            poly56.AddVertexAt(6, new Point2d(scale1 * 1.0450419085619, scale1 * 0.26237481419382), 0, 0, 0);
            poly56.AddVertexAt(7, new Point2d(scale1 * 1.04130145548847, scale1 * 0.26489962001839), 0, 0, 0);
            poly56.AddVertexAt(8, new Point2d(scale1 * 1.03078143121943, scale1 * 0.268967362735751), 0, 0, 0);
            poly56.AddVertexAt(9, new Point2d(scale1 * 1.02517075160927, scale1 * 0.272894838462859), 0, 0, 0);
            poly56.AddVertexAt(10, new Point2d(scale1 * 1.02657342151181, scale1 * 0.27471830933616), 0, 0, 0);
            poly56.AddVertexAt(11, new Point2d(scale1 * 1.03498944092704, scale1 * 0.278365251082757), 0, 0, 0);
            poly56.AddVertexAt(12, new Point2d(scale1 * 1.04153523380556, scale1 * 0.279066586034028), 0, 0, 0);
            poly56.AddVertexAt(13, new Point2d(scale1 * 1.04574324351317, scale1 * 0.277804183121743), 0, 0, 0);
            poly56.AddVertexAt(14, new Point2d(scale1 * 1.05135392312332, scale1 * 0.276541780209458), 0, 0, 0);
            poly56.AddVertexAt(15, new Point2d(scale1 * 1.05579571114803, scale1 * 0.275840445258188), 0, 0, 0);
            poly56.Closed = true;
            poly56.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly56.Layer = "0";
            poly56.Color = color_GP;
            poly56.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly56);
            Polyline poly57 = new Polyline();
            poly57.AddVertexAt(0, new Point2d(scale1 * 1.07889889791182, scale1 * 0.261890255500208), 0, 0, 0);
            poly57.AddVertexAt(1, new Point2d(scale1 * 1.07274102660718, scale1 * 0.262017770945895), 0, 0, 0);
            poly57.AddVertexAt(2, new Point2d(scale1 * 1.06743251686181, scale1 * 0.259722492923566), 0, 0, 0);
            poly57.AddVertexAt(3, new Point2d(scale1 * 1.06637081491274, scale1 * 0.256662122227115), 0, 0, 0);
            poly57.AddVertexAt(4, new Point2d(scale1 * 1.07167932465811, scale1 * 0.255769514107321), 0, 0, 0);
            poly57.AddVertexAt(5, new Point2d(scale1 * 1.07826187674237, scale1 * 0.2567896376728), 0, 0, 0);
            poly57.AddVertexAt(6, new Point2d(scale1 * 1.08208400375904, scale1 * 0.259467462032192), 0, 0, 0);
            poly57.AddVertexAt(7, new Point2d(scale1 * 1.07889889791182, scale1 * 0.261890255500208), 0, 0, 0);
            poly57.Closed = true;
            poly57.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly57.Layer = "0";
            poly57.Color = color_GP;
            poly57.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly57);
            Polyline poly58 = new Polyline();
            poly58.AddVertexAt(0, new Point2d(scale1 * 1.09647313832164, scale1 * 0.248628649148944), 0, 0, 0);
            poly58.AddVertexAt(1, new Point2d(scale1 * 1.09670691663873, scale1 * 0.243298503519296), 0, 0, 0);
            poly58.AddVertexAt(2, new Point2d(scale1 * 1.09507046841911, scale1 * 0.242036100607014), 0, 0, 0);
            poly58.AddVertexAt(3, new Point2d(scale1 * 1.09086245871149, scale1 * 0.240633430704475), 0, 0, 0);
            poly58.AddVertexAt(4, new Point2d(scale1 * 1.08688822732096, scale1 * 0.240773697694727), 0, 0, 0);
            poly58.AddVertexAt(5, new Point2d(scale1 * 1.0838491091988, scale1 * 0.242036100607014), 0, 0, 0);
            poly58.AddVertexAt(6, new Point2d(scale1 * 1.08057621275954, scale1 * 0.243298503519296), 0, 0, 0);
            poly58.AddVertexAt(7, new Point2d(scale1 * 1.07753709463737, scale1 * 0.244560906431583), 0, 0, 0);
            poly58.AddVertexAt(8, new Point2d(scale1 * 1.07566686810066, scale1 * 0.246945445265898), 0, 0, 0);
            poly58.AddVertexAt(9, new Point2d(scale1 * 1.07566686810066, scale1 * 0.248628649148944), 0, 0, 0);
            poly58.AddVertexAt(10, new Point2d(scale1 * 1.079173542857, scale1 * 0.252275590895543), 0, 0, 0);
            poly58.AddVertexAt(11, new Point2d(scale1 * 1.08595311405261, scale1 * 0.253818527788335), 0, 0, 0);
            poly58.AddVertexAt(12, new Point2d(scale1 * 1.09273268524821, scale1 * 0.252275590895543), 0, 0, 0);
            poly58.Closed = true;
            poly58.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly58.Layer = "0";
            poly58.Color = color_GP;
            poly58.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly58);
            Polyline poly59 = new Polyline();
            poly59.AddVertexAt(0, new Point2d(scale1 * 1.06210772570945, scale1 * 0.243438770509551), 0, 0, 0);
            poly59.AddVertexAt(1, new Point2d(scale1 * 1.05930238590438, scale1 * 0.241895833616759), 0, 0, 0);
            poly59.AddVertexAt(2, new Point2d(scale1 * 1.06023749917273, scale1 * 0.240072362743458), 0, 0, 0);
            poly59.AddVertexAt(3, new Point2d(scale1 * 1.06164016907527, scale1 * 0.237968357889651), 0, 0, 0);
            poly59.AddVertexAt(4, new Point2d(scale1 * 1.06467928719744, scale1 * 0.237687823909144), 0, 0, 0);
            poly59.AddVertexAt(5, new Point2d(scale1 * 1.06818596195379, scale1 * 0.23937102779219), 0, 0, 0);
            poly59.AddVertexAt(6, new Point2d(scale1 * 1.06818596195379, scale1 * 0.241615299636252), 0, 0, 0);
            poly59.AddVertexAt(7, new Point2d(scale1 * 1.06584817878289, scale1 * 0.244140105460821), 0, 0, 0);
            poly59.AddVertexAt(8, new Point2d(scale1 * 1.06210772570945, scale1 * 0.243438770509551), 0, 0, 0);
            poly59.Closed = true;
            poly59.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly59.Layer = "0";
            poly59.Color = color_GP;
            poly59.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly59);
            Polyline poly60 = new Polyline();
            poly60.AddVertexAt(0, new Point2d(scale1 * 1.01652095387695, scale1 * 0.248628649148944), 0, 0, 0);
            poly60.AddVertexAt(1, new Point2d(scale1 * 1.02306674675547, scale1 * 0.250452120022245), 0, 0, 0);
            poly60.AddVertexAt(2, new Point2d(scale1 * 1.0359245541954, scale1 * 0.249750785070974), 0, 0, 0);
            poly60.AddVertexAt(3, new Point2d(scale1 * 1.04247034707391, scale1 * 0.248628649148944), 0, 0, 0);
            poly60.AddVertexAt(4, new Point2d(scale1 * 1.04363923865936, scale1 * 0.246805178275643), 0, 0, 0);
            poly60.AddVertexAt(5, new Point2d(scale1 * 1.04387301697645, scale1 * 0.244981707402343), 0, 0, 0);
            poly60.AddVertexAt(6, new Point2d(scale1 * 1.04410679529354, scale1 * 0.242737435558282), 0, 0, 0);
            poly60.AddVertexAt(7, new Point2d(scale1 * 1.04597702183026, scale1 * 0.240773697694727), 0, 0, 0);
            poly60.AddVertexAt(8, new Point2d(scale1 * 1.04457435192772, scale1 * 0.237407289928637), 0, 0, 0);
            poly60.AddVertexAt(9, new Point2d(scale1 * 1.03943122895175, scale1 * 0.235583819055336), 0, 0, 0);
            poly60.AddVertexAt(10, new Point2d(scale1 * 1.02937876131689, scale1 * 0.235303285074829), 0, 0, 0);
            poly60.AddVertexAt(11, new Point2d(scale1 * 1.024469416658, scale1 * 0.236144887016352), 0, 0, 0);
            poly60.AddVertexAt(12, new Point2d(scale1 * 1.02096274190166, scale1 * 0.237828090899398), 0, 0, 0);
            poly60.AddVertexAt(13, new Point2d(scale1 * 1.0176898454624, scale1 * 0.241054231675236), 0, 0, 0);
            poly60.AddVertexAt(14, new Point2d(scale1 * 1.01581961892568, scale1 * 0.244420639441328), 0, 0, 0);
            poly60.Closed = true;
            poly60.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly60.Layer = "0";
            poly60.Color = color_GP;
            poly60.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly60);
            Polyline poly61 = new Polyline();
            poly61.AddVertexAt(0, new Point2d(scale1 * 0.960285741001707, scale1 * 0.3), 0, 0, 0);
            poly61.AddVertexAt(1, new Point2d(scale1 * 0.961640169075273, scale1 * 0.297968357889651), 0, 0, 0);
            poly61.AddVertexAt(2, new Point2d(scale1 * 0.964679287197439, scale1 * 0.297687823909144), 0, 0, 0);
            poly61.AddVertexAt(3, new Point2d(scale1 * 0.968185961953785, scale1 * 0.29937102779219), 0, 0, 0);
            poly61.AddVertexAt(4, new Point2d(scale1 * 0.968185961953785, scale1 * 0.3), 0, 0, 0);
            poly61.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly61.Layer = "0";
            poly61.Color = color_GP;
            poly61.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly61);
            Polyline poly62 = new Polyline();
            poly62.AddVertexAt(0, new Point2d(scale1 * 0.94565464779079, scale1 * 0.3), 0, 0, 0);
            poly62.AddVertexAt(1, new Point2d(scale1 * 0.944574351927722, scale1 * 0.297407289928637), 0, 0, 0);
            poly62.AddVertexAt(2, new Point2d(scale1 * 0.939431228951747, scale1 * 0.295583819055336), 0, 0, 0);
            poly62.AddVertexAt(3, new Point2d(scale1 * 0.929378761316889, scale1 * 0.295303285074829), 0, 0, 0);
            poly62.AddVertexAt(4, new Point2d(scale1 * 0.924469416658004, scale1 * 0.296144887016352), 0, 0, 0);
            poly62.AddVertexAt(5, new Point2d(scale1 * 0.920962741901658, scale1 * 0.297828090899398), 0, 0, 0);
            poly62.AddVertexAt(6, new Point2d(scale1 * 0.918759355857569, scale1 * 0.3), 0, 0, 0);
            poly62.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly62.Layer = "0";
            poly62.Color = color_GP;
            poly62.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly62);
            Polyline poly63 = new Polyline();
            poly63.AddVertexAt(0, new Point2d(scale1 * 0.958367272636016, scale1 * 0.282994061761134), 0, 0, 0);
            poly63.AddVertexAt(1, new Point2d(scale1 * 0.956497046099298, scale1 * 0.282853794770882), 0, 0, 0);
            poly63.AddVertexAt(2, new Point2d(scale1 * 0.954159262928401, scale1 * 0.28355512972215), 0, 0, 0);
            poly63.AddVertexAt(3, new Point2d(scale1 * 0.952289036391683, scale1 * 0.285518867585705), 0, 0, 0);
            poly63.AddVertexAt(4, new Point2d(scale1 * 0.953925484611311, scale1 * 0.287061804478495), 0, 0, 0);
            poly63.AddVertexAt(5, new Point2d(scale1 * 0.958367272636016, scale1 * 0.286220202536974), 0, 0, 0);
            poly63.AddVertexAt(6, new Point2d(scale1 * 0.959536164221465, scale1 * 0.284817532634435), 0, 0, 0);
            poly63.Closed = true;
            poly63.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly63.Layer = "0";
            poly63.Color = color_GP;
            poly63.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly63);
            Polyline poly64 = new Polyline();
            poly64.AddVertexAt(0, new Point2d(scale1 * 0.967952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly64.AddVertexAt(1, new Point2d(scale1 * 0.966081957099978, scale1 * 0.282012192829359), 0, 0, 0);
            poly64.AddVertexAt(2, new Point2d(scale1 * 0.965848178782888, scale1 * 0.280048454965803), 0, 0, 0);
            poly64.AddVertexAt(3, new Point2d(scale1 * 0.968419740270875, scale1 * 0.278505518073012), 0, 0, 0);
            poly64.AddVertexAt(4, new Point2d(scale1 * 0.971692636710131, scale1 * 0.278926319043773), 0, 0, 0);
            poly64.AddVertexAt(5, new Point2d(scale1 * 0.972393971661401, scale1 * 0.281030323897581), 0, 0, 0);
            poly64.AddVertexAt(6, new Point2d(scale1 * 0.971926415027221, scale1 * 0.282713527780627), 0, 0, 0);
            poly64.AddVertexAt(7, new Point2d(scale1 * 0.967952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly64.Closed = true;
            poly64.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly64.Layer = "0";
            poly64.Color = color_GP;
            poly64.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly64);
            Polyline poly65 = new Polyline();
            poly65.AddVertexAt(0, new Point2d(scale1 * 0.916577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly65.AddVertexAt(1, new Point2d(scale1 * 0.914239328956148, scale1 * 0.285251085149766), 0, 0, 0);
            poly65.AddVertexAt(2, new Point2d(scale1 * 0.915641998858686, scale1 * 0.283427614276465), 0, 0, 0);
            poly65.AddVertexAt(3, new Point2d(scale1 * 0.919616230249212, scale1 * 0.282726279325197), 0, 0, 0);
            poly65.AddVertexAt(4, new Point2d(scale1 * 0.922655348371378, scale1 * 0.283988682237481), 0, 0, 0);
            poly65.AddVertexAt(5, new Point2d(scale1 * 0.922655348371378, scale1 * 0.286653755052303), 0, 0, 0);
            poly65.AddVertexAt(6, new Point2d(scale1 * 0.916577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly65.Closed = true;
            poly65.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly65.Layer = "0";
            poly65.Color = color_GP;
            poly65.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly65);
            Polyline poly66 = new Polyline();
            poly66.AddVertexAt(0, new Point2d(scale1 * 0.911080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly66.AddVertexAt(1, new Point2d(scale1 * 0.909476961807309, scale1 * 0.263254670769046), 0, 0, 0);
            poly66.AddVertexAt(2, new Point2d(scale1 * 0.909710740124399, scale1 * 0.259607729022446), 0, 0, 0);
            poly66.AddVertexAt(3, new Point2d(scale1 * 0.910645853392758, scale1 * 0.256802389217368), 0, 0, 0);
            poly66.AddVertexAt(4, new Point2d(scale1 * 0.914853863100373, scale1 * 0.255960787275844), 0, 0, 0);
            poly66.AddVertexAt(5, new Point2d(scale1 * 0.919061872807989, scale1 * 0.256521855236861), 0, 0, 0);
            poly66.AddVertexAt(6, new Point2d(scale1 * 0.922568547564335, scale1 * 0.258625860090668), 0, 0, 0);
            poly66.AddVertexAt(7, new Point2d(scale1 * 0.919061872807989, scale1 * 0.263535204749552), 0, 0, 0);
            poly66.AddVertexAt(8, new Point2d(scale1 * 0.911080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly66.Closed = true;
            poly66.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly66.Layer = "0";
            poly66.Color = color_GP;
            poly66.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly66);
            Polyline poly67 = new Polyline();
            poly67.AddVertexAt(0, new Point2d(scale1 * 0.987975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly67.AddVertexAt(1, new Point2d(scale1 * 0.984702548523924, scale1 * 0.288936281530065), 0, 0, 0);
            poly67.AddVertexAt(2, new Point2d(scale1 * 0.982598543670117, scale1 * 0.283746402890672), 0, 0, 0);
            poly67.AddVertexAt(3, new Point2d(scale1 * 0.982598543670117, scale1 * 0.281221597066103), 0, 0, 0);
            poly67.AddVertexAt(4, new Point2d(scale1 * 0.981897208718848, scale1 * 0.275751184446203), 0, 0, 0);
            poly67.AddVertexAt(5, new Point2d(scale1 * 0.985403883475194, scale1 * 0.273366645611888), 0, 0, 0);
            poly67.AddVertexAt(6, new Point2d(scale1 * 0.990313228134079, scale1 * 0.272525043670365), 0, 0, 0);
            poly67.AddVertexAt(7, new Point2d(scale1 * 1.00176836567148, scale1 * 0.272945844641126), 0, 0, 0);
            poly67.AddVertexAt(8, new Point2d(scale1 * 1.00480748379364, scale1 * 0.274488781533918), 0, 0, 0);
            poly67.AddVertexAt(9, new Point2d(scale1 * 1.005742597062, scale1 * 0.278275990270772), 0, 0, 0);
            poly67.AddVertexAt(10, new Point2d(scale1 * 1.00270347893983, scale1 * 0.28037999512458), 0, 0, 0);
            poly67.AddVertexAt(11, new Point2d(scale1 * 0.999898139134758, scale1 * 0.282764533958894), 0, 0, 0);
            poly67.AddVertexAt(12, new Point2d(scale1 * 0.99756035596386, scale1 * 0.286271208715241), 0, 0, 0);
            poly67.AddVertexAt(13, new Point2d(scale1 * 0.995690129427142, scale1 * 0.288796014539811), 0, 0, 0);
            poly67.AddVertexAt(14, new Point2d(scale1 * 0.991482119719527, scale1 * 0.289777883471588), 0, 0, 0);
            poly67.AddVertexAt(15, new Point2d(scale1 * 0.987975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly67.Closed = true;
            poly67.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly67.Layer = "0";
            poly67.Color = color_GP;
            poly67.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly67);
            Polyline poly68 = new Polyline();
            poly68.AddVertexAt(0, new Point2d(scale1 * 0.958367272636016, scale1 * 0.274437775355651), 0, 0, 0);
            poly68.AddVertexAt(1, new Point2d(scale1 * 0.966081957099978, scale1 * 0.272754571472604), 0, 0, 0);
            poly68.AddVertexAt(2, new Point2d(scale1 * 0.969354853539234, scale1 * 0.268125760794228), 0, 0, 0);
            poly68.AddVertexAt(3, new Point2d(scale1 * 0.966081957099978, scale1 * 0.263637217106105), 0, 0, 0);
            poly68.AddVertexAt(4, new Point2d(scale1 * 0.958367272636016, scale1 * 0.261954013223059), 0, 0, 0);
            poly68.AddVertexAt(5, new Point2d(scale1 * 0.949717474903696, scale1 * 0.261392945262043), 0, 0, 0);
            poly68.AddVertexAt(6, new Point2d(scale1 * 0.945041908561901, scale1 * 0.26237481419382), 0, 0, 0);
            poly68.AddVertexAt(7, new Point2d(scale1 * 0.941301455488465, scale1 * 0.26489962001839), 0, 0, 0);
            poly68.AddVertexAt(8, new Point2d(scale1 * 0.930781431219427, scale1 * 0.268967362735751), 0, 0, 0);
            poly68.AddVertexAt(9, new Point2d(scale1 * 0.925170751609273, scale1 * 0.272894838462859), 0, 0, 0);
            poly68.AddVertexAt(10, new Point2d(scale1 * 0.926573421511812, scale1 * 0.27471830933616), 0, 0, 0);
            poly68.AddVertexAt(11, new Point2d(scale1 * 0.934989440927042, scale1 * 0.278365251082757), 0, 0, 0);
            poly68.AddVertexAt(12, new Point2d(scale1 * 0.941535233805555, scale1 * 0.279066586034028), 0, 0, 0);
            poly68.AddVertexAt(13, new Point2d(scale1 * 0.94574324351317, scale1 * 0.277804183121743), 0, 0, 0);
            poly68.AddVertexAt(14, new Point2d(scale1 * 0.951353923123324, scale1 * 0.276541780209458), 0, 0, 0);
            poly68.AddVertexAt(15, new Point2d(scale1 * 0.955795711148029, scale1 * 0.275840445258188), 0, 0, 0);
            poly68.Closed = true;
            poly68.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly68.Layer = "0";
            poly68.Color = color_GP;
            poly68.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly68);
            Polyline poly69 = new Polyline();
            poly69.AddVertexAt(0, new Point2d(scale1 * 0.978898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly69.AddVertexAt(1, new Point2d(scale1 * 0.972741026607183, scale1 * 0.262017770945895), 0, 0, 0);
            poly69.AddVertexAt(2, new Point2d(scale1 * 0.96743251686181, scale1 * 0.259722492923566), 0, 0, 0);
            poly69.AddVertexAt(3, new Point2d(scale1 * 0.966370814912736, scale1 * 0.256662122227115), 0, 0, 0);
            poly69.AddVertexAt(4, new Point2d(scale1 * 0.97167932465811, scale1 * 0.255769514107321), 0, 0, 0);
            poly69.AddVertexAt(5, new Point2d(scale1 * 0.978261876742373, scale1 * 0.2567896376728), 0, 0, 0);
            poly69.AddVertexAt(6, new Point2d(scale1 * 0.982084003759042, scale1 * 0.259467462032192), 0, 0, 0);
            poly69.AddVertexAt(7, new Point2d(scale1 * 0.978898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly69.Closed = true;
            poly69.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly69.Layer = "0";
            poly69.Color = color_GP;
            poly69.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly69);
            Polyline poly70 = new Polyline();
            poly70.AddVertexAt(0, new Point2d(scale1 * 0.996473138321644, scale1 * 0.248628649148944), 0, 0, 0);
            poly70.AddVertexAt(1, new Point2d(scale1 * 0.996706916638733, scale1 * 0.243298503519296), 0, 0, 0);
            poly70.AddVertexAt(2, new Point2d(scale1 * 0.995070468419106, scale1 * 0.242036100607014), 0, 0, 0);
            poly70.AddVertexAt(3, new Point2d(scale1 * 0.99086245871149, scale1 * 0.240633430704475), 0, 0, 0);
            poly70.AddVertexAt(4, new Point2d(scale1 * 0.986888227320964, scale1 * 0.240773697694727), 0, 0, 0);
            poly70.AddVertexAt(5, new Point2d(scale1 * 0.983849109198798, scale1 * 0.242036100607014), 0, 0, 0);
            poly70.AddVertexAt(6, new Point2d(scale1 * 0.980576212759542, scale1 * 0.243298503519296), 0, 0, 0);
            poly70.AddVertexAt(7, new Point2d(scale1 * 0.977537094637375, scale1 * 0.244560906431583), 0, 0, 0);
            poly70.AddVertexAt(8, new Point2d(scale1 * 0.975666868100657, scale1 * 0.246945445265898), 0, 0, 0);
            poly70.AddVertexAt(9, new Point2d(scale1 * 0.975666868100657, scale1 * 0.248628649148944), 0, 0, 0);
            poly70.AddVertexAt(10, new Point2d(scale1 * 0.979173542857003, scale1 * 0.252275590895543), 0, 0, 0);
            poly70.AddVertexAt(11, new Point2d(scale1 * 0.985953114052606, scale1 * 0.253818527788335), 0, 0, 0);
            poly70.AddVertexAt(12, new Point2d(scale1 * 0.992732685248208, scale1 * 0.252275590895543), 0, 0, 0);
            poly70.Closed = true;
            poly70.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly70.Layer = "0";
            poly70.Color = color_GP;
            poly70.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly70);
            Polyline poly71 = new Polyline();
            poly71.AddVertexAt(0, new Point2d(scale1 * 0.962107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly71.AddVertexAt(1, new Point2d(scale1 * 0.959302385904375, scale1 * 0.241895833616759), 0, 0, 0);
            poly71.AddVertexAt(2, new Point2d(scale1 * 0.960237499172734, scale1 * 0.240072362743458), 0, 0, 0);
            poly71.AddVertexAt(3, new Point2d(scale1 * 0.961640169075273, scale1 * 0.237968357889651), 0, 0, 0);
            poly71.AddVertexAt(4, new Point2d(scale1 * 0.964679287197439, scale1 * 0.237687823909144), 0, 0, 0);
            poly71.AddVertexAt(5, new Point2d(scale1 * 0.968185961953785, scale1 * 0.23937102779219), 0, 0, 0);
            poly71.AddVertexAt(6, new Point2d(scale1 * 0.968185961953785, scale1 * 0.241615299636252), 0, 0, 0);
            poly71.AddVertexAt(7, new Point2d(scale1 * 0.965848178782888, scale1 * 0.244140105460821), 0, 0, 0);
            poly71.AddVertexAt(8, new Point2d(scale1 * 0.962107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly71.Closed = true;
            poly71.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly71.Layer = "0";
            poly71.Color = color_GP;
            poly71.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly71);
            Polyline poly72 = new Polyline();
            poly72.AddVertexAt(0, new Point2d(scale1 * 0.916520953876953, scale1 * 0.248628649148944), 0, 0, 0);
            poly72.AddVertexAt(1, new Point2d(scale1 * 0.923066746755466, scale1 * 0.250452120022245), 0, 0, 0);
            poly72.AddVertexAt(2, new Point2d(scale1 * 0.935924554195401, scale1 * 0.249750785070974), 0, 0, 0);
            poly72.AddVertexAt(3, new Point2d(scale1 * 0.942470347073914, scale1 * 0.248628649148944), 0, 0, 0);
            poly72.AddVertexAt(4, new Point2d(scale1 * 0.943639238659363, scale1 * 0.246805178275643), 0, 0, 0);
            poly72.AddVertexAt(5, new Point2d(scale1 * 0.943873016976452, scale1 * 0.244981707402343), 0, 0, 0);
            poly72.AddVertexAt(6, new Point2d(scale1 * 0.944106795293542, scale1 * 0.242737435558282), 0, 0, 0);
            poly72.AddVertexAt(7, new Point2d(scale1 * 0.94597702183026, scale1 * 0.240773697694727), 0, 0, 0);
            poly72.AddVertexAt(8, new Point2d(scale1 * 0.944574351927722, scale1 * 0.237407289928637), 0, 0, 0);
            poly72.AddVertexAt(9, new Point2d(scale1 * 0.939431228951747, scale1 * 0.235583819055336), 0, 0, 0);
            poly72.AddVertexAt(10, new Point2d(scale1 * 0.929378761316889, scale1 * 0.235303285074829), 0, 0, 0);
            poly72.AddVertexAt(11, new Point2d(scale1 * 0.924469416658004, scale1 * 0.236144887016352), 0, 0, 0);
            poly72.AddVertexAt(12, new Point2d(scale1 * 0.920962741901658, scale1 * 0.237828090899398), 0, 0, 0);
            poly72.AddVertexAt(13, new Point2d(scale1 * 0.917689845462402, scale1 * 0.241054231675236), 0, 0, 0);
            poly72.AddVertexAt(14, new Point2d(scale1 * 0.915819618925684, scale1 * 0.244420639441328), 0, 0, 0);
            poly72.Closed = true;
            poly72.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly72.Layer = "0";
            poly72.Color = color_GP;
            poly72.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly72);
            Polyline poly73 = new Polyline();
            poly73.AddVertexAt(0, new Point2d(scale1 * 0.860285741001707, scale1 * 0.3), 0, 0, 0);
            poly73.AddVertexAt(1, new Point2d(scale1 * 0.861640169075273, scale1 * 0.297968357889651), 0, 0, 0);
            poly73.AddVertexAt(2, new Point2d(scale1 * 0.864679287197439, scale1 * 0.297687823909144), 0, 0, 0);
            poly73.AddVertexAt(3, new Point2d(scale1 * 0.868185961953785, scale1 * 0.29937102779219), 0, 0, 0);
            poly73.AddVertexAt(4, new Point2d(scale1 * 0.868185961953785, scale1 * 0.3), 0, 0, 0);
            poly73.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly73.Layer = "0";
            poly73.Color = color_GP;
            poly73.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly73);
            Polyline poly74 = new Polyline();
            poly74.AddVertexAt(0, new Point2d(scale1 * 0.84565464779079, scale1 * 0.3), 0, 0, 0);
            poly74.AddVertexAt(1, new Point2d(scale1 * 0.844574351927722, scale1 * 0.297407289928637), 0, 0, 0);
            poly74.AddVertexAt(2, new Point2d(scale1 * 0.839431228951747, scale1 * 0.295583819055336), 0, 0, 0);
            poly74.AddVertexAt(3, new Point2d(scale1 * 0.829378761316889, scale1 * 0.295303285074829), 0, 0, 0);
            poly74.AddVertexAt(4, new Point2d(scale1 * 0.824469416658004, scale1 * 0.296144887016352), 0, 0, 0);
            poly74.AddVertexAt(5, new Point2d(scale1 * 0.820962741901658, scale1 * 0.297828090899398), 0, 0, 0);
            poly74.AddVertexAt(6, new Point2d(scale1 * 0.818759355857569, scale1 * 0.3), 0, 0, 0);
            poly74.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly74.Layer = "0";
            poly74.Color = color_GP;
            poly74.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly74);
            Polyline poly75 = new Polyline();
            poly75.AddVertexAt(0, new Point2d(scale1 * 0.858367272636016, scale1 * 0.282994061761134), 0, 0, 0);
            poly75.AddVertexAt(1, new Point2d(scale1 * 0.856497046099298, scale1 * 0.282853794770882), 0, 0, 0);
            poly75.AddVertexAt(2, new Point2d(scale1 * 0.854159262928401, scale1 * 0.28355512972215), 0, 0, 0);
            poly75.AddVertexAt(3, new Point2d(scale1 * 0.852289036391683, scale1 * 0.285518867585705), 0, 0, 0);
            poly75.AddVertexAt(4, new Point2d(scale1 * 0.853925484611311, scale1 * 0.287061804478495), 0, 0, 0);
            poly75.AddVertexAt(5, new Point2d(scale1 * 0.858367272636016, scale1 * 0.286220202536974), 0, 0, 0);
            poly75.AddVertexAt(6, new Point2d(scale1 * 0.859536164221465, scale1 * 0.284817532634435), 0, 0, 0);
            poly75.Closed = true;
            poly75.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly75.Layer = "0";
            poly75.Color = color_GP;
            poly75.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly75);
            Polyline poly76 = new Polyline();
            poly76.AddVertexAt(0, new Point2d(scale1 * 0.867952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly76.AddVertexAt(1, new Point2d(scale1 * 0.866081957099978, scale1 * 0.282012192829359), 0, 0, 0);
            poly76.AddVertexAt(2, new Point2d(scale1 * 0.865848178782888, scale1 * 0.280048454965803), 0, 0, 0);
            poly76.AddVertexAt(3, new Point2d(scale1 * 0.868419740270875, scale1 * 0.278505518073012), 0, 0, 0);
            poly76.AddVertexAt(4, new Point2d(scale1 * 0.871692636710131, scale1 * 0.278926319043773), 0, 0, 0);
            poly76.AddVertexAt(5, new Point2d(scale1 * 0.872393971661401, scale1 * 0.281030323897581), 0, 0, 0);
            poly76.AddVertexAt(6, new Point2d(scale1 * 0.871926415027221, scale1 * 0.282713527780627), 0, 0, 0);
            poly76.AddVertexAt(7, new Point2d(scale1 * 0.867952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly76.Closed = true;
            poly76.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly76.Layer = "0";
            poly76.Color = color_GP;
            poly76.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly76);
            Polyline poly77 = new Polyline();
            poly77.AddVertexAt(0, new Point2d(scale1 * 0.816577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly77.AddVertexAt(1, new Point2d(scale1 * 0.814239328956148, scale1 * 0.285251085149766), 0, 0, 0);
            poly77.AddVertexAt(2, new Point2d(scale1 * 0.815641998858686, scale1 * 0.283427614276465), 0, 0, 0);
            poly77.AddVertexAt(3, new Point2d(scale1 * 0.819616230249212, scale1 * 0.282726279325197), 0, 0, 0);
            poly77.AddVertexAt(4, new Point2d(scale1 * 0.822655348371378, scale1 * 0.283988682237481), 0, 0, 0);
            poly77.AddVertexAt(5, new Point2d(scale1 * 0.822655348371378, scale1 * 0.286653755052303), 0, 0, 0);
            poly77.AddVertexAt(6, new Point2d(scale1 * 0.816577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly77.Closed = true;
            poly77.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly77.Layer = "0";
            poly77.Color = color_GP;
            poly77.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly77);
            Polyline poly78 = new Polyline();
            poly78.AddVertexAt(0, new Point2d(scale1 * 0.811080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly78.AddVertexAt(1, new Point2d(scale1 * 0.809476961807309, scale1 * 0.263254670769046), 0, 0, 0);
            poly78.AddVertexAt(2, new Point2d(scale1 * 0.809710740124399, scale1 * 0.259607729022446), 0, 0, 0);
            poly78.AddVertexAt(3, new Point2d(scale1 * 0.810645853392758, scale1 * 0.256802389217368), 0, 0, 0);
            poly78.AddVertexAt(4, new Point2d(scale1 * 0.814853863100373, scale1 * 0.255960787275844), 0, 0, 0);
            poly78.AddVertexAt(5, new Point2d(scale1 * 0.819061872807989, scale1 * 0.256521855236861), 0, 0, 0);
            poly78.AddVertexAt(6, new Point2d(scale1 * 0.822568547564335, scale1 * 0.258625860090668), 0, 0, 0);
            poly78.AddVertexAt(7, new Point2d(scale1 * 0.819061872807989, scale1 * 0.263535204749552), 0, 0, 0);
            poly78.AddVertexAt(8, new Point2d(scale1 * 0.811080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly78.Closed = true;
            poly78.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly78.Layer = "0";
            poly78.Color = color_GP;
            poly78.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly78);
            Polyline poly79 = new Polyline();
            poly79.AddVertexAt(0, new Point2d(scale1 * 0.887975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly79.AddVertexAt(1, new Point2d(scale1 * 0.884702548523924, scale1 * 0.288936281530065), 0, 0, 0);
            poly79.AddVertexAt(2, new Point2d(scale1 * 0.882598543670117, scale1 * 0.283746402890672), 0, 0, 0);
            poly79.AddVertexAt(3, new Point2d(scale1 * 0.882598543670117, scale1 * 0.281221597066103), 0, 0, 0);
            poly79.AddVertexAt(4, new Point2d(scale1 * 0.881897208718848, scale1 * 0.275751184446203), 0, 0, 0);
            poly79.AddVertexAt(5, new Point2d(scale1 * 0.885403883475194, scale1 * 0.273366645611888), 0, 0, 0);
            poly79.AddVertexAt(6, new Point2d(scale1 * 0.890313228134079, scale1 * 0.272525043670365), 0, 0, 0);
            poly79.AddVertexAt(7, new Point2d(scale1 * 0.901768365671475, scale1 * 0.272945844641126), 0, 0, 0);
            poly79.AddVertexAt(8, new Point2d(scale1 * 0.904807483793642, scale1 * 0.274488781533918), 0, 0, 0);
            poly79.AddVertexAt(9, new Point2d(scale1 * 0.905742597062001, scale1 * 0.278275990270772), 0, 0, 0);
            poly79.AddVertexAt(10, new Point2d(scale1 * 0.902703478939834, scale1 * 0.28037999512458), 0, 0, 0);
            poly79.AddVertexAt(11, new Point2d(scale1 * 0.899898139134758, scale1 * 0.282764533958894), 0, 0, 0);
            poly79.AddVertexAt(12, new Point2d(scale1 * 0.89756035596386, scale1 * 0.286271208715241), 0, 0, 0);
            poly79.AddVertexAt(13, new Point2d(scale1 * 0.895690129427142, scale1 * 0.288796014539811), 0, 0, 0);
            poly79.AddVertexAt(14, new Point2d(scale1 * 0.891482119719527, scale1 * 0.289777883471588), 0, 0, 0);
            poly79.AddVertexAt(15, new Point2d(scale1 * 0.887975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly79.Closed = true;
            poly79.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly79.Layer = "0";
            poly79.Color = color_GP;
            poly79.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly79);
            Polyline poly80 = new Polyline();
            poly80.AddVertexAt(0, new Point2d(scale1 * 0.858367272636016, scale1 * 0.274437775355651), 0, 0, 0);
            poly80.AddVertexAt(1, new Point2d(scale1 * 0.866081957099978, scale1 * 0.272754571472604), 0, 0, 0);
            poly80.AddVertexAt(2, new Point2d(scale1 * 0.869354853539234, scale1 * 0.268125760794228), 0, 0, 0);
            poly80.AddVertexAt(3, new Point2d(scale1 * 0.866081957099978, scale1 * 0.263637217106105), 0, 0, 0);
            poly80.AddVertexAt(4, new Point2d(scale1 * 0.858367272636016, scale1 * 0.261954013223059), 0, 0, 0);
            poly80.AddVertexAt(5, new Point2d(scale1 * 0.849717474903696, scale1 * 0.261392945262043), 0, 0, 0);
            poly80.AddVertexAt(6, new Point2d(scale1 * 0.845041908561901, scale1 * 0.26237481419382), 0, 0, 0);
            poly80.AddVertexAt(7, new Point2d(scale1 * 0.841301455488465, scale1 * 0.26489962001839), 0, 0, 0);
            poly80.AddVertexAt(8, new Point2d(scale1 * 0.830781431219427, scale1 * 0.268967362735751), 0, 0, 0);
            poly80.AddVertexAt(9, new Point2d(scale1 * 0.825170751609273, scale1 * 0.272894838462859), 0, 0, 0);
            poly80.AddVertexAt(10, new Point2d(scale1 * 0.826573421511812, scale1 * 0.27471830933616), 0, 0, 0);
            poly80.AddVertexAt(11, new Point2d(scale1 * 0.834989440927042, scale1 * 0.278365251082757), 0, 0, 0);
            poly80.AddVertexAt(12, new Point2d(scale1 * 0.841535233805555, scale1 * 0.279066586034028), 0, 0, 0);
            poly80.AddVertexAt(13, new Point2d(scale1 * 0.84574324351317, scale1 * 0.277804183121743), 0, 0, 0);
            poly80.AddVertexAt(14, new Point2d(scale1 * 0.851353923123324, scale1 * 0.276541780209458), 0, 0, 0);
            poly80.AddVertexAt(15, new Point2d(scale1 * 0.855795711148029, scale1 * 0.275840445258188), 0, 0, 0);
            poly80.Closed = true;
            poly80.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly80.Layer = "0";
            poly80.Color = color_GP;
            poly80.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly80);
            Polyline poly81 = new Polyline();
            poly81.AddVertexAt(0, new Point2d(scale1 * 0.878898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly81.AddVertexAt(1, new Point2d(scale1 * 0.872741026607183, scale1 * 0.262017770945895), 0, 0, 0);
            poly81.AddVertexAt(2, new Point2d(scale1 * 0.86743251686181, scale1 * 0.259722492923566), 0, 0, 0);
            poly81.AddVertexAt(3, new Point2d(scale1 * 0.866370814912736, scale1 * 0.256662122227115), 0, 0, 0);
            poly81.AddVertexAt(4, new Point2d(scale1 * 0.87167932465811, scale1 * 0.255769514107321), 0, 0, 0);
            poly81.AddVertexAt(5, new Point2d(scale1 * 0.878261876742373, scale1 * 0.2567896376728), 0, 0, 0);
            poly81.AddVertexAt(6, new Point2d(scale1 * 0.882084003759042, scale1 * 0.259467462032192), 0, 0, 0);
            poly81.AddVertexAt(7, new Point2d(scale1 * 0.878898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly81.Closed = true;
            poly81.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly81.Layer = "0";
            poly81.Color = color_GP;
            poly81.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly81);
            Polyline poly82 = new Polyline();
            poly82.AddVertexAt(0, new Point2d(scale1 * 0.896473138321644, scale1 * 0.248628649148944), 0, 0, 0);
            poly82.AddVertexAt(1, new Point2d(scale1 * 0.896706916638733, scale1 * 0.243298503519296), 0, 0, 0);
            poly82.AddVertexAt(2, new Point2d(scale1 * 0.895070468419106, scale1 * 0.242036100607014), 0, 0, 0);
            poly82.AddVertexAt(3, new Point2d(scale1 * 0.89086245871149, scale1 * 0.240633430704475), 0, 0, 0);
            poly82.AddVertexAt(4, new Point2d(scale1 * 0.886888227320964, scale1 * 0.240773697694727), 0, 0, 0);
            poly82.AddVertexAt(5, new Point2d(scale1 * 0.883849109198798, scale1 * 0.242036100607014), 0, 0, 0);
            poly82.AddVertexAt(6, new Point2d(scale1 * 0.880576212759542, scale1 * 0.243298503519296), 0, 0, 0);
            poly82.AddVertexAt(7, new Point2d(scale1 * 0.877537094637375, scale1 * 0.244560906431583), 0, 0, 0);
            poly82.AddVertexAt(8, new Point2d(scale1 * 0.875666868100657, scale1 * 0.246945445265898), 0, 0, 0);
            poly82.AddVertexAt(9, new Point2d(scale1 * 0.875666868100657, scale1 * 0.248628649148944), 0, 0, 0);
            poly82.AddVertexAt(10, new Point2d(scale1 * 0.879173542857003, scale1 * 0.252275590895543), 0, 0, 0);
            poly82.AddVertexAt(11, new Point2d(scale1 * 0.885953114052606, scale1 * 0.253818527788335), 0, 0, 0);
            poly82.AddVertexAt(12, new Point2d(scale1 * 0.892732685248208, scale1 * 0.252275590895543), 0, 0, 0);
            poly82.Closed = true;
            poly82.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly82.Layer = "0";
            poly82.Color = color_GP;
            poly82.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly82);
            Polyline poly83 = new Polyline();
            poly83.AddVertexAt(0, new Point2d(scale1 * 0.862107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly83.AddVertexAt(1, new Point2d(scale1 * 0.859302385904375, scale1 * 0.241895833616759), 0, 0, 0);
            poly83.AddVertexAt(2, new Point2d(scale1 * 0.860237499172734, scale1 * 0.240072362743458), 0, 0, 0);
            poly83.AddVertexAt(3, new Point2d(scale1 * 0.861640169075273, scale1 * 0.237968357889651), 0, 0, 0);
            poly83.AddVertexAt(4, new Point2d(scale1 * 0.864679287197439, scale1 * 0.237687823909144), 0, 0, 0);
            poly83.AddVertexAt(5, new Point2d(scale1 * 0.868185961953785, scale1 * 0.23937102779219), 0, 0, 0);
            poly83.AddVertexAt(6, new Point2d(scale1 * 0.868185961953785, scale1 * 0.241615299636252), 0, 0, 0);
            poly83.AddVertexAt(7, new Point2d(scale1 * 0.865848178782888, scale1 * 0.244140105460821), 0, 0, 0);
            poly83.AddVertexAt(8, new Point2d(scale1 * 0.862107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly83.Closed = true;
            poly83.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly83.Layer = "0";
            poly83.Color = color_GP;
            poly83.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly83);
            Polyline poly84 = new Polyline();
            poly84.AddVertexAt(0, new Point2d(scale1 * 0.816520953876953, scale1 * 0.248628649148944), 0, 0, 0);
            poly84.AddVertexAt(1, new Point2d(scale1 * 0.823066746755466, scale1 * 0.250452120022245), 0, 0, 0);
            poly84.AddVertexAt(2, new Point2d(scale1 * 0.835924554195401, scale1 * 0.249750785070974), 0, 0, 0);
            poly84.AddVertexAt(3, new Point2d(scale1 * 0.842470347073914, scale1 * 0.248628649148944), 0, 0, 0);
            poly84.AddVertexAt(4, new Point2d(scale1 * 0.843639238659363, scale1 * 0.246805178275643), 0, 0, 0);
            poly84.AddVertexAt(5, new Point2d(scale1 * 0.843873016976453, scale1 * 0.244981707402343), 0, 0, 0);
            poly84.AddVertexAt(6, new Point2d(scale1 * 0.844106795293542, scale1 * 0.242737435558282), 0, 0, 0);
            poly84.AddVertexAt(7, new Point2d(scale1 * 0.84597702183026, scale1 * 0.240773697694727), 0, 0, 0);
            poly84.AddVertexAt(8, new Point2d(scale1 * 0.844574351927722, scale1 * 0.237407289928637), 0, 0, 0);
            poly84.AddVertexAt(9, new Point2d(scale1 * 0.839431228951747, scale1 * 0.235583819055336), 0, 0, 0);
            poly84.AddVertexAt(10, new Point2d(scale1 * 0.829378761316889, scale1 * 0.235303285074829), 0, 0, 0);
            poly84.AddVertexAt(11, new Point2d(scale1 * 0.824469416658004, scale1 * 0.236144887016352), 0, 0, 0);
            poly84.AddVertexAt(12, new Point2d(scale1 * 0.820962741901658, scale1 * 0.237828090899398), 0, 0, 0);
            poly84.AddVertexAt(13, new Point2d(scale1 * 0.817689845462402, scale1 * 0.241054231675236), 0, 0, 0);
            poly84.AddVertexAt(14, new Point2d(scale1 * 0.815819618925684, scale1 * 0.244420639441328), 0, 0, 0);
            poly84.Closed = true;
            poly84.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly84.Layer = "0";
            poly84.Color = color_GP;
            poly84.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly84);
            Polyline poly85 = new Polyline();
            poly85.AddVertexAt(0, new Point2d(scale1 * 0.760285741001707, scale1 * 0.3), 0, 0, 0);
            poly85.AddVertexAt(1, new Point2d(scale1 * 0.761640169075273, scale1 * 0.297968357889651), 0, 0, 0);
            poly85.AddVertexAt(2, new Point2d(scale1 * 0.764679287197439, scale1 * 0.297687823909144), 0, 0, 0);
            poly85.AddVertexAt(3, new Point2d(scale1 * 0.768185961953785, scale1 * 0.29937102779219), 0, 0, 0);
            poly85.AddVertexAt(4, new Point2d(scale1 * 0.768185961953785, scale1 * 0.3), 0, 0, 0);
            poly85.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly85.Layer = "0";
            poly85.Color = color_GP;
            poly85.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly85);
            Polyline poly86 = new Polyline();
            poly86.AddVertexAt(0, new Point2d(scale1 * 0.74565464779079, scale1 * 0.3), 0, 0, 0);
            poly86.AddVertexAt(1, new Point2d(scale1 * 0.744574351927722, scale1 * 0.297407289928637), 0, 0, 0);
            poly86.AddVertexAt(2, new Point2d(scale1 * 0.739431228951748, scale1 * 0.295583819055336), 0, 0, 0);
            poly86.AddVertexAt(3, new Point2d(scale1 * 0.729378761316889, scale1 * 0.295303285074829), 0, 0, 0);
            poly86.AddVertexAt(4, new Point2d(scale1 * 0.724469416658004, scale1 * 0.296144887016352), 0, 0, 0);
            poly86.AddVertexAt(5, new Point2d(scale1 * 0.720962741901658, scale1 * 0.297828090899398), 0, 0, 0);
            poly86.AddVertexAt(6, new Point2d(scale1 * 0.718759355857569, scale1 * 0.3), 0, 0, 0);
            poly86.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly86.Layer = "0";
            poly86.Color = color_GP;
            poly86.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly86);
            Polyline poly87 = new Polyline();
            poly87.AddVertexAt(0, new Point2d(scale1 * 0.758367272636016, scale1 * 0.282994061761134), 0, 0, 0);
            poly87.AddVertexAt(1, new Point2d(scale1 * 0.756497046099298, scale1 * 0.282853794770882), 0, 0, 0);
            poly87.AddVertexAt(2, new Point2d(scale1 * 0.754159262928401, scale1 * 0.28355512972215), 0, 0, 0);
            poly87.AddVertexAt(3, new Point2d(scale1 * 0.752289036391683, scale1 * 0.285518867585705), 0, 0, 0);
            poly87.AddVertexAt(4, new Point2d(scale1 * 0.753925484611311, scale1 * 0.287061804478495), 0, 0, 0);
            poly87.AddVertexAt(5, new Point2d(scale1 * 0.758367272636016, scale1 * 0.286220202536974), 0, 0, 0);
            poly87.AddVertexAt(6, new Point2d(scale1 * 0.759536164221465, scale1 * 0.284817532634435), 0, 0, 0);
            poly87.Closed = true;
            poly87.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly87.Layer = "0";
            poly87.Color = color_GP;
            poly87.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly87);
            Polyline poly88 = new Polyline();
            poly88.AddVertexAt(0, new Point2d(scale1 * 0.767952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly88.AddVertexAt(1, new Point2d(scale1 * 0.766081957099978, scale1 * 0.282012192829359), 0, 0, 0);
            poly88.AddVertexAt(2, new Point2d(scale1 * 0.765848178782888, scale1 * 0.280048454965803), 0, 0, 0);
            poly88.AddVertexAt(3, new Point2d(scale1 * 0.768419740270875, scale1 * 0.278505518073012), 0, 0, 0);
            poly88.AddVertexAt(4, new Point2d(scale1 * 0.771692636710131, scale1 * 0.278926319043773), 0, 0, 0);
            poly88.AddVertexAt(5, new Point2d(scale1 * 0.772393971661401, scale1 * 0.281030323897581), 0, 0, 0);
            poly88.AddVertexAt(6, new Point2d(scale1 * 0.771926415027221, scale1 * 0.282713527780627), 0, 0, 0);
            poly88.AddVertexAt(7, new Point2d(scale1 * 0.767952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly88.Closed = true;
            poly88.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly88.Layer = "0";
            poly88.Color = color_GP;
            poly88.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly88);
            Polyline poly89 = new Polyline();
            poly89.AddVertexAt(0, new Point2d(scale1 * 0.716577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly89.AddVertexAt(1, new Point2d(scale1 * 0.714239328956148, scale1 * 0.285251085149766), 0, 0, 0);
            poly89.AddVertexAt(2, new Point2d(scale1 * 0.715641998858686, scale1 * 0.283427614276465), 0, 0, 0);
            poly89.AddVertexAt(3, new Point2d(scale1 * 0.719616230249212, scale1 * 0.282726279325197), 0, 0, 0);
            poly89.AddVertexAt(4, new Point2d(scale1 * 0.722655348371378, scale1 * 0.283988682237481), 0, 0, 0);
            poly89.AddVertexAt(5, new Point2d(scale1 * 0.722655348371378, scale1 * 0.286653755052303), 0, 0, 0);
            poly89.AddVertexAt(6, new Point2d(scale1 * 0.716577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly89.Closed = true;
            poly89.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly89.Layer = "0";
            poly89.Color = color_GP;
            poly89.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly89);
            Polyline poly90 = new Polyline();
            poly90.AddVertexAt(0, new Point2d(scale1 * 0.711080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly90.AddVertexAt(1, new Point2d(scale1 * 0.70947696180731, scale1 * 0.263254670769046), 0, 0, 0);
            poly90.AddVertexAt(2, new Point2d(scale1 * 0.709710740124399, scale1 * 0.259607729022446), 0, 0, 0);
            poly90.AddVertexAt(3, new Point2d(scale1 * 0.710645853392758, scale1 * 0.256802389217368), 0, 0, 0);
            poly90.AddVertexAt(4, new Point2d(scale1 * 0.714853863100373, scale1 * 0.255960787275844), 0, 0, 0);
            poly90.AddVertexAt(5, new Point2d(scale1 * 0.719061872807989, scale1 * 0.256521855236861), 0, 0, 0);
            poly90.AddVertexAt(6, new Point2d(scale1 * 0.722568547564335, scale1 * 0.258625860090668), 0, 0, 0);
            poly90.AddVertexAt(7, new Point2d(scale1 * 0.719061872807989, scale1 * 0.263535204749552), 0, 0, 0);
            poly90.AddVertexAt(8, new Point2d(scale1 * 0.711080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly90.Closed = true;
            poly90.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly90.Layer = "0";
            poly90.Color = color_GP;
            poly90.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly90);
            Polyline poly91 = new Polyline();
            poly91.AddVertexAt(0, new Point2d(scale1 * 0.787975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly91.AddVertexAt(1, new Point2d(scale1 * 0.784702548523924, scale1 * 0.288936281530065), 0, 0, 0);
            poly91.AddVertexAt(2, new Point2d(scale1 * 0.782598543670117, scale1 * 0.283746402890672), 0, 0, 0);
            poly91.AddVertexAt(3, new Point2d(scale1 * 0.782598543670117, scale1 * 0.281221597066103), 0, 0, 0);
            poly91.AddVertexAt(4, new Point2d(scale1 * 0.781897208718848, scale1 * 0.275751184446203), 0, 0, 0);
            poly91.AddVertexAt(5, new Point2d(scale1 * 0.785403883475194, scale1 * 0.273366645611888), 0, 0, 0);
            poly91.AddVertexAt(6, new Point2d(scale1 * 0.790313228134079, scale1 * 0.272525043670365), 0, 0, 0);
            poly91.AddVertexAt(7, new Point2d(scale1 * 0.801768365671475, scale1 * 0.272945844641126), 0, 0, 0);
            poly91.AddVertexAt(8, new Point2d(scale1 * 0.804807483793642, scale1 * 0.274488781533918), 0, 0, 0);
            poly91.AddVertexAt(9, new Point2d(scale1 * 0.805742597062001, scale1 * 0.278275990270772), 0, 0, 0);
            poly91.AddVertexAt(10, new Point2d(scale1 * 0.802703478939834, scale1 * 0.28037999512458), 0, 0, 0);
            poly91.AddVertexAt(11, new Point2d(scale1 * 0.799898139134758, scale1 * 0.282764533958894), 0, 0, 0);
            poly91.AddVertexAt(12, new Point2d(scale1 * 0.79756035596386, scale1 * 0.286271208715241), 0, 0, 0);
            poly91.AddVertexAt(13, new Point2d(scale1 * 0.795690129427142, scale1 * 0.288796014539811), 0, 0, 0);
            poly91.AddVertexAt(14, new Point2d(scale1 * 0.791482119719527, scale1 * 0.289777883471588), 0, 0, 0);
            poly91.AddVertexAt(15, new Point2d(scale1 * 0.787975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly91.Closed = true;
            poly91.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly91.Layer = "0";
            poly91.Color = color_GP;
            poly91.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly91);
            Polyline poly92 = new Polyline();
            poly92.AddVertexAt(0, new Point2d(scale1 * 0.758367272636016, scale1 * 0.274437775355651), 0, 0, 0);
            poly92.AddVertexAt(1, new Point2d(scale1 * 0.766081957099978, scale1 * 0.272754571472604), 0, 0, 0);
            poly92.AddVertexAt(2, new Point2d(scale1 * 0.769354853539234, scale1 * 0.268125760794228), 0, 0, 0);
            poly92.AddVertexAt(3, new Point2d(scale1 * 0.766081957099978, scale1 * 0.263637217106105), 0, 0, 0);
            poly92.AddVertexAt(4, new Point2d(scale1 * 0.758367272636016, scale1 * 0.261954013223059), 0, 0, 0);
            poly92.AddVertexAt(5, new Point2d(scale1 * 0.749717474903696, scale1 * 0.261392945262043), 0, 0, 0);
            poly92.AddVertexAt(6, new Point2d(scale1 * 0.745041908561901, scale1 * 0.26237481419382), 0, 0, 0);
            poly92.AddVertexAt(7, new Point2d(scale1 * 0.741301455488465, scale1 * 0.26489962001839), 0, 0, 0);
            poly92.AddVertexAt(8, new Point2d(scale1 * 0.730781431219427, scale1 * 0.268967362735751), 0, 0, 0);
            poly92.AddVertexAt(9, new Point2d(scale1 * 0.725170751609273, scale1 * 0.272894838462859), 0, 0, 0);
            poly92.AddVertexAt(10, new Point2d(scale1 * 0.726573421511812, scale1 * 0.27471830933616), 0, 0, 0);
            poly92.AddVertexAt(11, new Point2d(scale1 * 0.734989440927042, scale1 * 0.278365251082757), 0, 0, 0);
            poly92.AddVertexAt(12, new Point2d(scale1 * 0.741535233805555, scale1 * 0.279066586034028), 0, 0, 0);
            poly92.AddVertexAt(13, new Point2d(scale1 * 0.74574324351317, scale1 * 0.277804183121743), 0, 0, 0);
            poly92.AddVertexAt(14, new Point2d(scale1 * 0.751353923123324, scale1 * 0.276541780209458), 0, 0, 0);
            poly92.AddVertexAt(15, new Point2d(scale1 * 0.755795711148029, scale1 * 0.275840445258188), 0, 0, 0);
            poly92.Closed = true;
            poly92.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly92.Layer = "0";
            poly92.Color = color_GP;
            poly92.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly92);
            Polyline poly93 = new Polyline();
            poly93.AddVertexAt(0, new Point2d(scale1 * 0.778898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly93.AddVertexAt(1, new Point2d(scale1 * 0.772741026607183, scale1 * 0.262017770945895), 0, 0, 0);
            poly93.AddVertexAt(2, new Point2d(scale1 * 0.767432516861809, scale1 * 0.259722492923566), 0, 0, 0);
            poly93.AddVertexAt(3, new Point2d(scale1 * 0.766370814912736, scale1 * 0.256662122227115), 0, 0, 0);
            poly93.AddVertexAt(4, new Point2d(scale1 * 0.77167932465811, scale1 * 0.255769514107321), 0, 0, 0);
            poly93.AddVertexAt(5, new Point2d(scale1 * 0.778261876742373, scale1 * 0.2567896376728), 0, 0, 0);
            poly93.AddVertexAt(6, new Point2d(scale1 * 0.782084003759042, scale1 * 0.259467462032192), 0, 0, 0);
            poly93.AddVertexAt(7, new Point2d(scale1 * 0.778898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly93.Closed = true;
            poly93.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly93.Layer = "0";
            poly93.Color = color_GP;
            poly93.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly93);
            Polyline poly94 = new Polyline();
            poly94.AddVertexAt(0, new Point2d(scale1 * 0.796473138321644, scale1 * 0.248628649148944), 0, 0, 0);
            poly94.AddVertexAt(1, new Point2d(scale1 * 0.796706916638733, scale1 * 0.243298503519296), 0, 0, 0);
            poly94.AddVertexAt(2, new Point2d(scale1 * 0.795070468419106, scale1 * 0.242036100607014), 0, 0, 0);
            poly94.AddVertexAt(3, new Point2d(scale1 * 0.79086245871149, scale1 * 0.240633430704475), 0, 0, 0);
            poly94.AddVertexAt(4, new Point2d(scale1 * 0.786888227320964, scale1 * 0.240773697694727), 0, 0, 0);
            poly94.AddVertexAt(5, new Point2d(scale1 * 0.783849109198798, scale1 * 0.242036100607014), 0, 0, 0);
            poly94.AddVertexAt(6, new Point2d(scale1 * 0.780576212759542, scale1 * 0.243298503519296), 0, 0, 0);
            poly94.AddVertexAt(7, new Point2d(scale1 * 0.777537094637375, scale1 * 0.244560906431583), 0, 0, 0);
            poly94.AddVertexAt(8, new Point2d(scale1 * 0.775666868100657, scale1 * 0.246945445265898), 0, 0, 0);
            poly94.AddVertexAt(9, new Point2d(scale1 * 0.775666868100657, scale1 * 0.248628649148944), 0, 0, 0);
            poly94.AddVertexAt(10, new Point2d(scale1 * 0.779173542857003, scale1 * 0.252275590895543), 0, 0, 0);
            poly94.AddVertexAt(11, new Point2d(scale1 * 0.785953114052606, scale1 * 0.253818527788335), 0, 0, 0);
            poly94.AddVertexAt(12, new Point2d(scale1 * 0.792732685248208, scale1 * 0.252275590895543), 0, 0, 0);
            poly94.Closed = true;
            poly94.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly94.Layer = "0";
            poly94.Color = color_GP;
            poly94.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly94);
            Polyline poly95 = new Polyline();
            poly95.AddVertexAt(0, new Point2d(scale1 * 0.762107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly95.AddVertexAt(1, new Point2d(scale1 * 0.759302385904375, scale1 * 0.241895833616759), 0, 0, 0);
            poly95.AddVertexAt(2, new Point2d(scale1 * 0.760237499172734, scale1 * 0.240072362743458), 0, 0, 0);
            poly95.AddVertexAt(3, new Point2d(scale1 * 0.761640169075273, scale1 * 0.237968357889651), 0, 0, 0);
            poly95.AddVertexAt(4, new Point2d(scale1 * 0.764679287197439, scale1 * 0.237687823909144), 0, 0, 0);
            poly95.AddVertexAt(5, new Point2d(scale1 * 0.768185961953785, scale1 * 0.23937102779219), 0, 0, 0);
            poly95.AddVertexAt(6, new Point2d(scale1 * 0.768185961953785, scale1 * 0.241615299636252), 0, 0, 0);
            poly95.AddVertexAt(7, new Point2d(scale1 * 0.765848178782888, scale1 * 0.244140105460821), 0, 0, 0);
            poly95.AddVertexAt(8, new Point2d(scale1 * 0.762107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly95.Closed = true;
            poly95.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly95.Layer = "0";
            poly95.Color = color_GP;
            poly95.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly95);
            Polyline poly96 = new Polyline();
            poly96.AddVertexAt(0, new Point2d(scale1 * 0.716520953876953, scale1 * 0.248628649148944), 0, 0, 0);
            poly96.AddVertexAt(1, new Point2d(scale1 * 0.723066746755466, scale1 * 0.250452120022245), 0, 0, 0);
            poly96.AddVertexAt(2, new Point2d(scale1 * 0.735924554195401, scale1 * 0.249750785070974), 0, 0, 0);
            poly96.AddVertexAt(3, new Point2d(scale1 * 0.742470347073914, scale1 * 0.248628649148944), 0, 0, 0);
            poly96.AddVertexAt(4, new Point2d(scale1 * 0.743639238659363, scale1 * 0.246805178275643), 0, 0, 0);
            poly96.AddVertexAt(5, new Point2d(scale1 * 0.743873016976452, scale1 * 0.244981707402343), 0, 0, 0);
            poly96.AddVertexAt(6, new Point2d(scale1 * 0.744106795293542, scale1 * 0.242737435558282), 0, 0, 0);
            poly96.AddVertexAt(7, new Point2d(scale1 * 0.74597702183026, scale1 * 0.240773697694727), 0, 0, 0);
            poly96.AddVertexAt(8, new Point2d(scale1 * 0.744574351927722, scale1 * 0.237407289928637), 0, 0, 0);
            poly96.AddVertexAt(9, new Point2d(scale1 * 0.739431228951748, scale1 * 0.235583819055336), 0, 0, 0);
            poly96.AddVertexAt(10, new Point2d(scale1 * 0.729378761316889, scale1 * 0.235303285074829), 0, 0, 0);
            poly96.AddVertexAt(11, new Point2d(scale1 * 0.724469416658004, scale1 * 0.236144887016352), 0, 0, 0);
            poly96.AddVertexAt(12, new Point2d(scale1 * 0.720962741901658, scale1 * 0.237828090899398), 0, 0, 0);
            poly96.AddVertexAt(13, new Point2d(scale1 * 0.717689845462402, scale1 * 0.241054231675236), 0, 0, 0);
            poly96.AddVertexAt(14, new Point2d(scale1 * 0.715819618925684, scale1 * 0.244420639441328), 0, 0, 0);
            poly96.Closed = true;
            poly96.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly96.Layer = "0";
            poly96.Color = color_GP;
            poly96.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly96);
            Polyline poly97 = new Polyline();
            poly97.AddVertexAt(0, new Point2d(scale1 * 0.660285741001707, scale1 * 0.3), 0, 0, 0);
            poly97.AddVertexAt(1, new Point2d(scale1 * 0.661640169075273, scale1 * 0.297968357889651), 0, 0, 0);
            poly97.AddVertexAt(2, new Point2d(scale1 * 0.664679287197439, scale1 * 0.297687823909144), 0, 0, 0);
            poly97.AddVertexAt(3, new Point2d(scale1 * 0.668185961953786, scale1 * 0.29937102779219), 0, 0, 0);
            poly97.AddVertexAt(4, new Point2d(scale1 * 0.668185961953786, scale1 * 0.3), 0, 0, 0);
            poly97.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly97.Layer = "0";
            poly97.Color = color_GP;
            poly97.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly97);
            Polyline poly98 = new Polyline();
            poly98.AddVertexAt(0, new Point2d(scale1 * 0.645654647790791, scale1 * 0.3), 0, 0, 0);
            poly98.AddVertexAt(1, new Point2d(scale1 * 0.644574351927722, scale1 * 0.297407289928637), 0, 0, 0);
            poly98.AddVertexAt(2, new Point2d(scale1 * 0.639431228951748, scale1 * 0.295583819055336), 0, 0, 0);
            poly98.AddVertexAt(3, new Point2d(scale1 * 0.629378761316889, scale1 * 0.295303285074829), 0, 0, 0);
            poly98.AddVertexAt(4, new Point2d(scale1 * 0.624469416658004, scale1 * 0.296144887016352), 0, 0, 0);
            poly98.AddVertexAt(5, new Point2d(scale1 * 0.620962741901658, scale1 * 0.297828090899398), 0, 0, 0);
            poly98.AddVertexAt(6, new Point2d(scale1 * 0.618759355857569, scale1 * 0.3), 0, 0, 0);
            poly98.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly98.Layer = "0";
            poly98.Color = color_GP;
            poly98.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly98);
            Polyline poly99 = new Polyline();
            poly99.AddVertexAt(0, new Point2d(scale1 * 0.658367272636016, scale1 * 0.282994061761134), 0, 0, 0);
            poly99.AddVertexAt(1, new Point2d(scale1 * 0.656497046099298, scale1 * 0.282853794770882), 0, 0, 0);
            poly99.AddVertexAt(2, new Point2d(scale1 * 0.654159262928401, scale1 * 0.28355512972215), 0, 0, 0);
            poly99.AddVertexAt(3, new Point2d(scale1 * 0.652289036391683, scale1 * 0.285518867585705), 0, 0, 0);
            poly99.AddVertexAt(4, new Point2d(scale1 * 0.653925484611311, scale1 * 0.287061804478495), 0, 0, 0);
            poly99.AddVertexAt(5, new Point2d(scale1 * 0.658367272636016, scale1 * 0.286220202536974), 0, 0, 0);
            poly99.AddVertexAt(6, new Point2d(scale1 * 0.659536164221465, scale1 * 0.284817532634435), 0, 0, 0);
            poly99.Closed = true;
            poly99.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly99.Layer = "0";
            poly99.Color = color_GP;
            poly99.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly99);
            Polyline poly100 = new Polyline();
            poly100.AddVertexAt(0, new Point2d(scale1 * 0.667952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly100.AddVertexAt(1, new Point2d(scale1 * 0.666081957099978, scale1 * 0.282012192829359), 0, 0, 0);
            poly100.AddVertexAt(2, new Point2d(scale1 * 0.665848178782888, scale1 * 0.280048454965803), 0, 0, 0);
            poly100.AddVertexAt(3, new Point2d(scale1 * 0.668419740270875, scale1 * 0.278505518073012), 0, 0, 0);
            poly100.AddVertexAt(4, new Point2d(scale1 * 0.671692636710131, scale1 * 0.278926319043773), 0, 0, 0);
            poly100.AddVertexAt(5, new Point2d(scale1 * 0.672393971661401, scale1 * 0.281030323897581), 0, 0, 0);
            poly100.AddVertexAt(6, new Point2d(scale1 * 0.671926415027221, scale1 * 0.282713527780627), 0, 0, 0);
            poly100.AddVertexAt(7, new Point2d(scale1 * 0.667952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly100.Closed = true;
            poly100.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly100.Layer = "0";
            poly100.Color = color_GP;
            poly100.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly100);
            Polyline poly101 = new Polyline();
            poly101.AddVertexAt(0, new Point2d(scale1 * 0.616577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly101.AddVertexAt(1, new Point2d(scale1 * 0.614239328956148, scale1 * 0.285251085149766), 0, 0, 0);
            poly101.AddVertexAt(2, new Point2d(scale1 * 0.615641998858686, scale1 * 0.283427614276465), 0, 0, 0);
            poly101.AddVertexAt(3, new Point2d(scale1 * 0.619616230249212, scale1 * 0.282726279325197), 0, 0, 0);
            poly101.AddVertexAt(4, new Point2d(scale1 * 0.622655348371378, scale1 * 0.283988682237481), 0, 0, 0);
            poly101.AddVertexAt(5, new Point2d(scale1 * 0.622655348371378, scale1 * 0.286653755052303), 0, 0, 0);
            poly101.AddVertexAt(6, new Point2d(scale1 * 0.616577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly101.Closed = true;
            poly101.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly101.Layer = "0";
            poly101.Color = color_GP;
            poly101.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly101);
            Polyline poly102 = new Polyline();
            poly102.AddVertexAt(0, new Point2d(scale1 * 0.611080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly102.AddVertexAt(1, new Point2d(scale1 * 0.609476961807309, scale1 * 0.263254670769046), 0, 0, 0);
            poly102.AddVertexAt(2, new Point2d(scale1 * 0.609710740124399, scale1 * 0.259607729022446), 0, 0, 0);
            poly102.AddVertexAt(3, new Point2d(scale1 * 0.610645853392758, scale1 * 0.256802389217368), 0, 0, 0);
            poly102.AddVertexAt(4, new Point2d(scale1 * 0.614853863100373, scale1 * 0.255960787275844), 0, 0, 0);
            poly102.AddVertexAt(5, new Point2d(scale1 * 0.619061872807989, scale1 * 0.256521855236861), 0, 0, 0);
            poly102.AddVertexAt(6, new Point2d(scale1 * 0.622568547564335, scale1 * 0.258625860090668), 0, 0, 0);
            poly102.AddVertexAt(7, new Point2d(scale1 * 0.619061872807989, scale1 * 0.263535204749552), 0, 0, 0);
            poly102.AddVertexAt(8, new Point2d(scale1 * 0.611080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly102.Closed = true;
            poly102.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly102.Layer = "0";
            poly102.Color = color_GP;
            poly102.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly102);
            Polyline poly103 = new Polyline();
            poly103.AddVertexAt(0, new Point2d(scale1 * 0.687975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly103.AddVertexAt(1, new Point2d(scale1 * 0.684702548523924, scale1 * 0.288936281530065), 0, 0, 0);
            poly103.AddVertexAt(2, new Point2d(scale1 * 0.682598543670117, scale1 * 0.283746402890672), 0, 0, 0);
            poly103.AddVertexAt(3, new Point2d(scale1 * 0.682598543670117, scale1 * 0.281221597066103), 0, 0, 0);
            poly103.AddVertexAt(4, new Point2d(scale1 * 0.681897208718848, scale1 * 0.275751184446203), 0, 0, 0);
            poly103.AddVertexAt(5, new Point2d(scale1 * 0.685403883475194, scale1 * 0.273366645611888), 0, 0, 0);
            poly103.AddVertexAt(6, new Point2d(scale1 * 0.690313228134079, scale1 * 0.272525043670365), 0, 0, 0);
            poly103.AddVertexAt(7, new Point2d(scale1 * 0.701768365671475, scale1 * 0.272945844641126), 0, 0, 0);
            poly103.AddVertexAt(8, new Point2d(scale1 * 0.704807483793642, scale1 * 0.274488781533918), 0, 0, 0);
            poly103.AddVertexAt(9, new Point2d(scale1 * 0.705742597062001, scale1 * 0.278275990270772), 0, 0, 0);
            poly103.AddVertexAt(10, new Point2d(scale1 * 0.702703478939834, scale1 * 0.28037999512458), 0, 0, 0);
            poly103.AddVertexAt(11, new Point2d(scale1 * 0.699898139134758, scale1 * 0.282764533958894), 0, 0, 0);
            poly103.AddVertexAt(12, new Point2d(scale1 * 0.69756035596386, scale1 * 0.286271208715241), 0, 0, 0);
            poly103.AddVertexAt(13, new Point2d(scale1 * 0.695690129427142, scale1 * 0.288796014539811), 0, 0, 0);
            poly103.AddVertexAt(14, new Point2d(scale1 * 0.691482119719527, scale1 * 0.289777883471588), 0, 0, 0);
            poly103.AddVertexAt(15, new Point2d(scale1 * 0.687975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly103.Closed = true;
            poly103.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly103.Layer = "0";
            poly103.Color = color_GP;
            poly103.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly103);
            Polyline poly104 = new Polyline();
            poly104.AddVertexAt(0, new Point2d(scale1 * 0.658367272636016, scale1 * 0.274437775355651), 0, 0, 0);
            poly104.AddVertexAt(1, new Point2d(scale1 * 0.666081957099978, scale1 * 0.272754571472604), 0, 0, 0);
            poly104.AddVertexAt(2, new Point2d(scale1 * 0.669354853539234, scale1 * 0.268125760794228), 0, 0, 0);
            poly104.AddVertexAt(3, new Point2d(scale1 * 0.666081957099978, scale1 * 0.263637217106105), 0, 0, 0);
            poly104.AddVertexAt(4, new Point2d(scale1 * 0.658367272636016, scale1 * 0.261954013223059), 0, 0, 0);
            poly104.AddVertexAt(5, new Point2d(scale1 * 0.649717474903696, scale1 * 0.261392945262043), 0, 0, 0);
            poly104.AddVertexAt(6, new Point2d(scale1 * 0.645041908561901, scale1 * 0.26237481419382), 0, 0, 0);
            poly104.AddVertexAt(7, new Point2d(scale1 * 0.641301455488465, scale1 * 0.26489962001839), 0, 0, 0);
            poly104.AddVertexAt(8, new Point2d(scale1 * 0.630781431219427, scale1 * 0.268967362735751), 0, 0, 0);
            poly104.AddVertexAt(9, new Point2d(scale1 * 0.625170751609273, scale1 * 0.272894838462859), 0, 0, 0);
            poly104.AddVertexAt(10, new Point2d(scale1 * 0.626573421511812, scale1 * 0.27471830933616), 0, 0, 0);
            poly104.AddVertexAt(11, new Point2d(scale1 * 0.634989440927042, scale1 * 0.278365251082757), 0, 0, 0);
            poly104.AddVertexAt(12, new Point2d(scale1 * 0.641535233805555, scale1 * 0.279066586034028), 0, 0, 0);
            poly104.AddVertexAt(13, new Point2d(scale1 * 0.64574324351317, scale1 * 0.277804183121743), 0, 0, 0);
            poly104.AddVertexAt(14, new Point2d(scale1 * 0.651353923123324, scale1 * 0.276541780209458), 0, 0, 0);
            poly104.AddVertexAt(15, new Point2d(scale1 * 0.655795711148029, scale1 * 0.275840445258188), 0, 0, 0);
            poly104.Closed = true;
            poly104.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly104.Layer = "0";
            poly104.Color = color_GP;
            poly104.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly104);
            Polyline poly105 = new Polyline();
            poly105.AddVertexAt(0, new Point2d(scale1 * 0.678898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly105.AddVertexAt(1, new Point2d(scale1 * 0.672741026607184, scale1 * 0.262017770945895), 0, 0, 0);
            poly105.AddVertexAt(2, new Point2d(scale1 * 0.66743251686181, scale1 * 0.259722492923566), 0, 0, 0);
            poly105.AddVertexAt(3, new Point2d(scale1 * 0.666370814912736, scale1 * 0.256662122227115), 0, 0, 0);
            poly105.AddVertexAt(4, new Point2d(scale1 * 0.67167932465811, scale1 * 0.255769514107321), 0, 0, 0);
            poly105.AddVertexAt(5, new Point2d(scale1 * 0.678261876742373, scale1 * 0.2567896376728), 0, 0, 0);
            poly105.AddVertexAt(6, new Point2d(scale1 * 0.682084003759042, scale1 * 0.259467462032192), 0, 0, 0);
            poly105.AddVertexAt(7, new Point2d(scale1 * 0.678898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly105.Closed = true;
            poly105.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly105.Layer = "0";
            poly105.Color = color_GP;
            poly105.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly105);
            Polyline poly106 = new Polyline();
            poly106.AddVertexAt(0, new Point2d(scale1 * 0.696473138321644, scale1 * 0.248628649148944), 0, 0, 0);
            poly106.AddVertexAt(1, new Point2d(scale1 * 0.696706916638733, scale1 * 0.243298503519296), 0, 0, 0);
            poly106.AddVertexAt(2, new Point2d(scale1 * 0.695070468419106, scale1 * 0.242036100607014), 0, 0, 0);
            poly106.AddVertexAt(3, new Point2d(scale1 * 0.69086245871149, scale1 * 0.240633430704475), 0, 0, 0);
            poly106.AddVertexAt(4, new Point2d(scale1 * 0.686888227320964, scale1 * 0.240773697694727), 0, 0, 0);
            poly106.AddVertexAt(5, new Point2d(scale1 * 0.683849109198798, scale1 * 0.242036100607014), 0, 0, 0);
            poly106.AddVertexAt(6, new Point2d(scale1 * 0.680576212759542, scale1 * 0.243298503519296), 0, 0, 0);
            poly106.AddVertexAt(7, new Point2d(scale1 * 0.677537094637375, scale1 * 0.244560906431583), 0, 0, 0);
            poly106.AddVertexAt(8, new Point2d(scale1 * 0.675666868100657, scale1 * 0.246945445265898), 0, 0, 0);
            poly106.AddVertexAt(9, new Point2d(scale1 * 0.675666868100657, scale1 * 0.248628649148944), 0, 0, 0);
            poly106.AddVertexAt(10, new Point2d(scale1 * 0.679173542857003, scale1 * 0.252275590895543), 0, 0, 0);
            poly106.AddVertexAt(11, new Point2d(scale1 * 0.685953114052606, scale1 * 0.253818527788335), 0, 0, 0);
            poly106.AddVertexAt(12, new Point2d(scale1 * 0.692732685248208, scale1 * 0.252275590895543), 0, 0, 0);
            poly106.Closed = true;
            poly106.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly106.Layer = "0";
            poly106.Color = color_GP;
            poly106.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly106);
            Polyline poly107 = new Polyline();
            poly107.AddVertexAt(0, new Point2d(scale1 * 0.662107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly107.AddVertexAt(1, new Point2d(scale1 * 0.659302385904375, scale1 * 0.241895833616759), 0, 0, 0);
            poly107.AddVertexAt(2, new Point2d(scale1 * 0.660237499172734, scale1 * 0.240072362743458), 0, 0, 0);
            poly107.AddVertexAt(3, new Point2d(scale1 * 0.661640169075273, scale1 * 0.237968357889651), 0, 0, 0);
            poly107.AddVertexAt(4, new Point2d(scale1 * 0.664679287197439, scale1 * 0.237687823909144), 0, 0, 0);
            poly107.AddVertexAt(5, new Point2d(scale1 * 0.668185961953786, scale1 * 0.23937102779219), 0, 0, 0);
            poly107.AddVertexAt(6, new Point2d(scale1 * 0.668185961953786, scale1 * 0.241615299636252), 0, 0, 0);
            poly107.AddVertexAt(7, new Point2d(scale1 * 0.665848178782888, scale1 * 0.244140105460821), 0, 0, 0);
            poly107.AddVertexAt(8, new Point2d(scale1 * 0.662107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly107.Closed = true;
            poly107.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly107.Layer = "0";
            poly107.Color = color_GP;
            poly107.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly107);
            Polyline poly108 = new Polyline();
            poly108.AddVertexAt(0, new Point2d(scale1 * 0.616520953876953, scale1 * 0.248628649148944), 0, 0, 0);
            poly108.AddVertexAt(1, new Point2d(scale1 * 0.623066746755466, scale1 * 0.250452120022245), 0, 0, 0);
            poly108.AddVertexAt(2, new Point2d(scale1 * 0.635924554195402, scale1 * 0.249750785070974), 0, 0, 0);
            poly108.AddVertexAt(3, new Point2d(scale1 * 0.642470347073914, scale1 * 0.248628649148944), 0, 0, 0);
            poly108.AddVertexAt(4, new Point2d(scale1 * 0.643639238659363, scale1 * 0.246805178275643), 0, 0, 0);
            poly108.AddVertexAt(5, new Point2d(scale1 * 0.643873016976453, scale1 * 0.244981707402343), 0, 0, 0);
            poly108.AddVertexAt(6, new Point2d(scale1 * 0.644106795293542, scale1 * 0.242737435558282), 0, 0, 0);
            poly108.AddVertexAt(7, new Point2d(scale1 * 0.64597702183026, scale1 * 0.240773697694727), 0, 0, 0);
            poly108.AddVertexAt(8, new Point2d(scale1 * 0.644574351927722, scale1 * 0.237407289928637), 0, 0, 0);
            poly108.AddVertexAt(9, new Point2d(scale1 * 0.639431228951748, scale1 * 0.235583819055336), 0, 0, 0);
            poly108.AddVertexAt(10, new Point2d(scale1 * 0.629378761316889, scale1 * 0.235303285074829), 0, 0, 0);
            poly108.AddVertexAt(11, new Point2d(scale1 * 0.624469416658004, scale1 * 0.236144887016352), 0, 0, 0);
            poly108.AddVertexAt(12, new Point2d(scale1 * 0.620962741901658, scale1 * 0.237828090899398), 0, 0, 0);
            poly108.AddVertexAt(13, new Point2d(scale1 * 0.617689845462402, scale1 * 0.241054231675236), 0, 0, 0);
            poly108.AddVertexAt(14, new Point2d(scale1 * 0.615819618925684, scale1 * 0.244420639441328), 0, 0, 0);
            poly108.Closed = true;
            poly108.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly108.Layer = "0";
            poly108.Color = color_GP;
            poly108.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly108);
            Polyline poly109 = new Polyline();
            poly109.AddVertexAt(0, new Point2d(scale1 * 0.560285741001707, scale1 * 0.3), 0, 0, 0);
            poly109.AddVertexAt(1, new Point2d(scale1 * 0.561640169075273, scale1 * 0.297968357889651), 0, 0, 0);
            poly109.AddVertexAt(2, new Point2d(scale1 * 0.564679287197439, scale1 * 0.297687823909144), 0, 0, 0);
            poly109.AddVertexAt(3, new Point2d(scale1 * 0.568185961953785, scale1 * 0.29937102779219), 0, 0, 0);
            poly109.AddVertexAt(4, new Point2d(scale1 * 0.568185961953785, scale1 * 0.3), 0, 0, 0);
            poly109.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly109.Layer = "0";
            poly109.Color = color_GP;
            poly109.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly109);
            Polyline poly110 = new Polyline();
            poly110.AddVertexAt(0, new Point2d(scale1 * 0.54565464779079, scale1 * 0.3), 0, 0, 0);
            poly110.AddVertexAt(1, new Point2d(scale1 * 0.544574351927722, scale1 * 0.297407289928637), 0, 0, 0);
            poly110.AddVertexAt(2, new Point2d(scale1 * 0.539431228951748, scale1 * 0.295583819055336), 0, 0, 0);
            poly110.AddVertexAt(3, new Point2d(scale1 * 0.529378761316889, scale1 * 0.295303285074829), 0, 0, 0);
            poly110.AddVertexAt(4, new Point2d(scale1 * 0.524469416658004, scale1 * 0.296144887016352), 0, 0, 0);
            poly110.AddVertexAt(5, new Point2d(scale1 * 0.520962741901658, scale1 * 0.297828090899398), 0, 0, 0);
            poly110.AddVertexAt(6, new Point2d(scale1 * 0.518759355857569, scale1 * 0.3), 0, 0, 0);
            poly110.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly110.Layer = "0";
            poly110.Color = color_GP;
            poly110.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly110);
            Polyline poly111 = new Polyline();
            poly111.AddVertexAt(0, new Point2d(scale1 * 0.558367272636016, scale1 * 0.282994061761134), 0, 0, 0);
            poly111.AddVertexAt(1, new Point2d(scale1 * 0.556497046099298, scale1 * 0.282853794770882), 0, 0, 0);
            poly111.AddVertexAt(2, new Point2d(scale1 * 0.554159262928401, scale1 * 0.28355512972215), 0, 0, 0);
            poly111.AddVertexAt(3, new Point2d(scale1 * 0.552289036391683, scale1 * 0.285518867585705), 0, 0, 0);
            poly111.AddVertexAt(4, new Point2d(scale1 * 0.553925484611311, scale1 * 0.287061804478495), 0, 0, 0);
            poly111.AddVertexAt(5, new Point2d(scale1 * 0.558367272636016, scale1 * 0.286220202536974), 0, 0, 0);
            poly111.AddVertexAt(6, new Point2d(scale1 * 0.559536164221465, scale1 * 0.284817532634435), 0, 0, 0);
            poly111.Closed = true;
            poly111.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly111.Layer = "0";
            poly111.Color = color_GP;
            poly111.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly111);
            Polyline poly112 = new Polyline();
            poly112.AddVertexAt(0, new Point2d(scale1 * 0.567952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly112.AddVertexAt(1, new Point2d(scale1 * 0.566081957099978, scale1 * 0.282012192829359), 0, 0, 0);
            poly112.AddVertexAt(2, new Point2d(scale1 * 0.565848178782888, scale1 * 0.280048454965803), 0, 0, 0);
            poly112.AddVertexAt(3, new Point2d(scale1 * 0.568419740270875, scale1 * 0.278505518073012), 0, 0, 0);
            poly112.AddVertexAt(4, new Point2d(scale1 * 0.571692636710131, scale1 * 0.278926319043773), 0, 0, 0);
            poly112.AddVertexAt(5, new Point2d(scale1 * 0.572393971661401, scale1 * 0.281030323897581), 0, 0, 0);
            poly112.AddVertexAt(6, new Point2d(scale1 * 0.571926415027221, scale1 * 0.282713527780627), 0, 0, 0);
            poly112.AddVertexAt(7, new Point2d(scale1 * 0.567952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly112.Closed = true;
            poly112.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly112.Layer = "0";
            poly112.Color = color_GP;
            poly112.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly112);
            Polyline poly113 = new Polyline();
            poly113.AddVertexAt(0, new Point2d(scale1 * 0.516577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly113.AddVertexAt(1, new Point2d(scale1 * 0.514239328956148, scale1 * 0.285251085149766), 0, 0, 0);
            poly113.AddVertexAt(2, new Point2d(scale1 * 0.515641998858686, scale1 * 0.283427614276465), 0, 0, 0);
            poly113.AddVertexAt(3, new Point2d(scale1 * 0.519616230249212, scale1 * 0.282726279325197), 0, 0, 0);
            poly113.AddVertexAt(4, new Point2d(scale1 * 0.522655348371378, scale1 * 0.283988682237481), 0, 0, 0);
            poly113.AddVertexAt(5, new Point2d(scale1 * 0.522655348371378, scale1 * 0.286653755052303), 0, 0, 0);
            poly113.AddVertexAt(6, new Point2d(scale1 * 0.516577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly113.Closed = true;
            poly113.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly113.Layer = "0";
            poly113.Color = color_GP;
            poly113.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly113);
            Polyline poly114 = new Polyline();
            poly114.AddVertexAt(0, new Point2d(scale1 * 0.511080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly114.AddVertexAt(1, new Point2d(scale1 * 0.509476961807309, scale1 * 0.263254670769046), 0, 0, 0);
            poly114.AddVertexAt(2, new Point2d(scale1 * 0.509710740124399, scale1 * 0.259607729022446), 0, 0, 0);
            poly114.AddVertexAt(3, new Point2d(scale1 * 0.510645853392758, scale1 * 0.256802389217368), 0, 0, 0);
            poly114.AddVertexAt(4, new Point2d(scale1 * 0.514853863100373, scale1 * 0.255960787275844), 0, 0, 0);
            poly114.AddVertexAt(5, new Point2d(scale1 * 0.519061872807989, scale1 * 0.256521855236861), 0, 0, 0);
            poly114.AddVertexAt(6, new Point2d(scale1 * 0.522568547564335, scale1 * 0.258625860090668), 0, 0, 0);
            poly114.AddVertexAt(7, new Point2d(scale1 * 0.519061872807989, scale1 * 0.263535204749552), 0, 0, 0);
            poly114.AddVertexAt(8, new Point2d(scale1 * 0.511080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly114.Closed = true;
            poly114.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly114.Layer = "0";
            poly114.Color = color_GP;
            poly114.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly114);
            Polyline poly115 = new Polyline();
            poly115.AddVertexAt(0, new Point2d(scale1 * 0.587975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly115.AddVertexAt(1, new Point2d(scale1 * 0.584702548523924, scale1 * 0.288936281530065), 0, 0, 0);
            poly115.AddVertexAt(2, new Point2d(scale1 * 0.582598543670117, scale1 * 0.283746402890672), 0, 0, 0);
            poly115.AddVertexAt(3, new Point2d(scale1 * 0.582598543670117, scale1 * 0.281221597066103), 0, 0, 0);
            poly115.AddVertexAt(4, new Point2d(scale1 * 0.581897208718847, scale1 * 0.275751184446203), 0, 0, 0);
            poly115.AddVertexAt(5, new Point2d(scale1 * 0.585403883475194, scale1 * 0.273366645611888), 0, 0, 0);
            poly115.AddVertexAt(6, new Point2d(scale1 * 0.590313228134079, scale1 * 0.272525043670365), 0, 0, 0);
            poly115.AddVertexAt(7, new Point2d(scale1 * 0.601768365671475, scale1 * 0.272945844641126), 0, 0, 0);
            poly115.AddVertexAt(8, new Point2d(scale1 * 0.604807483793642, scale1 * 0.274488781533918), 0, 0, 0);
            poly115.AddVertexAt(9, new Point2d(scale1 * 0.605742597062001, scale1 * 0.278275990270772), 0, 0, 0);
            poly115.AddVertexAt(10, new Point2d(scale1 * 0.602703478939834, scale1 * 0.28037999512458), 0, 0, 0);
            poly115.AddVertexAt(11, new Point2d(scale1 * 0.599898139134758, scale1 * 0.282764533958894), 0, 0, 0);
            poly115.AddVertexAt(12, new Point2d(scale1 * 0.59756035596386, scale1 * 0.286271208715241), 0, 0, 0);
            poly115.AddVertexAt(13, new Point2d(scale1 * 0.595690129427142, scale1 * 0.288796014539811), 0, 0, 0);
            poly115.AddVertexAt(14, new Point2d(scale1 * 0.591482119719527, scale1 * 0.289777883471588), 0, 0, 0);
            poly115.AddVertexAt(15, new Point2d(scale1 * 0.587975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly115.Closed = true;
            poly115.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly115.Layer = "0";
            poly115.Color = color_GP;
            poly115.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly115);
            Polyline poly116 = new Polyline();
            poly116.AddVertexAt(0, new Point2d(scale1 * 0.558367272636016, scale1 * 0.274437775355651), 0, 0, 0);
            poly116.AddVertexAt(1, new Point2d(scale1 * 0.566081957099978, scale1 * 0.272754571472604), 0, 0, 0);
            poly116.AddVertexAt(2, new Point2d(scale1 * 0.569354853539234, scale1 * 0.268125760794228), 0, 0, 0);
            poly116.AddVertexAt(3, new Point2d(scale1 * 0.566081957099978, scale1 * 0.263637217106105), 0, 0, 0);
            poly116.AddVertexAt(4, new Point2d(scale1 * 0.558367272636016, scale1 * 0.261954013223059), 0, 0, 0);
            poly116.AddVertexAt(5, new Point2d(scale1 * 0.549717474903696, scale1 * 0.261392945262043), 0, 0, 0);
            poly116.AddVertexAt(6, new Point2d(scale1 * 0.545041908561901, scale1 * 0.26237481419382), 0, 0, 0);
            poly116.AddVertexAt(7, new Point2d(scale1 * 0.541301455488465, scale1 * 0.26489962001839), 0, 0, 0);
            poly116.AddVertexAt(8, new Point2d(scale1 * 0.530781431219427, scale1 * 0.268967362735751), 0, 0, 0);
            poly116.AddVertexAt(9, new Point2d(scale1 * 0.525170751609273, scale1 * 0.272894838462859), 0, 0, 0);
            poly116.AddVertexAt(10, new Point2d(scale1 * 0.526573421511812, scale1 * 0.27471830933616), 0, 0, 0);
            poly116.AddVertexAt(11, new Point2d(scale1 * 0.534989440927042, scale1 * 0.278365251082757), 0, 0, 0);
            poly116.AddVertexAt(12, new Point2d(scale1 * 0.541535233805555, scale1 * 0.279066586034028), 0, 0, 0);
            poly116.AddVertexAt(13, new Point2d(scale1 * 0.54574324351317, scale1 * 0.277804183121743), 0, 0, 0);
            poly116.AddVertexAt(14, new Point2d(scale1 * 0.551353923123324, scale1 * 0.276541780209458), 0, 0, 0);
            poly116.AddVertexAt(15, new Point2d(scale1 * 0.555795711148029, scale1 * 0.275840445258188), 0, 0, 0);
            poly116.Closed = true;
            poly116.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly116.Layer = "0";
            poly116.Color = color_GP;
            poly116.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly116);
            Polyline poly117 = new Polyline();
            poly117.AddVertexAt(0, new Point2d(scale1 * 0.578898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly117.AddVertexAt(1, new Point2d(scale1 * 0.572741026607184, scale1 * 0.262017770945895), 0, 0, 0);
            poly117.AddVertexAt(2, new Point2d(scale1 * 0.56743251686181, scale1 * 0.259722492923566), 0, 0, 0);
            poly117.AddVertexAt(3, new Point2d(scale1 * 0.566370814912736, scale1 * 0.256662122227115), 0, 0, 0);
            poly117.AddVertexAt(4, new Point2d(scale1 * 0.57167932465811, scale1 * 0.255769514107321), 0, 0, 0);
            poly117.AddVertexAt(5, new Point2d(scale1 * 0.578261876742373, scale1 * 0.2567896376728), 0, 0, 0);
            poly117.AddVertexAt(6, new Point2d(scale1 * 0.582084003759042, scale1 * 0.259467462032192), 0, 0, 0);
            poly117.AddVertexAt(7, new Point2d(scale1 * 0.578898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly117.Closed = true;
            poly117.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly117.Layer = "0";
            poly117.Color = color_GP;
            poly117.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly117);
            Polyline poly118 = new Polyline();
            poly118.AddVertexAt(0, new Point2d(scale1 * 0.596473138321644, scale1 * 0.248628649148944), 0, 0, 0);
            poly118.AddVertexAt(1, new Point2d(scale1 * 0.596706916638733, scale1 * 0.243298503519296), 0, 0, 0);
            poly118.AddVertexAt(2, new Point2d(scale1 * 0.595070468419106, scale1 * 0.242036100607014), 0, 0, 0);
            poly118.AddVertexAt(3, new Point2d(scale1 * 0.59086245871149, scale1 * 0.240633430704475), 0, 0, 0);
            poly118.AddVertexAt(4, new Point2d(scale1 * 0.586888227320964, scale1 * 0.240773697694727), 0, 0, 0);
            poly118.AddVertexAt(5, new Point2d(scale1 * 0.583849109198798, scale1 * 0.242036100607014), 0, 0, 0);
            poly118.AddVertexAt(6, new Point2d(scale1 * 0.580576212759542, scale1 * 0.243298503519296), 0, 0, 0);
            poly118.AddVertexAt(7, new Point2d(scale1 * 0.577537094637375, scale1 * 0.244560906431583), 0, 0, 0);
            poly118.AddVertexAt(8, new Point2d(scale1 * 0.575666868100657, scale1 * 0.246945445265898), 0, 0, 0);
            poly118.AddVertexAt(9, new Point2d(scale1 * 0.575666868100657, scale1 * 0.248628649148944), 0, 0, 0);
            poly118.AddVertexAt(10, new Point2d(scale1 * 0.579173542857003, scale1 * 0.252275590895543), 0, 0, 0);
            poly118.AddVertexAt(11, new Point2d(scale1 * 0.585953114052606, scale1 * 0.253818527788335), 0, 0, 0);
            poly118.AddVertexAt(12, new Point2d(scale1 * 0.592732685248208, scale1 * 0.252275590895543), 0, 0, 0);
            poly118.Closed = true;
            poly118.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly118.Layer = "0";
            poly118.Color = color_GP;
            poly118.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly118);
            Polyline poly119 = new Polyline();
            poly119.AddVertexAt(0, new Point2d(scale1 * 0.562107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly119.AddVertexAt(1, new Point2d(scale1 * 0.559302385904375, scale1 * 0.241895833616759), 0, 0, 0);
            poly119.AddVertexAt(2, new Point2d(scale1 * 0.560237499172734, scale1 * 0.240072362743458), 0, 0, 0);
            poly119.AddVertexAt(3, new Point2d(scale1 * 0.561640169075273, scale1 * 0.237968357889651), 0, 0, 0);
            poly119.AddVertexAt(4, new Point2d(scale1 * 0.564679287197439, scale1 * 0.237687823909144), 0, 0, 0);
            poly119.AddVertexAt(5, new Point2d(scale1 * 0.568185961953785, scale1 * 0.23937102779219), 0, 0, 0);
            poly119.AddVertexAt(6, new Point2d(scale1 * 0.568185961953785, scale1 * 0.241615299636252), 0, 0, 0);
            poly119.AddVertexAt(7, new Point2d(scale1 * 0.565848178782888, scale1 * 0.244140105460821), 0, 0, 0);
            poly119.AddVertexAt(8, new Point2d(scale1 * 0.562107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly119.Closed = true;
            poly119.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly119.Layer = "0";
            poly119.Color = color_GP;
            poly119.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly119);
            Polyline poly120 = new Polyline();
            poly120.AddVertexAt(0, new Point2d(scale1 * 0.516520953876953, scale1 * 0.248628649148944), 0, 0, 0);
            poly120.AddVertexAt(1, new Point2d(scale1 * 0.523066746755466, scale1 * 0.250452120022245), 0, 0, 0);
            poly120.AddVertexAt(2, new Point2d(scale1 * 0.535924554195401, scale1 * 0.249750785070974), 0, 0, 0);
            poly120.AddVertexAt(3, new Point2d(scale1 * 0.542470347073914, scale1 * 0.248628649148944), 0, 0, 0);
            poly120.AddVertexAt(4, new Point2d(scale1 * 0.543639238659362, scale1 * 0.246805178275643), 0, 0, 0);
            poly120.AddVertexAt(5, new Point2d(scale1 * 0.543873016976453, scale1 * 0.244981707402343), 0, 0, 0);
            poly120.AddVertexAt(6, new Point2d(scale1 * 0.544106795293542, scale1 * 0.242737435558282), 0, 0, 0);
            poly120.AddVertexAt(7, new Point2d(scale1 * 0.54597702183026, scale1 * 0.240773697694727), 0, 0, 0);
            poly120.AddVertexAt(8, new Point2d(scale1 * 0.544574351927722, scale1 * 0.237407289928637), 0, 0, 0);
            poly120.AddVertexAt(9, new Point2d(scale1 * 0.539431228951748, scale1 * 0.235583819055336), 0, 0, 0);
            poly120.AddVertexAt(10, new Point2d(scale1 * 0.529378761316889, scale1 * 0.235303285074829), 0, 0, 0);
            poly120.AddVertexAt(11, new Point2d(scale1 * 0.524469416658004, scale1 * 0.236144887016352), 0, 0, 0);
            poly120.AddVertexAt(12, new Point2d(scale1 * 0.520962741901658, scale1 * 0.237828090899398), 0, 0, 0);
            poly120.AddVertexAt(13, new Point2d(scale1 * 0.517689845462401, scale1 * 0.241054231675236), 0, 0, 0);
            poly120.AddVertexAt(14, new Point2d(scale1 * 0.515819618925684, scale1 * 0.244420639441328), 0, 0, 0);
            poly120.Closed = true;
            poly120.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly120.Layer = "0";
            poly120.Color = color_GP;
            poly120.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly120);
            Polyline poly121 = new Polyline();
            poly121.AddVertexAt(0, new Point2d(scale1 * 0.460285741001707, scale1 * 0.3), 0, 0, 0);
            poly121.AddVertexAt(1, new Point2d(scale1 * 0.461640169075273, scale1 * 0.297968357889651), 0, 0, 0);
            poly121.AddVertexAt(2, new Point2d(scale1 * 0.464679287197439, scale1 * 0.297687823909144), 0, 0, 0);
            poly121.AddVertexAt(3, new Point2d(scale1 * 0.468185961953785, scale1 * 0.29937102779219), 0, 0, 0);
            poly121.AddVertexAt(4, new Point2d(scale1 * 0.468185961953785, scale1 * 0.3), 0, 0, 0);
            poly121.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly121.Layer = "0";
            poly121.Color = color_GP;
            poly121.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly121);
            Polyline poly122 = new Polyline();
            poly122.AddVertexAt(0, new Point2d(scale1 * 0.44565464779079, scale1 * 0.3), 0, 0, 0);
            poly122.AddVertexAt(1, new Point2d(scale1 * 0.444574351927722, scale1 * 0.297407289928637), 0, 0, 0);
            poly122.AddVertexAt(2, new Point2d(scale1 * 0.439431228951747, scale1 * 0.295583819055336), 0, 0, 0);
            poly122.AddVertexAt(3, new Point2d(scale1 * 0.429378761316889, scale1 * 0.295303285074829), 0, 0, 0);
            poly122.AddVertexAt(4, new Point2d(scale1 * 0.424469416658004, scale1 * 0.296144887016352), 0, 0, 0);
            poly122.AddVertexAt(5, new Point2d(scale1 * 0.420962741901658, scale1 * 0.297828090899398), 0, 0, 0);
            poly122.AddVertexAt(6, new Point2d(scale1 * 0.418759355857569, scale1 * 0.3), 0, 0, 0);
            poly122.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly122.Layer = "0";
            poly122.Color = color_GP;
            poly122.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly122);
            Polyline poly123 = new Polyline();
            poly123.AddVertexAt(0, new Point2d(scale1 * 0.458367272636016, scale1 * 0.282994061761134), 0, 0, 0);
            poly123.AddVertexAt(1, new Point2d(scale1 * 0.456497046099298, scale1 * 0.282853794770882), 0, 0, 0);
            poly123.AddVertexAt(2, new Point2d(scale1 * 0.454159262928401, scale1 * 0.28355512972215), 0, 0, 0);
            poly123.AddVertexAt(3, new Point2d(scale1 * 0.452289036391683, scale1 * 0.285518867585705), 0, 0, 0);
            poly123.AddVertexAt(4, new Point2d(scale1 * 0.453925484611311, scale1 * 0.287061804478495), 0, 0, 0);
            poly123.AddVertexAt(5, new Point2d(scale1 * 0.458367272636016, scale1 * 0.286220202536974), 0, 0, 0);
            poly123.AddVertexAt(6, new Point2d(scale1 * 0.459536164221465, scale1 * 0.284817532634435), 0, 0, 0);
            poly123.Closed = true;
            poly123.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly123.Layer = "0";
            poly123.Color = color_GP;
            poly123.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly123);
            Polyline poly124 = new Polyline();
            poly124.AddVertexAt(0, new Point2d(scale1 * 0.467952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly124.AddVertexAt(1, new Point2d(scale1 * 0.466081957099978, scale1 * 0.282012192829359), 0, 0, 0);
            poly124.AddVertexAt(2, new Point2d(scale1 * 0.465848178782888, scale1 * 0.280048454965803), 0, 0, 0);
            poly124.AddVertexAt(3, new Point2d(scale1 * 0.468419740270875, scale1 * 0.278505518073012), 0, 0, 0);
            poly124.AddVertexAt(4, new Point2d(scale1 * 0.471692636710132, scale1 * 0.278926319043773), 0, 0, 0);
            poly124.AddVertexAt(5, new Point2d(scale1 * 0.472393971661401, scale1 * 0.281030323897581), 0, 0, 0);
            poly124.AddVertexAt(6, new Point2d(scale1 * 0.471926415027221, scale1 * 0.282713527780627), 0, 0, 0);
            poly124.AddVertexAt(7, new Point2d(scale1 * 0.467952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly124.Closed = true;
            poly124.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly124.Layer = "0";
            poly124.Color = color_GP;
            poly124.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly124);
            Polyline poly125 = new Polyline();
            poly125.AddVertexAt(0, new Point2d(scale1 * 0.416577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly125.AddVertexAt(1, new Point2d(scale1 * 0.414239328956148, scale1 * 0.285251085149766), 0, 0, 0);
            poly125.AddVertexAt(2, new Point2d(scale1 * 0.415641998858686, scale1 * 0.283427614276465), 0, 0, 0);
            poly125.AddVertexAt(3, new Point2d(scale1 * 0.419616230249211, scale1 * 0.282726279325197), 0, 0, 0);
            poly125.AddVertexAt(4, new Point2d(scale1 * 0.422655348371378, scale1 * 0.283988682237481), 0, 0, 0);
            poly125.AddVertexAt(5, new Point2d(scale1 * 0.422655348371378, scale1 * 0.286653755052303), 0, 0, 0);
            poly125.AddVertexAt(6, new Point2d(scale1 * 0.416577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly125.Closed = true;
            poly125.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly125.Layer = "0";
            poly125.Color = color_GP;
            poly125.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly125);
            Polyline poly126 = new Polyline();
            poly126.AddVertexAt(0, new Point2d(scale1 * 0.411080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly126.AddVertexAt(1, new Point2d(scale1 * 0.409476961807309, scale1 * 0.263254670769046), 0, 0, 0);
            poly126.AddVertexAt(2, new Point2d(scale1 * 0.409710740124399, scale1 * 0.259607729022446), 0, 0, 0);
            poly126.AddVertexAt(3, new Point2d(scale1 * 0.410645853392758, scale1 * 0.256802389217368), 0, 0, 0);
            poly126.AddVertexAt(4, new Point2d(scale1 * 0.414853863100374, scale1 * 0.255960787275844), 0, 0, 0);
            poly126.AddVertexAt(5, new Point2d(scale1 * 0.419061872807989, scale1 * 0.256521855236861), 0, 0, 0);
            poly126.AddVertexAt(6, new Point2d(scale1 * 0.422568547564335, scale1 * 0.258625860090668), 0, 0, 0);
            poly126.AddVertexAt(7, new Point2d(scale1 * 0.419061872807989, scale1 * 0.263535204749552), 0, 0, 0);
            poly126.AddVertexAt(8, new Point2d(scale1 * 0.411080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly126.Closed = true;
            poly126.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly126.Layer = "0";
            poly126.Color = color_GP;
            poly126.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly126);
            Polyline poly127 = new Polyline();
            poly127.AddVertexAt(0, new Point2d(scale1 * 0.487975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly127.AddVertexAt(1, new Point2d(scale1 * 0.484702548523924, scale1 * 0.288936281530065), 0, 0, 0);
            poly127.AddVertexAt(2, new Point2d(scale1 * 0.482598543670117, scale1 * 0.283746402890672), 0, 0, 0);
            poly127.AddVertexAt(3, new Point2d(scale1 * 0.482598543670117, scale1 * 0.281221597066103), 0, 0, 0);
            poly127.AddVertexAt(4, new Point2d(scale1 * 0.481897208718848, scale1 * 0.275751184446203), 0, 0, 0);
            poly127.AddVertexAt(5, new Point2d(scale1 * 0.485403883475193, scale1 * 0.273366645611888), 0, 0, 0);
            poly127.AddVertexAt(6, new Point2d(scale1 * 0.490313228134079, scale1 * 0.272525043670365), 0, 0, 0);
            poly127.AddVertexAt(7, new Point2d(scale1 * 0.501768365671475, scale1 * 0.272945844641126), 0, 0, 0);
            poly127.AddVertexAt(8, new Point2d(scale1 * 0.504807483793642, scale1 * 0.274488781533918), 0, 0, 0);
            poly127.AddVertexAt(9, new Point2d(scale1 * 0.505742597062001, scale1 * 0.278275990270772), 0, 0, 0);
            poly127.AddVertexAt(10, new Point2d(scale1 * 0.502703478939834, scale1 * 0.28037999512458), 0, 0, 0);
            poly127.AddVertexAt(11, new Point2d(scale1 * 0.499898139134758, scale1 * 0.282764533958894), 0, 0, 0);
            poly127.AddVertexAt(12, new Point2d(scale1 * 0.49756035596386, scale1 * 0.286271208715241), 0, 0, 0);
            poly127.AddVertexAt(13, new Point2d(scale1 * 0.495690129427142, scale1 * 0.288796014539811), 0, 0, 0);
            poly127.AddVertexAt(14, new Point2d(scale1 * 0.491482119719527, scale1 * 0.289777883471588), 0, 0, 0);
            poly127.AddVertexAt(15, new Point2d(scale1 * 0.487975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly127.Closed = true;
            poly127.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly127.Layer = "0";
            poly127.Color = color_GP;
            poly127.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly127);
            Polyline poly128 = new Polyline();
            poly128.AddVertexAt(0, new Point2d(scale1 * 0.458367272636016, scale1 * 0.274437775355651), 0, 0, 0);
            poly128.AddVertexAt(1, new Point2d(scale1 * 0.466081957099978, scale1 * 0.272754571472604), 0, 0, 0);
            poly128.AddVertexAt(2, new Point2d(scale1 * 0.469354853539234, scale1 * 0.268125760794228), 0, 0, 0);
            poly128.AddVertexAt(3, new Point2d(scale1 * 0.466081957099978, scale1 * 0.263637217106105), 0, 0, 0);
            poly128.AddVertexAt(4, new Point2d(scale1 * 0.458367272636016, scale1 * 0.261954013223059), 0, 0, 0);
            poly128.AddVertexAt(5, new Point2d(scale1 * 0.449717474903696, scale1 * 0.261392945262043), 0, 0, 0);
            poly128.AddVertexAt(6, new Point2d(scale1 * 0.445041908561901, scale1 * 0.26237481419382), 0, 0, 0);
            poly128.AddVertexAt(7, new Point2d(scale1 * 0.441301455488465, scale1 * 0.26489962001839), 0, 0, 0);
            poly128.AddVertexAt(8, new Point2d(scale1 * 0.430781431219427, scale1 * 0.268967362735751), 0, 0, 0);
            poly128.AddVertexAt(9, new Point2d(scale1 * 0.425170751609273, scale1 * 0.272894838462859), 0, 0, 0);
            poly128.AddVertexAt(10, new Point2d(scale1 * 0.426573421511812, scale1 * 0.27471830933616), 0, 0, 0);
            poly128.AddVertexAt(11, new Point2d(scale1 * 0.434989440927043, scale1 * 0.278365251082757), 0, 0, 0);
            poly128.AddVertexAt(12, new Point2d(scale1 * 0.441535233805555, scale1 * 0.279066586034028), 0, 0, 0);
            poly128.AddVertexAt(13, new Point2d(scale1 * 0.445743243513171, scale1 * 0.277804183121743), 0, 0, 0);
            poly128.AddVertexAt(14, new Point2d(scale1 * 0.451353923123324, scale1 * 0.276541780209458), 0, 0, 0);
            poly128.AddVertexAt(15, new Point2d(scale1 * 0.455795711148029, scale1 * 0.275840445258188), 0, 0, 0);
            poly128.Closed = true;
            poly128.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly128.Layer = "0";
            poly128.Color = color_GP;
            poly128.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly128);
            Polyline poly129 = new Polyline();
            poly129.AddVertexAt(0, new Point2d(scale1 * 0.478898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly129.AddVertexAt(1, new Point2d(scale1 * 0.472741026607183, scale1 * 0.262017770945895), 0, 0, 0);
            poly129.AddVertexAt(2, new Point2d(scale1 * 0.46743251686181, scale1 * 0.259722492923566), 0, 0, 0);
            poly129.AddVertexAt(3, new Point2d(scale1 * 0.466370814912736, scale1 * 0.256662122227115), 0, 0, 0);
            poly129.AddVertexAt(4, new Point2d(scale1 * 0.47167932465811, scale1 * 0.255769514107321), 0, 0, 0);
            poly129.AddVertexAt(5, new Point2d(scale1 * 0.478261876742373, scale1 * 0.2567896376728), 0, 0, 0);
            poly129.AddVertexAt(6, new Point2d(scale1 * 0.482084003759041, scale1 * 0.259467462032192), 0, 0, 0);
            poly129.AddVertexAt(7, new Point2d(scale1 * 0.478898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly129.Closed = true;
            poly129.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly129.Layer = "0";
            poly129.Color = color_GP;
            poly129.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly129);
            Polyline poly130 = new Polyline();
            poly130.AddVertexAt(0, new Point2d(scale1 * 0.496473138321644, scale1 * 0.248628649148944), 0, 0, 0);
            poly130.AddVertexAt(1, new Point2d(scale1 * 0.496706916638733, scale1 * 0.243298503519296), 0, 0, 0);
            poly130.AddVertexAt(2, new Point2d(scale1 * 0.495070468419106, scale1 * 0.242036100607014), 0, 0, 0);
            poly130.AddVertexAt(3, new Point2d(scale1 * 0.49086245871149, scale1 * 0.240633430704475), 0, 0, 0);
            poly130.AddVertexAt(4, new Point2d(scale1 * 0.486888227320964, scale1 * 0.240773697694727), 0, 0, 0);
            poly130.AddVertexAt(5, new Point2d(scale1 * 0.483849109198798, scale1 * 0.242036100607014), 0, 0, 0);
            poly130.AddVertexAt(6, new Point2d(scale1 * 0.480576212759541, scale1 * 0.243298503519296), 0, 0, 0);
            poly130.AddVertexAt(7, new Point2d(scale1 * 0.477537094637375, scale1 * 0.244560906431583), 0, 0, 0);
            poly130.AddVertexAt(8, new Point2d(scale1 * 0.475666868100657, scale1 * 0.246945445265898), 0, 0, 0);
            poly130.AddVertexAt(9, new Point2d(scale1 * 0.475666868100657, scale1 * 0.248628649148944), 0, 0, 0);
            poly130.AddVertexAt(10, new Point2d(scale1 * 0.479173542857003, scale1 * 0.252275590895543), 0, 0, 0);
            poly130.AddVertexAt(11, new Point2d(scale1 * 0.485953114052605, scale1 * 0.253818527788335), 0, 0, 0);
            poly130.AddVertexAt(12, new Point2d(scale1 * 0.492732685248208, scale1 * 0.252275590895543), 0, 0, 0);
            poly130.Closed = true;
            poly130.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly130.Layer = "0";
            poly130.Color = color_GP;
            poly130.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly130);
            Polyline poly131 = new Polyline();
            poly131.AddVertexAt(0, new Point2d(scale1 * 0.462107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly131.AddVertexAt(1, new Point2d(scale1 * 0.459302385904375, scale1 * 0.241895833616759), 0, 0, 0);
            poly131.AddVertexAt(2, new Point2d(scale1 * 0.460237499172734, scale1 * 0.240072362743458), 0, 0, 0);
            poly131.AddVertexAt(3, new Point2d(scale1 * 0.461640169075273, scale1 * 0.237968357889651), 0, 0, 0);
            poly131.AddVertexAt(4, new Point2d(scale1 * 0.464679287197439, scale1 * 0.237687823909144), 0, 0, 0);
            poly131.AddVertexAt(5, new Point2d(scale1 * 0.468185961953785, scale1 * 0.23937102779219), 0, 0, 0);
            poly131.AddVertexAt(6, new Point2d(scale1 * 0.468185961953785, scale1 * 0.241615299636252), 0, 0, 0);
            poly131.AddVertexAt(7, new Point2d(scale1 * 0.465848178782888, scale1 * 0.244140105460821), 0, 0, 0);
            poly131.AddVertexAt(8, new Point2d(scale1 * 0.462107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly131.Closed = true;
            poly131.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly131.Layer = "0";
            poly131.Color = color_GP;
            poly131.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly131);
            Polyline poly132 = new Polyline();
            poly132.AddVertexAt(0, new Point2d(scale1 * 0.416520953876953, scale1 * 0.248628649148944), 0, 0, 0);
            poly132.AddVertexAt(1, new Point2d(scale1 * 0.423066746755466, scale1 * 0.250452120022245), 0, 0, 0);
            poly132.AddVertexAt(2, new Point2d(scale1 * 0.435924554195401, scale1 * 0.249750785070974), 0, 0, 0);
            poly132.AddVertexAt(3, new Point2d(scale1 * 0.442470347073914, scale1 * 0.248628649148944), 0, 0, 0);
            poly132.AddVertexAt(4, new Point2d(scale1 * 0.443639238659363, scale1 * 0.246805178275643), 0, 0, 0);
            poly132.AddVertexAt(5, new Point2d(scale1 * 0.443873016976452, scale1 * 0.244981707402343), 0, 0, 0);
            poly132.AddVertexAt(6, new Point2d(scale1 * 0.444106795293542, scale1 * 0.242737435558282), 0, 0, 0);
            poly132.AddVertexAt(7, new Point2d(scale1 * 0.44597702183026, scale1 * 0.240773697694727), 0, 0, 0);
            poly132.AddVertexAt(8, new Point2d(scale1 * 0.444574351927722, scale1 * 0.237407289928637), 0, 0, 0);
            poly132.AddVertexAt(9, new Point2d(scale1 * 0.439431228951747, scale1 * 0.235583819055336), 0, 0, 0);
            poly132.AddVertexAt(10, new Point2d(scale1 * 0.429378761316889, scale1 * 0.235303285074829), 0, 0, 0);
            poly132.AddVertexAt(11, new Point2d(scale1 * 0.424469416658004, scale1 * 0.236144887016352), 0, 0, 0);
            poly132.AddVertexAt(12, new Point2d(scale1 * 0.420962741901658, scale1 * 0.237828090899398), 0, 0, 0);
            poly132.AddVertexAt(13, new Point2d(scale1 * 0.417689845462402, scale1 * 0.241054231675236), 0, 0, 0);
            poly132.AddVertexAt(14, new Point2d(scale1 * 0.415819618925684, scale1 * 0.244420639441328), 0, 0, 0);
            poly132.Closed = true;
            poly132.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly132.Layer = "0";
            poly132.Color = color_GP;
            poly132.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly132);
            Polyline poly133 = new Polyline();
            poly133.AddVertexAt(0, new Point2d(scale1 * 0.360285741001707, scale1 * 0.3), 0, 0, 0);
            poly133.AddVertexAt(1, new Point2d(scale1 * 0.361640169075273, scale1 * 0.297968357889651), 0, 0, 0);
            poly133.AddVertexAt(2, new Point2d(scale1 * 0.364679287197439, scale1 * 0.297687823909144), 0, 0, 0);
            poly133.AddVertexAt(3, new Point2d(scale1 * 0.368185961953785, scale1 * 0.29937102779219), 0, 0, 0);
            poly133.AddVertexAt(4, new Point2d(scale1 * 0.368185961953785, scale1 * 0.3), 0, 0, 0);
            poly133.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly133.Layer = "0";
            poly133.Color = color_GP;
            poly133.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly133);
            Polyline poly134 = new Polyline();
            poly134.AddVertexAt(0, new Point2d(scale1 * 0.34565464779079, scale1 * 0.3), 0, 0, 0);
            poly134.AddVertexAt(1, new Point2d(scale1 * 0.344574351927722, scale1 * 0.297407289928637), 0, 0, 0);
            poly134.AddVertexAt(2, new Point2d(scale1 * 0.339431228951747, scale1 * 0.295583819055336), 0, 0, 0);
            poly134.AddVertexAt(3, new Point2d(scale1 * 0.329378761316889, scale1 * 0.295303285074829), 0, 0, 0);
            poly134.AddVertexAt(4, new Point2d(scale1 * 0.324469416658004, scale1 * 0.296144887016352), 0, 0, 0);
            poly134.AddVertexAt(5, new Point2d(scale1 * 0.320962741901658, scale1 * 0.297828090899398), 0, 0, 0);
            poly134.AddVertexAt(6, new Point2d(scale1 * 0.318759355857569, scale1 * 0.3), 0, 0, 0);
            poly134.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly134.Layer = "0";
            poly134.Color = color_GP;
            poly134.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly134);
            Polyline poly135 = new Polyline();
            poly135.AddVertexAt(0, new Point2d(scale1 * 0.358367272636016, scale1 * 0.282994061761134), 0, 0, 0);
            poly135.AddVertexAt(1, new Point2d(scale1 * 0.356497046099298, scale1 * 0.282853794770882), 0, 0, 0);
            poly135.AddVertexAt(2, new Point2d(scale1 * 0.354159262928401, scale1 * 0.28355512972215), 0, 0, 0);
            poly135.AddVertexAt(3, new Point2d(scale1 * 0.352289036391683, scale1 * 0.285518867585705), 0, 0, 0);
            poly135.AddVertexAt(4, new Point2d(scale1 * 0.353925484611311, scale1 * 0.287061804478495), 0, 0, 0);
            poly135.AddVertexAt(5, new Point2d(scale1 * 0.358367272636016, scale1 * 0.286220202536974), 0, 0, 0);
            poly135.AddVertexAt(6, new Point2d(scale1 * 0.359536164221465, scale1 * 0.284817532634435), 0, 0, 0);
            poly135.Closed = true;
            poly135.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly135.Layer = "0";
            poly135.Color = color_GP;
            poly135.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly135);
            Polyline poly136 = new Polyline();
            poly136.AddVertexAt(0, new Point2d(scale1 * 0.367952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly136.AddVertexAt(1, new Point2d(scale1 * 0.366081957099978, scale1 * 0.282012192829359), 0, 0, 0);
            poly136.AddVertexAt(2, new Point2d(scale1 * 0.365848178782888, scale1 * 0.280048454965803), 0, 0, 0);
            poly136.AddVertexAt(3, new Point2d(scale1 * 0.368419740270875, scale1 * 0.278505518073012), 0, 0, 0);
            poly136.AddVertexAt(4, new Point2d(scale1 * 0.371692636710131, scale1 * 0.278926319043773), 0, 0, 0);
            poly136.AddVertexAt(5, new Point2d(scale1 * 0.372393971661401, scale1 * 0.281030323897581), 0, 0, 0);
            poly136.AddVertexAt(6, new Point2d(scale1 * 0.371926415027221, scale1 * 0.282713527780627), 0, 0, 0);
            poly136.AddVertexAt(7, new Point2d(scale1 * 0.367952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly136.Closed = true;
            poly136.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly136.Layer = "0";
            poly136.Color = color_GP;
            poly136.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly136);
            Polyline poly137 = new Polyline();
            poly137.AddVertexAt(0, new Point2d(scale1 * 0.316577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly137.AddVertexAt(1, new Point2d(scale1 * 0.314239328956148, scale1 * 0.285251085149766), 0, 0, 0);
            poly137.AddVertexAt(2, new Point2d(scale1 * 0.315641998858686, scale1 * 0.283427614276465), 0, 0, 0);
            poly137.AddVertexAt(3, new Point2d(scale1 * 0.319616230249212, scale1 * 0.282726279325197), 0, 0, 0);
            poly137.AddVertexAt(4, new Point2d(scale1 * 0.322655348371378, scale1 * 0.283988682237481), 0, 0, 0);
            poly137.AddVertexAt(5, new Point2d(scale1 * 0.322655348371378, scale1 * 0.286653755052303), 0, 0, 0);
            poly137.AddVertexAt(6, new Point2d(scale1 * 0.316577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly137.Closed = true;
            poly137.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly137.Layer = "0";
            poly137.Color = color_GP;
            poly137.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly137);
            Polyline poly138 = new Polyline();
            poly138.AddVertexAt(0, new Point2d(scale1 * 0.311080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly138.AddVertexAt(1, new Point2d(scale1 * 0.309476961807309, scale1 * 0.263254670769046), 0, 0, 0);
            poly138.AddVertexAt(2, new Point2d(scale1 * 0.309710740124399, scale1 * 0.259607729022446), 0, 0, 0);
            poly138.AddVertexAt(3, new Point2d(scale1 * 0.310645853392758, scale1 * 0.256802389217368), 0, 0, 0);
            poly138.AddVertexAt(4, new Point2d(scale1 * 0.314853863100373, scale1 * 0.255960787275844), 0, 0, 0);
            poly138.AddVertexAt(5, new Point2d(scale1 * 0.319061872807989, scale1 * 0.256521855236861), 0, 0, 0);
            poly138.AddVertexAt(6, new Point2d(scale1 * 0.322568547564335, scale1 * 0.258625860090668), 0, 0, 0);
            poly138.AddVertexAt(7, new Point2d(scale1 * 0.319061872807989, scale1 * 0.263535204749552), 0, 0, 0);
            poly138.AddVertexAt(8, new Point2d(scale1 * 0.311080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly138.Closed = true;
            poly138.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly138.Layer = "0";
            poly138.Color = color_GP;
            poly138.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly138);
            Polyline poly139 = new Polyline();
            poly139.AddVertexAt(0, new Point2d(scale1 * 0.387975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly139.AddVertexAt(1, new Point2d(scale1 * 0.384702548523924, scale1 * 0.288936281530065), 0, 0, 0);
            poly139.AddVertexAt(2, new Point2d(scale1 * 0.382598543670117, scale1 * 0.283746402890672), 0, 0, 0);
            poly139.AddVertexAt(3, new Point2d(scale1 * 0.382598543670117, scale1 * 0.281221597066103), 0, 0, 0);
            poly139.AddVertexAt(4, new Point2d(scale1 * 0.381897208718848, scale1 * 0.275751184446203), 0, 0, 0);
            poly139.AddVertexAt(5, new Point2d(scale1 * 0.385403883475194, scale1 * 0.273366645611888), 0, 0, 0);
            poly139.AddVertexAt(6, new Point2d(scale1 * 0.390313228134078, scale1 * 0.272525043670365), 0, 0, 0);
            poly139.AddVertexAt(7, new Point2d(scale1 * 0.401768365671475, scale1 * 0.272945844641126), 0, 0, 0);
            poly139.AddVertexAt(8, new Point2d(scale1 * 0.404807483793642, scale1 * 0.274488781533918), 0, 0, 0);
            poly139.AddVertexAt(9, new Point2d(scale1 * 0.405742597062001, scale1 * 0.278275990270772), 0, 0, 0);
            poly139.AddVertexAt(10, new Point2d(scale1 * 0.402703478939834, scale1 * 0.28037999512458), 0, 0, 0);
            poly139.AddVertexAt(11, new Point2d(scale1 * 0.399898139134758, scale1 * 0.282764533958894), 0, 0, 0);
            poly139.AddVertexAt(12, new Point2d(scale1 * 0.39756035596386, scale1 * 0.286271208715241), 0, 0, 0);
            poly139.AddVertexAt(13, new Point2d(scale1 * 0.395690129427142, scale1 * 0.288796014539811), 0, 0, 0);
            poly139.AddVertexAt(14, new Point2d(scale1 * 0.391482119719527, scale1 * 0.289777883471588), 0, 0, 0);
            poly139.AddVertexAt(15, new Point2d(scale1 * 0.387975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly139.Closed = true;
            poly139.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly139.Layer = "0";
            poly139.Color = color_GP;
            poly139.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly139);
            Polyline poly140 = new Polyline();
            poly140.AddVertexAt(0, new Point2d(scale1 * 0.358367272636016, scale1 * 0.274437775355651), 0, 0, 0);
            poly140.AddVertexAt(1, new Point2d(scale1 * 0.366081957099978, scale1 * 0.272754571472604), 0, 0, 0);
            poly140.AddVertexAt(2, new Point2d(scale1 * 0.369354853539234, scale1 * 0.268125760794228), 0, 0, 0);
            poly140.AddVertexAt(3, new Point2d(scale1 * 0.366081957099978, scale1 * 0.263637217106105), 0, 0, 0);
            poly140.AddVertexAt(4, new Point2d(scale1 * 0.358367272636016, scale1 * 0.261954013223059), 0, 0, 0);
            poly140.AddVertexAt(5, new Point2d(scale1 * 0.349717474903696, scale1 * 0.261392945262043), 0, 0, 0);
            poly140.AddVertexAt(6, new Point2d(scale1 * 0.345041908561901, scale1 * 0.26237481419382), 0, 0, 0);
            poly140.AddVertexAt(7, new Point2d(scale1 * 0.341301455488465, scale1 * 0.26489962001839), 0, 0, 0);
            poly140.AddVertexAt(8, new Point2d(scale1 * 0.330781431219427, scale1 * 0.268967362735751), 0, 0, 0);
            poly140.AddVertexAt(9, new Point2d(scale1 * 0.325170751609273, scale1 * 0.272894838462859), 0, 0, 0);
            poly140.AddVertexAt(10, new Point2d(scale1 * 0.326573421511812, scale1 * 0.27471830933616), 0, 0, 0);
            poly140.AddVertexAt(11, new Point2d(scale1 * 0.334989440927042, scale1 * 0.278365251082757), 0, 0, 0);
            poly140.AddVertexAt(12, new Point2d(scale1 * 0.341535233805555, scale1 * 0.279066586034028), 0, 0, 0);
            poly140.AddVertexAt(13, new Point2d(scale1 * 0.34574324351317, scale1 * 0.277804183121743), 0, 0, 0);
            poly140.AddVertexAt(14, new Point2d(scale1 * 0.351353923123324, scale1 * 0.276541780209458), 0, 0, 0);
            poly140.AddVertexAt(15, new Point2d(scale1 * 0.355795711148029, scale1 * 0.275840445258188), 0, 0, 0);
            poly140.Closed = true;
            poly140.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly140.Layer = "0";
            poly140.Color = color_GP;
            poly140.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly140);
            Polyline poly141 = new Polyline();
            poly141.AddVertexAt(0, new Point2d(scale1 * 0.378898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly141.AddVertexAt(1, new Point2d(scale1 * 0.372741026607184, scale1 * 0.262017770945895), 0, 0, 0);
            poly141.AddVertexAt(2, new Point2d(scale1 * 0.36743251686181, scale1 * 0.259722492923566), 0, 0, 0);
            poly141.AddVertexAt(3, new Point2d(scale1 * 0.366370814912736, scale1 * 0.256662122227115), 0, 0, 0);
            poly141.AddVertexAt(4, new Point2d(scale1 * 0.37167932465811, scale1 * 0.255769514107321), 0, 0, 0);
            poly141.AddVertexAt(5, new Point2d(scale1 * 0.378261876742373, scale1 * 0.2567896376728), 0, 0, 0);
            poly141.AddVertexAt(6, new Point2d(scale1 * 0.382084003759042, scale1 * 0.259467462032192), 0, 0, 0);
            poly141.AddVertexAt(7, new Point2d(scale1 * 0.378898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly141.Closed = true;
            poly141.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly141.Layer = "0";
            poly141.Color = color_GP;
            poly141.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly141);
            Polyline poly142 = new Polyline();
            poly142.AddVertexAt(0, new Point2d(scale1 * 0.396473138321644, scale1 * 0.248628649148944), 0, 0, 0);
            poly142.AddVertexAt(1, new Point2d(scale1 * 0.396706916638733, scale1 * 0.243298503519296), 0, 0, 0);
            poly142.AddVertexAt(2, new Point2d(scale1 * 0.395070468419106, scale1 * 0.242036100607014), 0, 0, 0);
            poly142.AddVertexAt(3, new Point2d(scale1 * 0.39086245871149, scale1 * 0.240633430704475), 0, 0, 0);
            poly142.AddVertexAt(4, new Point2d(scale1 * 0.386888227320964, scale1 * 0.240773697694727), 0, 0, 0);
            poly142.AddVertexAt(5, new Point2d(scale1 * 0.383849109198798, scale1 * 0.242036100607014), 0, 0, 0);
            poly142.AddVertexAt(6, new Point2d(scale1 * 0.380576212759542, scale1 * 0.243298503519296), 0, 0, 0);
            poly142.AddVertexAt(7, new Point2d(scale1 * 0.377537094637375, scale1 * 0.244560906431583), 0, 0, 0);
            poly142.AddVertexAt(8, new Point2d(scale1 * 0.375666868100657, scale1 * 0.246945445265898), 0, 0, 0);
            poly142.AddVertexAt(9, new Point2d(scale1 * 0.375666868100657, scale1 * 0.248628649148944), 0, 0, 0);
            poly142.AddVertexAt(10, new Point2d(scale1 * 0.379173542857003, scale1 * 0.252275590895543), 0, 0, 0);
            poly142.AddVertexAt(11, new Point2d(scale1 * 0.385953114052606, scale1 * 0.253818527788335), 0, 0, 0);
            poly142.AddVertexAt(12, new Point2d(scale1 * 0.392732685248208, scale1 * 0.252275590895543), 0, 0, 0);
            poly142.Closed = true;
            poly142.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly142.Layer = "0";
            poly142.Color = color_GP;
            poly142.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly142);
            Polyline poly143 = new Polyline();
            poly143.AddVertexAt(0, new Point2d(scale1 * 0.362107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly143.AddVertexAt(1, new Point2d(scale1 * 0.359302385904375, scale1 * 0.241895833616759), 0, 0, 0);
            poly143.AddVertexAt(2, new Point2d(scale1 * 0.360237499172734, scale1 * 0.240072362743458), 0, 0, 0);
            poly143.AddVertexAt(3, new Point2d(scale1 * 0.361640169075273, scale1 * 0.237968357889651), 0, 0, 0);
            poly143.AddVertexAt(4, new Point2d(scale1 * 0.364679287197439, scale1 * 0.237687823909144), 0, 0, 0);
            poly143.AddVertexAt(5, new Point2d(scale1 * 0.368185961953785, scale1 * 0.23937102779219), 0, 0, 0);
            poly143.AddVertexAt(6, new Point2d(scale1 * 0.368185961953785, scale1 * 0.241615299636252), 0, 0, 0);
            poly143.AddVertexAt(7, new Point2d(scale1 * 0.365848178782888, scale1 * 0.244140105460821), 0, 0, 0);
            poly143.AddVertexAt(8, new Point2d(scale1 * 0.362107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly143.Closed = true;
            poly143.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly143.Layer = "0";
            poly143.Color = color_GP;
            poly143.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly143);
            Polyline poly144 = new Polyline();
            poly144.AddVertexAt(0, new Point2d(scale1 * 0.316520953876953, scale1 * 0.248628649148944), 0, 0, 0);
            poly144.AddVertexAt(1, new Point2d(scale1 * 0.323066746755466, scale1 * 0.250452120022245), 0, 0, 0);
            poly144.AddVertexAt(2, new Point2d(scale1 * 0.335924554195401, scale1 * 0.249750785070974), 0, 0, 0);
            poly144.AddVertexAt(3, new Point2d(scale1 * 0.342470347073914, scale1 * 0.248628649148944), 0, 0, 0);
            poly144.AddVertexAt(4, new Point2d(scale1 * 0.343639238659363, scale1 * 0.246805178275643), 0, 0, 0);
            poly144.AddVertexAt(5, new Point2d(scale1 * 0.343873016976453, scale1 * 0.244981707402343), 0, 0, 0);
            poly144.AddVertexAt(6, new Point2d(scale1 * 0.344106795293542, scale1 * 0.242737435558282), 0, 0, 0);
            poly144.AddVertexAt(7, new Point2d(scale1 * 0.34597702183026, scale1 * 0.240773697694727), 0, 0, 0);
            poly144.AddVertexAt(8, new Point2d(scale1 * 0.344574351927722, scale1 * 0.237407289928637), 0, 0, 0);
            poly144.AddVertexAt(9, new Point2d(scale1 * 0.339431228951747, scale1 * 0.235583819055336), 0, 0, 0);
            poly144.AddVertexAt(10, new Point2d(scale1 * 0.329378761316889, scale1 * 0.235303285074829), 0, 0, 0);
            poly144.AddVertexAt(11, new Point2d(scale1 * 0.324469416658004, scale1 * 0.236144887016352), 0, 0, 0);
            poly144.AddVertexAt(12, new Point2d(scale1 * 0.320962741901658, scale1 * 0.237828090899398), 0, 0, 0);
            poly144.AddVertexAt(13, new Point2d(scale1 * 0.317689845462402, scale1 * 0.241054231675236), 0, 0, 0);
            poly144.AddVertexAt(14, new Point2d(scale1 * 0.315819618925684, scale1 * 0.244420639441328), 0, 0, 0);
            poly144.Closed = true;
            poly144.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly144.Layer = "0";
            poly144.Color = color_GP;
            poly144.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly144);
            Polyline poly145 = new Polyline();
            poly145.AddVertexAt(0, new Point2d(scale1 * 0.260285741001707, scale1 * 0.3), 0, 0, 0);
            poly145.AddVertexAt(1, new Point2d(scale1 * 0.261640169075273, scale1 * 0.297968357889651), 0, 0, 0);
            poly145.AddVertexAt(2, new Point2d(scale1 * 0.264679287197439, scale1 * 0.297687823909144), 0, 0, 0);
            poly145.AddVertexAt(3, new Point2d(scale1 * 0.268185961953785, scale1 * 0.29937102779219), 0, 0, 0);
            poly145.AddVertexAt(4, new Point2d(scale1 * 0.268185961953785, scale1 * 0.3), 0, 0, 0);
            poly145.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly145.Layer = "0";
            poly145.Color = color_GP;
            poly145.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly145);
            Polyline poly146 = new Polyline();
            poly146.AddVertexAt(0, new Point2d(scale1 * 0.24565464779079, scale1 * 0.3), 0, 0, 0);
            poly146.AddVertexAt(1, new Point2d(scale1 * 0.244574351927722, scale1 * 0.297407289928637), 0, 0, 0);
            poly146.AddVertexAt(2, new Point2d(scale1 * 0.239431228951748, scale1 * 0.295583819055336), 0, 0, 0);
            poly146.AddVertexAt(3, new Point2d(scale1 * 0.229378761316889, scale1 * 0.295303285074829), 0, 0, 0);
            poly146.AddVertexAt(4, new Point2d(scale1 * 0.224469416658004, scale1 * 0.296144887016352), 0, 0, 0);
            poly146.AddVertexAt(5, new Point2d(scale1 * 0.220962741901658, scale1 * 0.297828090899398), 0, 0, 0);
            poly146.AddVertexAt(6, new Point2d(scale1 * 0.218759355857569, scale1 * 0.3), 0, 0, 0);
            poly146.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly146.Layer = "0";
            poly146.Color = color_GP;
            poly146.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly146);
            Polyline poly147 = new Polyline();
            poly147.AddVertexAt(0, new Point2d(scale1 * 0.258367272636016, scale1 * 0.282994061761134), 0, 0, 0);
            poly147.AddVertexAt(1, new Point2d(scale1 * 0.256497046099298, scale1 * 0.282853794770882), 0, 0, 0);
            poly147.AddVertexAt(2, new Point2d(scale1 * 0.254159262928401, scale1 * 0.28355512972215), 0, 0, 0);
            poly147.AddVertexAt(3, new Point2d(scale1 * 0.252289036391683, scale1 * 0.285518867585705), 0, 0, 0);
            poly147.AddVertexAt(4, new Point2d(scale1 * 0.253925484611311, scale1 * 0.287061804478495), 0, 0, 0);
            poly147.AddVertexAt(5, new Point2d(scale1 * 0.258367272636016, scale1 * 0.286220202536974), 0, 0, 0);
            poly147.AddVertexAt(6, new Point2d(scale1 * 0.259536164221465, scale1 * 0.284817532634435), 0, 0, 0);
            poly147.Closed = true;
            poly147.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly147.Layer = "0";
            poly147.Color = color_GP;
            poly147.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly147);
            Polyline poly148 = new Polyline();
            poly148.AddVertexAt(0, new Point2d(scale1 * 0.267952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly148.AddVertexAt(1, new Point2d(scale1 * 0.266081957099978, scale1 * 0.282012192829359), 0, 0, 0);
            poly148.AddVertexAt(2, new Point2d(scale1 * 0.265848178782888, scale1 * 0.280048454965803), 0, 0, 0);
            poly148.AddVertexAt(3, new Point2d(scale1 * 0.268419740270875, scale1 * 0.278505518073012), 0, 0, 0);
            poly148.AddVertexAt(4, new Point2d(scale1 * 0.271692636710131, scale1 * 0.278926319043773), 0, 0, 0);
            poly148.AddVertexAt(5, new Point2d(scale1 * 0.272393971661401, scale1 * 0.281030323897581), 0, 0, 0);
            poly148.AddVertexAt(6, new Point2d(scale1 * 0.271926415027221, scale1 * 0.282713527780627), 0, 0, 0);
            poly148.AddVertexAt(7, new Point2d(scale1 * 0.267952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly148.Closed = true;
            poly148.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly148.Layer = "0";
            poly148.Color = color_GP;
            poly148.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly148);
            Polyline poly149 = new Polyline();
            poly149.AddVertexAt(0, new Point2d(scale1 * 0.216577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly149.AddVertexAt(1, new Point2d(scale1 * 0.214239328956148, scale1 * 0.285251085149766), 0, 0, 0);
            poly149.AddVertexAt(2, new Point2d(scale1 * 0.215641998858686, scale1 * 0.283427614276465), 0, 0, 0);
            poly149.AddVertexAt(3, new Point2d(scale1 * 0.219616230249212, scale1 * 0.282726279325197), 0, 0, 0);
            poly149.AddVertexAt(4, new Point2d(scale1 * 0.222655348371378, scale1 * 0.283988682237481), 0, 0, 0);
            poly149.AddVertexAt(5, new Point2d(scale1 * 0.222655348371378, scale1 * 0.286653755052303), 0, 0, 0);
            poly149.AddVertexAt(6, new Point2d(scale1 * 0.216577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly149.Closed = true;
            poly149.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly149.Layer = "0";
            poly149.Color = color_GP;
            poly149.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly149);
            Polyline poly150 = new Polyline();
            poly150.AddVertexAt(0, new Point2d(scale1 * 0.211080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly150.AddVertexAt(1, new Point2d(scale1 * 0.209476961807309, scale1 * 0.263254670769046), 0, 0, 0);
            poly150.AddVertexAt(2, new Point2d(scale1 * 0.209710740124399, scale1 * 0.259607729022446), 0, 0, 0);
            poly150.AddVertexAt(3, new Point2d(scale1 * 0.210645853392758, scale1 * 0.256802389217368), 0, 0, 0);
            poly150.AddVertexAt(4, new Point2d(scale1 * 0.214853863100373, scale1 * 0.255960787275844), 0, 0, 0);
            poly150.AddVertexAt(5, new Point2d(scale1 * 0.219061872807989, scale1 * 0.256521855236861), 0, 0, 0);
            poly150.AddVertexAt(6, new Point2d(scale1 * 0.222568547564335, scale1 * 0.258625860090668), 0, 0, 0);
            poly150.AddVertexAt(7, new Point2d(scale1 * 0.219061872807989, scale1 * 0.263535204749552), 0, 0, 0);
            poly150.AddVertexAt(8, new Point2d(scale1 * 0.211080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly150.Closed = true;
            poly150.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly150.Layer = "0";
            poly150.Color = color_GP;
            poly150.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly150);
            Polyline poly151 = new Polyline();
            poly151.AddVertexAt(0, new Point2d(scale1 * 0.287975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly151.AddVertexAt(1, new Point2d(scale1 * 0.284702548523924, scale1 * 0.288936281530065), 0, 0, 0);
            poly151.AddVertexAt(2, new Point2d(scale1 * 0.282598543670117, scale1 * 0.283746402890672), 0, 0, 0);
            poly151.AddVertexAt(3, new Point2d(scale1 * 0.282598543670117, scale1 * 0.281221597066103), 0, 0, 0);
            poly151.AddVertexAt(4, new Point2d(scale1 * 0.281897208718848, scale1 * 0.275751184446203), 0, 0, 0);
            poly151.AddVertexAt(5, new Point2d(scale1 * 0.285403883475194, scale1 * 0.273366645611888), 0, 0, 0);
            poly151.AddVertexAt(6, new Point2d(scale1 * 0.290313228134079, scale1 * 0.272525043670365), 0, 0, 0);
            poly151.AddVertexAt(7, new Point2d(scale1 * 0.301768365671475, scale1 * 0.272945844641126), 0, 0, 0);
            poly151.AddVertexAt(8, new Point2d(scale1 * 0.304807483793642, scale1 * 0.274488781533918), 0, 0, 0);
            poly151.AddVertexAt(9, new Point2d(scale1 * 0.305742597062001, scale1 * 0.278275990270772), 0, 0, 0);
            poly151.AddVertexAt(10, new Point2d(scale1 * 0.302703478939834, scale1 * 0.28037999512458), 0, 0, 0);
            poly151.AddVertexAt(11, new Point2d(scale1 * 0.299898139134757, scale1 * 0.282764533958894), 0, 0, 0);
            poly151.AddVertexAt(12, new Point2d(scale1 * 0.29756035596386, scale1 * 0.286271208715241), 0, 0, 0);
            poly151.AddVertexAt(13, new Point2d(scale1 * 0.295690129427142, scale1 * 0.288796014539811), 0, 0, 0);
            poly151.AddVertexAt(14, new Point2d(scale1 * 0.291482119719527, scale1 * 0.289777883471588), 0, 0, 0);
            poly151.AddVertexAt(15, new Point2d(scale1 * 0.287975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly151.Closed = true;
            poly151.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly151.Layer = "0";
            poly151.Color = color_GP;
            poly151.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly151);
            Polyline poly152 = new Polyline();
            poly152.AddVertexAt(0, new Point2d(scale1 * 0.258367272636016, scale1 * 0.274437775355651), 0, 0, 0);
            poly152.AddVertexAt(1, new Point2d(scale1 * 0.266081957099978, scale1 * 0.272754571472604), 0, 0, 0);
            poly152.AddVertexAt(2, new Point2d(scale1 * 0.269354853539234, scale1 * 0.268125760794228), 0, 0, 0);
            poly152.AddVertexAt(3, new Point2d(scale1 * 0.266081957099978, scale1 * 0.263637217106105), 0, 0, 0);
            poly152.AddVertexAt(4, new Point2d(scale1 * 0.258367272636016, scale1 * 0.261954013223059), 0, 0, 0);
            poly152.AddVertexAt(5, new Point2d(scale1 * 0.249717474903696, scale1 * 0.261392945262043), 0, 0, 0);
            poly152.AddVertexAt(6, new Point2d(scale1 * 0.245041908561901, scale1 * 0.26237481419382), 0, 0, 0);
            poly152.AddVertexAt(7, new Point2d(scale1 * 0.241301455488465, scale1 * 0.26489962001839), 0, 0, 0);
            poly152.AddVertexAt(8, new Point2d(scale1 * 0.230781431219427, scale1 * 0.268967362735751), 0, 0, 0);
            poly152.AddVertexAt(9, new Point2d(scale1 * 0.225170751609273, scale1 * 0.272894838462859), 0, 0, 0);
            poly152.AddVertexAt(10, new Point2d(scale1 * 0.226573421511812, scale1 * 0.27471830933616), 0, 0, 0);
            poly152.AddVertexAt(11, new Point2d(scale1 * 0.234989440927042, scale1 * 0.278365251082757), 0, 0, 0);
            poly152.AddVertexAt(12, new Point2d(scale1 * 0.241535233805555, scale1 * 0.279066586034028), 0, 0, 0);
            poly152.AddVertexAt(13, new Point2d(scale1 * 0.24574324351317, scale1 * 0.277804183121743), 0, 0, 0);
            poly152.AddVertexAt(14, new Point2d(scale1 * 0.251353923123324, scale1 * 0.276541780209458), 0, 0, 0);
            poly152.AddVertexAt(15, new Point2d(scale1 * 0.255795711148029, scale1 * 0.275840445258188), 0, 0, 0);
            poly152.Closed = true;
            poly152.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly152.Layer = "0";
            poly152.Color = color_GP;
            poly152.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly152);
            Polyline poly153 = new Polyline();
            poly153.AddVertexAt(0, new Point2d(scale1 * 0.278898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly153.AddVertexAt(1, new Point2d(scale1 * 0.272741026607183, scale1 * 0.262017770945895), 0, 0, 0);
            poly153.AddVertexAt(2, new Point2d(scale1 * 0.267432516861809, scale1 * 0.259722492923566), 0, 0, 0);
            poly153.AddVertexAt(3, new Point2d(scale1 * 0.266370814912736, scale1 * 0.256662122227115), 0, 0, 0);
            poly153.AddVertexAt(4, new Point2d(scale1 * 0.27167932465811, scale1 * 0.255769514107321), 0, 0, 0);
            poly153.AddVertexAt(5, new Point2d(scale1 * 0.278261876742373, scale1 * 0.2567896376728), 0, 0, 0);
            poly153.AddVertexAt(6, new Point2d(scale1 * 0.282084003759042, scale1 * 0.259467462032192), 0, 0, 0);
            poly153.AddVertexAt(7, new Point2d(scale1 * 0.278898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly153.Closed = true;
            poly153.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly153.Layer = "0";
            poly153.Color = color_GP;
            poly153.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly153);
            Polyline poly154 = new Polyline();
            poly154.AddVertexAt(0, new Point2d(scale1 * 0.296473138321644, scale1 * 0.248628649148944), 0, 0, 0);
            poly154.AddVertexAt(1, new Point2d(scale1 * 0.296706916638733, scale1 * 0.243298503519296), 0, 0, 0);
            poly154.AddVertexAt(2, new Point2d(scale1 * 0.295070468419106, scale1 * 0.242036100607014), 0, 0, 0);
            poly154.AddVertexAt(3, new Point2d(scale1 * 0.29086245871149, scale1 * 0.240633430704475), 0, 0, 0);
            poly154.AddVertexAt(4, new Point2d(scale1 * 0.286888227320964, scale1 * 0.240773697694727), 0, 0, 0);
            poly154.AddVertexAt(5, new Point2d(scale1 * 0.283849109198798, scale1 * 0.242036100607014), 0, 0, 0);
            poly154.AddVertexAt(6, new Point2d(scale1 * 0.280576212759541, scale1 * 0.243298503519296), 0, 0, 0);
            poly154.AddVertexAt(7, new Point2d(scale1 * 0.277537094637375, scale1 * 0.244560906431583), 0, 0, 0);
            poly154.AddVertexAt(8, new Point2d(scale1 * 0.275666868100657, scale1 * 0.246945445265898), 0, 0, 0);
            poly154.AddVertexAt(9, new Point2d(scale1 * 0.275666868100657, scale1 * 0.248628649148944), 0, 0, 0);
            poly154.AddVertexAt(10, new Point2d(scale1 * 0.279173542857003, scale1 * 0.252275590895543), 0, 0, 0);
            poly154.AddVertexAt(11, new Point2d(scale1 * 0.285953114052605, scale1 * 0.253818527788335), 0, 0, 0);
            poly154.AddVertexAt(12, new Point2d(scale1 * 0.292732685248208, scale1 * 0.252275590895543), 0, 0, 0);
            poly154.Closed = true;
            poly154.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly154.Layer = "0";
            poly154.Color = color_GP;
            poly154.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly154);
            Polyline poly155 = new Polyline();
            poly155.AddVertexAt(0, new Point2d(scale1 * 0.262107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly155.AddVertexAt(1, new Point2d(scale1 * 0.259302385904375, scale1 * 0.241895833616759), 0, 0, 0);
            poly155.AddVertexAt(2, new Point2d(scale1 * 0.260237499172734, scale1 * 0.240072362743458), 0, 0, 0);
            poly155.AddVertexAt(3, new Point2d(scale1 * 0.261640169075273, scale1 * 0.237968357889651), 0, 0, 0);
            poly155.AddVertexAt(4, new Point2d(scale1 * 0.264679287197439, scale1 * 0.237687823909144), 0, 0, 0);
            poly155.AddVertexAt(5, new Point2d(scale1 * 0.268185961953785, scale1 * 0.23937102779219), 0, 0, 0);
            poly155.AddVertexAt(6, new Point2d(scale1 * 0.268185961953785, scale1 * 0.241615299636252), 0, 0, 0);
            poly155.AddVertexAt(7, new Point2d(scale1 * 0.265848178782888, scale1 * 0.244140105460821), 0, 0, 0);
            poly155.AddVertexAt(8, new Point2d(scale1 * 0.262107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly155.Closed = true;
            poly155.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly155.Layer = "0";
            poly155.Color = color_GP;
            poly155.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly155);
            Polyline poly156 = new Polyline();
            poly156.AddVertexAt(0, new Point2d(scale1 * 0.216520953876953, scale1 * 0.248628649148944), 0, 0, 0);
            poly156.AddVertexAt(1, new Point2d(scale1 * 0.223066746755466, scale1 * 0.250452120022245), 0, 0, 0);
            poly156.AddVertexAt(2, new Point2d(scale1 * 0.235924554195401, scale1 * 0.249750785070974), 0, 0, 0);
            poly156.AddVertexAt(3, new Point2d(scale1 * 0.242470347073914, scale1 * 0.248628649148944), 0, 0, 0);
            poly156.AddVertexAt(4, new Point2d(scale1 * 0.243639238659363, scale1 * 0.246805178275643), 0, 0, 0);
            poly156.AddVertexAt(5, new Point2d(scale1 * 0.243873016976452, scale1 * 0.244981707402343), 0, 0, 0);
            poly156.AddVertexAt(6, new Point2d(scale1 * 0.244106795293542, scale1 * 0.242737435558282), 0, 0, 0);
            poly156.AddVertexAt(7, new Point2d(scale1 * 0.24597702183026, scale1 * 0.240773697694727), 0, 0, 0);
            poly156.AddVertexAt(8, new Point2d(scale1 * 0.244574351927722, scale1 * 0.237407289928637), 0, 0, 0);
            poly156.AddVertexAt(9, new Point2d(scale1 * 0.239431228951748, scale1 * 0.235583819055336), 0, 0, 0);
            poly156.AddVertexAt(10, new Point2d(scale1 * 0.229378761316889, scale1 * 0.235303285074829), 0, 0, 0);
            poly156.AddVertexAt(11, new Point2d(scale1 * 0.224469416658004, scale1 * 0.236144887016352), 0, 0, 0);
            poly156.AddVertexAt(12, new Point2d(scale1 * 0.220962741901658, scale1 * 0.237828090899398), 0, 0, 0);
            poly156.AddVertexAt(13, new Point2d(scale1 * 0.217689845462402, scale1 * 0.241054231675236), 0, 0, 0);
            poly156.AddVertexAt(14, new Point2d(scale1 * 0.215819618925684, scale1 * 0.244420639441328), 0, 0, 0);
            poly156.Closed = true;
            poly156.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly156.Layer = "0";
            poly156.Color = color_GP;
            poly156.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly156);
            Polyline poly157 = new Polyline();
            poly157.AddVertexAt(0, new Point2d(scale1 * 0.160285741001707, scale1 * 0.3), 0, 0, 0);
            poly157.AddVertexAt(1, new Point2d(scale1 * 0.161640169075273, scale1 * 0.297968357889651), 0, 0, 0);
            poly157.AddVertexAt(2, new Point2d(scale1 * 0.164679287197439, scale1 * 0.297687823909144), 0, 0, 0);
            poly157.AddVertexAt(3, new Point2d(scale1 * 0.168185961953785, scale1 * 0.29937102779219), 0, 0, 0);
            poly157.AddVertexAt(4, new Point2d(scale1 * 0.168185961953785, scale1 * 0.3), 0, 0, 0);
            poly157.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly157.Layer = "0";
            poly157.Color = color_GP;
            poly157.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly157);
            Polyline poly158 = new Polyline();
            poly158.AddVertexAt(0, new Point2d(scale1 * 0.14565464779079, scale1 * 0.3), 0, 0, 0);
            poly158.AddVertexAt(1, new Point2d(scale1 * 0.144574351927722, scale1 * 0.297407289928637), 0, 0, 0);
            poly158.AddVertexAt(2, new Point2d(scale1 * 0.139431228951747, scale1 * 0.295583819055336), 0, 0, 0);
            poly158.AddVertexAt(3, new Point2d(scale1 * 0.129378761316889, scale1 * 0.295303285074829), 0, 0, 0);
            poly158.AddVertexAt(4, new Point2d(scale1 * 0.124469416658004, scale1 * 0.296144887016352), 0, 0, 0);
            poly158.AddVertexAt(5, new Point2d(scale1 * 0.120962741901658, scale1 * 0.297828090899398), 0, 0, 0);
            poly158.AddVertexAt(6, new Point2d(scale1 * 0.118759355857569, scale1 * 0.3), 0, 0, 0);
            poly158.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly158.Layer = "0";
            poly158.Color = color_GP;
            poly158.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly158);
            Polyline poly159 = new Polyline();
            poly159.AddVertexAt(0, new Point2d(scale1 * 0.158367272636016, scale1 * 0.282994061761134), 0, 0, 0);
            poly159.AddVertexAt(1, new Point2d(scale1 * 0.156497046099298, scale1 * 0.282853794770882), 0, 0, 0);
            poly159.AddVertexAt(2, new Point2d(scale1 * 0.154159262928401, scale1 * 0.28355512972215), 0, 0, 0);
            poly159.AddVertexAt(3, new Point2d(scale1 * 0.152289036391683, scale1 * 0.285518867585705), 0, 0, 0);
            poly159.AddVertexAt(4, new Point2d(scale1 * 0.153925484611311, scale1 * 0.287061804478495), 0, 0, 0);
            poly159.AddVertexAt(5, new Point2d(scale1 * 0.158367272636016, scale1 * 0.286220202536974), 0, 0, 0);
            poly159.AddVertexAt(6, new Point2d(scale1 * 0.159536164221465, scale1 * 0.284817532634435), 0, 0, 0);
            poly159.Closed = true;
            poly159.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly159.Layer = "0";
            poly159.Color = color_GP;
            poly159.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly159);
            Polyline poly160 = new Polyline();
            poly160.AddVertexAt(0, new Point2d(scale1 * 0.167952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly160.AddVertexAt(1, new Point2d(scale1 * 0.166081957099978, scale1 * 0.282012192829359), 0, 0, 0);
            poly160.AddVertexAt(2, new Point2d(scale1 * 0.165848178782888, scale1 * 0.280048454965803), 0, 0, 0);
            poly160.AddVertexAt(3, new Point2d(scale1 * 0.168419740270875, scale1 * 0.278505518073012), 0, 0, 0);
            poly160.AddVertexAt(4, new Point2d(scale1 * 0.171692636710131, scale1 * 0.278926319043773), 0, 0, 0);
            poly160.AddVertexAt(5, new Point2d(scale1 * 0.172393971661401, scale1 * 0.281030323897581), 0, 0, 0);
            poly160.AddVertexAt(6, new Point2d(scale1 * 0.171926415027221, scale1 * 0.282713527780627), 0, 0, 0);
            poly160.AddVertexAt(7, new Point2d(scale1 * 0.167952183636696, scale1 * 0.28355512972215), 0, 0, 0);
            poly160.Closed = true;
            poly160.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly160.Layer = "0";
            poly160.Color = color_GP;
            poly160.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly160);
            Polyline poly161 = new Polyline();
            poly161.AddVertexAt(0, new Point2d(scale1 * 0.116577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly161.AddVertexAt(1, new Point2d(scale1 * 0.114239328956148, scale1 * 0.285251085149766), 0, 0, 0);
            poly161.AddVertexAt(2, new Point2d(scale1 * 0.115641998858686, scale1 * 0.283427614276465), 0, 0, 0);
            poly161.AddVertexAt(3, new Point2d(scale1 * 0.119616230249212, scale1 * 0.282726279325197), 0, 0, 0);
            poly161.AddVertexAt(4, new Point2d(scale1 * 0.122655348371378, scale1 * 0.283988682237481), 0, 0, 0);
            poly161.AddVertexAt(5, new Point2d(scale1 * 0.122655348371378, scale1 * 0.286653755052303), 0, 0, 0);
            poly161.AddVertexAt(6, new Point2d(scale1 * 0.116577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly161.Closed = true;
            poly161.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly161.Layer = "0";
            poly161.Color = color_GP;
            poly161.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly161);
            Polyline poly162 = new Polyline();
            poly162.AddVertexAt(0, new Point2d(scale1 * 0.111080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly162.AddVertexAt(1, new Point2d(scale1 * 0.109476961807309, scale1 * 0.263254670769046), 0, 0, 0);
            poly162.AddVertexAt(2, new Point2d(scale1 * 0.109710740124399, scale1 * 0.259607729022446), 0, 0, 0);
            poly162.AddVertexAt(3, new Point2d(scale1 * 0.110645853392758, scale1 * 0.256802389217368), 0, 0, 0);
            poly162.AddVertexAt(4, new Point2d(scale1 * 0.114853863100373, scale1 * 0.255960787275844), 0, 0, 0);
            poly162.AddVertexAt(5, new Point2d(scale1 * 0.119061872807989, scale1 * 0.256521855236861), 0, 0, 0);
            poly162.AddVertexAt(6, new Point2d(scale1 * 0.122568547564335, scale1 * 0.258625860090668), 0, 0, 0);
            poly162.AddVertexAt(7, new Point2d(scale1 * 0.119061872807989, scale1 * 0.263535204749552), 0, 0, 0);
            poly162.AddVertexAt(8, new Point2d(scale1 * 0.111080758243531, scale1 * 0.266735842436258), 0, 0, 0);
            poly162.Closed = true;
            poly162.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly162.Layer = "0";
            poly162.Color = color_GP;
            poly162.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly162);
            Polyline poly163 = new Polyline();
            poly163.AddVertexAt(0, new Point2d(scale1 * 0.187975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly163.AddVertexAt(1, new Point2d(scale1 * 0.184702548523924, scale1 * 0.288936281530065), 0, 0, 0);
            poly163.AddVertexAt(2, new Point2d(scale1 * 0.182598543670117, scale1 * 0.283746402890672), 0, 0, 0);
            poly163.AddVertexAt(3, new Point2d(scale1 * 0.182598543670117, scale1 * 0.281221597066103), 0, 0, 0);
            poly163.AddVertexAt(4, new Point2d(scale1 * 0.181897208718848, scale1 * 0.275751184446203), 0, 0, 0);
            poly163.AddVertexAt(5, new Point2d(scale1 * 0.185403883475194, scale1 * 0.273366645611888), 0, 0, 0);
            poly163.AddVertexAt(6, new Point2d(scale1 * 0.190313228134078, scale1 * 0.272525043670365), 0, 0, 0);
            poly163.AddVertexAt(7, new Point2d(scale1 * 0.201768365671475, scale1 * 0.272945844641126), 0, 0, 0);
            poly163.AddVertexAt(8, new Point2d(scale1 * 0.204807483793642, scale1 * 0.274488781533918), 0, 0, 0);
            poly163.AddVertexAt(9, new Point2d(scale1 * 0.205742597062001, scale1 * 0.278275990270772), 0, 0, 0);
            poly163.AddVertexAt(10, new Point2d(scale1 * 0.202703478939834, scale1 * 0.28037999512458), 0, 0, 0);
            poly163.AddVertexAt(11, new Point2d(scale1 * 0.199898139134758, scale1 * 0.282764533958894), 0, 0, 0);
            poly163.AddVertexAt(12, new Point2d(scale1 * 0.19756035596386, scale1 * 0.286271208715241), 0, 0, 0);
            poly163.AddVertexAt(13, new Point2d(scale1 * 0.195690129427142, scale1 * 0.288796014539811), 0, 0, 0);
            poly163.AddVertexAt(14, new Point2d(scale1 * 0.191482119719527, scale1 * 0.289777883471588), 0, 0, 0);
            poly163.AddVertexAt(15, new Point2d(scale1 * 0.187975444963181, scale1 * 0.290759752403364), 0, 0, 0);
            poly163.Closed = true;
            poly163.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly163.Layer = "0";
            poly163.Color = color_GP;
            poly163.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly163);
            Polyline poly164 = new Polyline();
            poly164.AddVertexAt(0, new Point2d(scale1 * 0.158367272636016, scale1 * 0.274437775355651), 0, 0, 0);
            poly164.AddVertexAt(1, new Point2d(scale1 * 0.166081957099978, scale1 * 0.272754571472604), 0, 0, 0);
            poly164.AddVertexAt(2, new Point2d(scale1 * 0.169354853539234, scale1 * 0.268125760794228), 0, 0, 0);
            poly164.AddVertexAt(3, new Point2d(scale1 * 0.166081957099978, scale1 * 0.263637217106105), 0, 0, 0);
            poly164.AddVertexAt(4, new Point2d(scale1 * 0.158367272636016, scale1 * 0.261954013223059), 0, 0, 0);
            poly164.AddVertexAt(5, new Point2d(scale1 * 0.149717474903696, scale1 * 0.261392945262043), 0, 0, 0);
            poly164.AddVertexAt(6, new Point2d(scale1 * 0.145041908561901, scale1 * 0.26237481419382), 0, 0, 0);
            poly164.AddVertexAt(7, new Point2d(scale1 * 0.141301455488465, scale1 * 0.26489962001839), 0, 0, 0);
            poly164.AddVertexAt(8, new Point2d(scale1 * 0.130781431219427, scale1 * 0.268967362735751), 0, 0, 0);
            poly164.AddVertexAt(9, new Point2d(scale1 * 0.125170751609273, scale1 * 0.272894838462859), 0, 0, 0);
            poly164.AddVertexAt(10, new Point2d(scale1 * 0.126573421511812, scale1 * 0.27471830933616), 0, 0, 0);
            poly164.AddVertexAt(11, new Point2d(scale1 * 0.134989440927042, scale1 * 0.278365251082757), 0, 0, 0);
            poly164.AddVertexAt(12, new Point2d(scale1 * 0.141535233805555, scale1 * 0.279066586034028), 0, 0, 0);
            poly164.AddVertexAt(13, new Point2d(scale1 * 0.14574324351317, scale1 * 0.277804183121743), 0, 0, 0);
            poly164.AddVertexAt(14, new Point2d(scale1 * 0.151353923123324, scale1 * 0.276541780209458), 0, 0, 0);
            poly164.AddVertexAt(15, new Point2d(scale1 * 0.155795711148029, scale1 * 0.275840445258188), 0, 0, 0);
            poly164.Closed = true;
            poly164.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly164.Layer = "0";
            poly164.Color = color_GP;
            poly164.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly164);
            Polyline poly165 = new Polyline();
            poly165.AddVertexAt(0, new Point2d(scale1 * 0.178898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly165.AddVertexAt(1, new Point2d(scale1 * 0.172741026607183, scale1 * 0.262017770945895), 0, 0, 0);
            poly165.AddVertexAt(2, new Point2d(scale1 * 0.167432516861809, scale1 * 0.259722492923566), 0, 0, 0);
            poly165.AddVertexAt(3, new Point2d(scale1 * 0.166370814912736, scale1 * 0.256662122227115), 0, 0, 0);
            poly165.AddVertexAt(4, new Point2d(scale1 * 0.17167932465811, scale1 * 0.255769514107321), 0, 0, 0);
            poly165.AddVertexAt(5, new Point2d(scale1 * 0.178261876742373, scale1 * 0.2567896376728), 0, 0, 0);
            poly165.AddVertexAt(6, new Point2d(scale1 * 0.182084003759041, scale1 * 0.259467462032192), 0, 0, 0);
            poly165.AddVertexAt(7, new Point2d(scale1 * 0.178898897911817, scale1 * 0.261890255500208), 0, 0, 0);
            poly165.Closed = true;
            poly165.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly165.Layer = "0";
            poly165.Color = color_GP;
            poly165.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly165);
            Polyline poly166 = new Polyline();
            poly166.AddVertexAt(0, new Point2d(scale1 * 0.196473138321644, scale1 * 0.248628649148944), 0, 0, 0);
            poly166.AddVertexAt(1, new Point2d(scale1 * 0.196706916638733, scale1 * 0.243298503519296), 0, 0, 0);
            poly166.AddVertexAt(2, new Point2d(scale1 * 0.195070468419106, scale1 * 0.242036100607014), 0, 0, 0);
            poly166.AddVertexAt(3, new Point2d(scale1 * 0.19086245871149, scale1 * 0.240633430704475), 0, 0, 0);
            poly166.AddVertexAt(4, new Point2d(scale1 * 0.186888227320964, scale1 * 0.240773697694727), 0, 0, 0);
            poly166.AddVertexAt(5, new Point2d(scale1 * 0.183849109198798, scale1 * 0.242036100607014), 0, 0, 0);
            poly166.AddVertexAt(6, new Point2d(scale1 * 0.180576212759542, scale1 * 0.243298503519296), 0, 0, 0);
            poly166.AddVertexAt(7, new Point2d(scale1 * 0.177537094637375, scale1 * 0.244560906431583), 0, 0, 0);
            poly166.AddVertexAt(8, new Point2d(scale1 * 0.175666868100657, scale1 * 0.246945445265898), 0, 0, 0);
            poly166.AddVertexAt(9, new Point2d(scale1 * 0.175666868100657, scale1 * 0.248628649148944), 0, 0, 0);
            poly166.AddVertexAt(10, new Point2d(scale1 * 0.179173542857003, scale1 * 0.252275590895543), 0, 0, 0);
            poly166.AddVertexAt(11, new Point2d(scale1 * 0.185953114052605, scale1 * 0.253818527788335), 0, 0, 0);
            poly166.AddVertexAt(12, new Point2d(scale1 * 0.192732685248208, scale1 * 0.252275590895543), 0, 0, 0);
            poly166.Closed = true;
            poly166.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly166.Layer = "0";
            poly166.Color = color_GP;
            poly166.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly166);
            Polyline poly167 = new Polyline();
            poly167.AddVertexAt(0, new Point2d(scale1 * 0.162107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly167.AddVertexAt(1, new Point2d(scale1 * 0.159302385904375, scale1 * 0.241895833616759), 0, 0, 0);
            poly167.AddVertexAt(2, new Point2d(scale1 * 0.160237499172734, scale1 * 0.240072362743458), 0, 0, 0);
            poly167.AddVertexAt(3, new Point2d(scale1 * 0.161640169075273, scale1 * 0.237968357889651), 0, 0, 0);
            poly167.AddVertexAt(4, new Point2d(scale1 * 0.164679287197439, scale1 * 0.237687823909144), 0, 0, 0);
            poly167.AddVertexAt(5, new Point2d(scale1 * 0.168185961953785, scale1 * 0.23937102779219), 0, 0, 0);
            poly167.AddVertexAt(6, new Point2d(scale1 * 0.168185961953785, scale1 * 0.241615299636252), 0, 0, 0);
            poly167.AddVertexAt(7, new Point2d(scale1 * 0.165848178782888, scale1 * 0.244140105460821), 0, 0, 0);
            poly167.AddVertexAt(8, new Point2d(scale1 * 0.162107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly167.Closed = true;
            poly167.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly167.Layer = "0";
            poly167.Color = color_GP;
            poly167.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly167);
            Polyline poly168 = new Polyline();
            poly168.AddVertexAt(0, new Point2d(scale1 * 0.116520953876953, scale1 * 0.248628649148944), 0, 0, 0);
            poly168.AddVertexAt(1, new Point2d(scale1 * 0.123066746755466, scale1 * 0.250452120022245), 0, 0, 0);
            poly168.AddVertexAt(2, new Point2d(scale1 * 0.135924554195401, scale1 * 0.249750785070974), 0, 0, 0);
            poly168.AddVertexAt(3, new Point2d(scale1 * 0.142470347073914, scale1 * 0.248628649148944), 0, 0, 0);
            poly168.AddVertexAt(4, new Point2d(scale1 * 0.143639238659363, scale1 * 0.246805178275643), 0, 0, 0);
            poly168.AddVertexAt(5, new Point2d(scale1 * 0.143873016976452, scale1 * 0.244981707402343), 0, 0, 0);
            poly168.AddVertexAt(6, new Point2d(scale1 * 0.144106795293542, scale1 * 0.242737435558282), 0, 0, 0);
            poly168.AddVertexAt(7, new Point2d(scale1 * 0.14597702183026, scale1 * 0.240773697694727), 0, 0, 0);
            poly168.AddVertexAt(8, new Point2d(scale1 * 0.144574351927722, scale1 * 0.237407289928637), 0, 0, 0);
            poly168.AddVertexAt(9, new Point2d(scale1 * 0.139431228951747, scale1 * 0.235583819055336), 0, 0, 0);
            poly168.AddVertexAt(10, new Point2d(scale1 * 0.129378761316889, scale1 * 0.235303285074829), 0, 0, 0);
            poly168.AddVertexAt(11, new Point2d(scale1 * 0.124469416658004, scale1 * 0.236144887016352), 0, 0, 0);
            poly168.AddVertexAt(12, new Point2d(scale1 * 0.120962741901658, scale1 * 0.237828090899398), 0, 0, 0);
            poly168.AddVertexAt(13, new Point2d(scale1 * 0.117689845462402, scale1 * 0.241054231675236), 0, 0, 0);
            poly168.AddVertexAt(14, new Point2d(scale1 * 0.115819618925684, scale1 * 0.244420639441328), 0, 0, 0);
            poly168.Closed = true;
            poly168.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly168.Layer = "0";
            poly168.Color = color_GP;
            poly168.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly168);
            Polyline poly169 = new Polyline();
            poly169.AddVertexAt(0, new Point2d(scale1 * 0.0602857410017066, scale1 * 0.3), 0, 0, 0);
            poly169.AddVertexAt(1, new Point2d(scale1 * 0.0616401690752726, scale1 * 0.297968357889651), 0, 0, 0);
            poly169.AddVertexAt(2, new Point2d(scale1 * 0.0646792871974391, scale1 * 0.297687823909144), 0, 0, 0);
            poly169.AddVertexAt(3, new Point2d(scale1 * 0.0681859619537855, scale1 * 0.29937102779219), 0, 0, 0);
            poly169.AddVertexAt(4, new Point2d(scale1 * 0.0681859619537855, scale1 * 0.3), 0, 0, 0);
            poly169.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly169.Layer = "0";
            poly169.Color = color_GP;
            poly169.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly169);
            Polyline poly170 = new Polyline();
            poly170.AddVertexAt(0, new Point2d(scale1 * 0.0456546477907903, scale1 * 0.3), 0, 0, 0);
            poly170.AddVertexAt(1, new Point2d(scale1 * 0.0445743519277215, scale1 * 0.297407289928637), 0, 0, 0);
            poly170.AddVertexAt(2, new Point2d(scale1 * 0.0394312289517476, scale1 * 0.295583819055336), 0, 0, 0);
            poly170.AddVertexAt(3, new Point2d(scale1 * 0.0293787613168888, scale1 * 0.295303285074829), 0, 0, 0);
            poly170.AddVertexAt(4, new Point2d(scale1 * 0.024469416658004, scale1 * 0.296144887016352), 0, 0, 0);
            poly170.AddVertexAt(5, new Point2d(scale1 * 0.020962741901658, scale1 * 0.297828090899398), 0, 0, 0);
            poly170.AddVertexAt(6, new Point2d(scale1 * 0.0187593558575691, scale1 * 0.3), 0, 0, 0);
            poly170.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly170.Layer = "0";
            poly170.Color = color_GP;
            poly170.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly170);
            Polyline poly171 = new Polyline();
            poly171.AddVertexAt(0, new Point2d(scale1 * 0.0583672726360163, scale1 * 0.282994061761134), 0, 0, 0);
            poly171.AddVertexAt(1, new Point2d(scale1 * 0.0564970460992982, scale1 * 0.282853794770882), 0, 0, 0);
            poly171.AddVertexAt(2, new Point2d(scale1 * 0.0541592629284009, scale1 * 0.28355512972215), 0, 0, 0);
            poly171.AddVertexAt(3, new Point2d(scale1 * 0.052289036391683, scale1 * 0.285518867585705), 0, 0, 0);
            poly171.AddVertexAt(4, new Point2d(scale1 * 0.0539254846113113, scale1 * 0.287061804478495), 0, 0, 0);
            poly171.AddVertexAt(5, new Point2d(scale1 * 0.0583672726360163, scale1 * 0.286220202536974), 0, 0, 0);
            poly171.AddVertexAt(6, new Point2d(scale1 * 0.0595361642214651, scale1 * 0.284817532634435), 0, 0, 0);
            poly171.Closed = true;
            poly171.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly171.Layer = "0";
            poly171.Color = color_GP;
            poly171.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly171);
            Polyline poly172 = new Polyline();
            poly172.AddVertexAt(0, new Point2d(scale1 * 0.0679521836366956, scale1 * 0.28355512972215), 0, 0, 0);
            poly172.AddVertexAt(1, new Point2d(scale1 * 0.0660819570999778, scale1 * 0.282012192829359), 0, 0, 0);
            poly172.AddVertexAt(2, new Point2d(scale1 * 0.065848178782888, scale1 * 0.280048454965803), 0, 0, 0);
            poly172.AddVertexAt(3, new Point2d(scale1 * 0.0684197402708751, scale1 * 0.278505518073012), 0, 0, 0);
            poly172.AddVertexAt(4, new Point2d(scale1 * 0.0716926367101314, scale1 * 0.278926319043773), 0, 0, 0);
            poly172.AddVertexAt(5, new Point2d(scale1 * 0.0723939716614006, scale1 * 0.281030323897581), 0, 0, 0);
            poly172.AddVertexAt(6, new Point2d(scale1 * 0.0719264150272212, scale1 * 0.282713527780627), 0, 0, 0);
            poly172.AddVertexAt(7, new Point2d(scale1 * 0.0679521836366956, scale1 * 0.28355512972215), 0, 0, 0);
            poly172.Closed = true;
            poly172.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly172.Layer = "0";
            poly172.Color = color_GP;
            poly172.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly172);
            Polyline poly173 = new Polyline();
            poly173.AddVertexAt(0, new Point2d(scale1 * 0.016577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly173.AddVertexAt(1, new Point2d(scale1 * 0.0142393289561475, scale1 * 0.285251085149766), 0, 0, 0);
            poly173.AddVertexAt(2, new Point2d(scale1 * 0.015641998858686, scale1 * 0.283427614276465), 0, 0, 0);
            poly173.AddVertexAt(3, new Point2d(scale1 * 0.0196162302492116, scale1 * 0.282726279325197), 0, 0, 0);
            poly173.AddVertexAt(4, new Point2d(scale1 * 0.0226553483713783, scale1 * 0.283988682237481), 0, 0, 0);
            poly173.AddVertexAt(5, new Point2d(scale1 * 0.0226553483713783, scale1 * 0.286653755052303), 0, 0, 0);
            poly173.AddVertexAt(6, new Point2d(scale1 * 0.016577112127045, scale1 * 0.287355090003573), 0, 0, 0);
            poly173.Closed = true;
            poly173.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly173.Layer = "0";
            poly173.Color = color_GP;
            poly173.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly173);
            Polyline poly174 = new Polyline();
            poly174.AddVertexAt(0, new Point2d(scale1 * 0.0110807582435306, scale1 * 0.266735842436258), 0, 0, 0);
            poly174.AddVertexAt(1, new Point2d(scale1 * 0.00947696180730939, scale1 * 0.263254670769046), 0, 0, 0);
            poly174.AddVertexAt(2, new Point2d(scale1 * 0.00971074012439921, scale1 * 0.259607729022446), 0, 0, 0);
            poly174.AddVertexAt(3, new Point2d(scale1 * 0.010645853392758, scale1 * 0.256802389217368), 0, 0, 0);
            poly174.AddVertexAt(4, new Point2d(scale1 * 0.0148538631003734, scale1 * 0.255960787275844), 0, 0, 0);
            poly174.AddVertexAt(5, new Point2d(scale1 * 0.0190618728079888, scale1 * 0.256521855236861), 0, 0, 0);
            poly174.AddVertexAt(6, new Point2d(scale1 * 0.0225685475643347, scale1 * 0.258625860090668), 0, 0, 0);
            poly174.AddVertexAt(7, new Point2d(scale1 * 0.0190618728079888, scale1 * 0.263535204749552), 0, 0, 0);
            poly174.AddVertexAt(8, new Point2d(scale1 * 0.0110807582435306, scale1 * 0.266735842436258), 0, 0, 0);
            poly174.Closed = true;
            poly174.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly174.Layer = "0";
            poly174.Color = color_GP;
            poly174.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly174);
            Polyline poly175 = new Polyline();
            poly175.AddVertexAt(0, new Point2d(scale1 * 0.0879754449631811, scale1 * 0.290759752403364), 0, 0, 0);
            poly175.AddVertexAt(1, new Point2d(scale1 * 0.0847025485239243, scale1 * 0.288936281530065), 0, 0, 0);
            poly175.AddVertexAt(2, new Point2d(scale1 * 0.0825985436701167, scale1 * 0.283746402890672), 0, 0, 0);
            poly175.AddVertexAt(3, new Point2d(scale1 * 0.0825985436701167, scale1 * 0.281221597066103), 0, 0, 0);
            poly175.AddVertexAt(4, new Point2d(scale1 * 0.0818972087188476, scale1 * 0.275751184446203), 0, 0, 0);
            poly175.AddVertexAt(5, new Point2d(scale1 * 0.0854038834751936, scale1 * 0.273366645611888), 0, 0, 0);
            poly175.AddVertexAt(6, new Point2d(scale1 * 0.0903132281340784, scale1 * 0.272525043670365), 0, 0, 0);
            poly175.AddVertexAt(7, new Point2d(scale1 * 0.101768365671475, scale1 * 0.272945844641126), 0, 0, 0);
            poly175.AddVertexAt(8, new Point2d(scale1 * 0.104807483793642, scale1 * 0.274488781533918), 0, 0, 0);
            poly175.AddVertexAt(9, new Point2d(scale1 * 0.105742597062001, scale1 * 0.278275990270772), 0, 0, 0);
            poly175.AddVertexAt(10, new Point2d(scale1 * 0.102703478939834, scale1 * 0.28037999512458), 0, 0, 0);
            poly175.AddVertexAt(11, new Point2d(scale1 * 0.0998981391347575, scale1 * 0.282764533958894), 0, 0, 0);
            poly175.AddVertexAt(12, new Point2d(scale1 * 0.09756035596386, scale1 * 0.286271208715241), 0, 0, 0);
            poly175.AddVertexAt(13, new Point2d(scale1 * 0.0956901294271417, scale1 * 0.288796014539811), 0, 0, 0);
            poly175.AddVertexAt(14, new Point2d(scale1 * 0.0914821197195268, scale1 * 0.289777883471588), 0, 0, 0);
            poly175.AddVertexAt(15, new Point2d(scale1 * 0.0879754449631811, scale1 * 0.290759752403364), 0, 0, 0);
            poly175.Closed = true;
            poly175.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly175.Layer = "0";
            poly175.Color = color_GP;
            poly175.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly175);
            Polyline poly176 = new Polyline();
            poly176.AddVertexAt(0, new Point2d(scale1 * 0.0583672726360163, scale1 * 0.274437775355651), 0, 0, 0);
            poly176.AddVertexAt(1, new Point2d(scale1 * 0.0660819570999778, scale1 * 0.272754571472604), 0, 0, 0);
            poly176.AddVertexAt(2, new Point2d(scale1 * 0.0693548535392339, scale1 * 0.268125760794228), 0, 0, 0);
            poly176.AddVertexAt(3, new Point2d(scale1 * 0.0660819570999778, scale1 * 0.263637217106105), 0, 0, 0);
            poly176.AddVertexAt(4, new Point2d(scale1 * 0.0583672726360163, scale1 * 0.261954013223059), 0, 0, 0);
            poly176.AddVertexAt(5, new Point2d(scale1 * 0.0497174749036959, scale1 * 0.261392945262043), 0, 0, 0);
            poly176.AddVertexAt(6, new Point2d(scale1 * 0.0450419085619012, scale1 * 0.26237481419382), 0, 0, 0);
            poly176.AddVertexAt(7, new Point2d(scale1 * 0.0413014554884654, scale1 * 0.26489962001839), 0, 0, 0);
            poly176.AddVertexAt(8, new Point2d(scale1 * 0.030781431219427, scale1 * 0.268967362735751), 0, 0, 0);
            poly176.AddVertexAt(9, new Point2d(scale1 * 0.0251707516092734, scale1 * 0.272894838462859), 0, 0, 0);
            poly176.AddVertexAt(10, new Point2d(scale1 * 0.0265734215118119, scale1 * 0.27471830933616), 0, 0, 0);
            poly176.AddVertexAt(11, new Point2d(scale1 * 0.0349894409270424, scale1 * 0.278365251082757), 0, 0, 0);
            poly176.AddVertexAt(12, new Point2d(scale1 * 0.041535233805555, scale1 * 0.279066586034028), 0, 0, 0);
            poly176.AddVertexAt(13, new Point2d(scale1 * 0.0457432435131704, scale1 * 0.277804183121743), 0, 0, 0);
            poly176.AddVertexAt(14, new Point2d(scale1 * 0.0513539231233242, scale1 * 0.276541780209458), 0, 0, 0);
            poly176.AddVertexAt(15, new Point2d(scale1 * 0.0557957111480292, scale1 * 0.275840445258188), 0, 0, 0);
            poly176.Closed = true;
            poly176.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly176.Layer = "0";
            poly176.Color = color_GP;
            poly176.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly176);
            Polyline poly177 = new Polyline();
            poly177.AddVertexAt(0, new Point2d(scale1 * 0.0788988979118168, scale1 * 0.261890255500208), 0, 0, 0);
            poly177.AddVertexAt(1, new Point2d(scale1 * 0.0727410266071835, scale1 * 0.262017770945895), 0, 0, 0);
            poly177.AddVertexAt(2, new Point2d(scale1 * 0.0674325168618095, scale1 * 0.259722492923566), 0, 0, 0);
            poly177.AddVertexAt(3, new Point2d(scale1 * 0.0663708149127362, scale1 * 0.256662122227115), 0, 0, 0);
            poly177.AddVertexAt(4, new Point2d(scale1 * 0.07167932465811, scale1 * 0.255769514107321), 0, 0, 0);
            poly177.AddVertexAt(5, new Point2d(scale1 * 0.0782618767423733, scale1 * 0.2567896376728), 0, 0, 0);
            poly177.AddVertexAt(6, new Point2d(scale1 * 0.0820840037590416, scale1 * 0.259467462032192), 0, 0, 0);
            poly177.AddVertexAt(7, new Point2d(scale1 * 0.0788988979118168, scale1 * 0.261890255500208), 0, 0, 0);
            poly177.Closed = true;
            poly177.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly177.Layer = "0";
            poly177.Color = color_GP;
            poly177.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly177);
            Polyline poly178 = new Polyline();
            poly178.AddVertexAt(0, new Point2d(scale1 * 0.0964731383216439, scale1 * 0.248628649148944), 0, 0, 0);
            poly178.AddVertexAt(1, new Point2d(scale1 * 0.0967069166387331, scale1 * 0.243298503519296), 0, 0, 0);
            poly178.AddVertexAt(2, new Point2d(scale1 * 0.0950704684191055, scale1 * 0.242036100607014), 0, 0, 0);
            poly178.AddVertexAt(3, new Point2d(scale1 * 0.0908624587114899, scale1 * 0.240633430704475), 0, 0, 0);
            poly178.AddVertexAt(4, new Point2d(scale1 * 0.0868882273209644, scale1 * 0.240773697694727), 0, 0, 0);
            poly178.AddVertexAt(5, new Point2d(scale1 * 0.0838491091987978, scale1 * 0.242036100607014), 0, 0, 0);
            poly178.AddVertexAt(6, new Point2d(scale1 * 0.0805762127595415, scale1 * 0.243298503519296), 0, 0, 0);
            poly178.AddVertexAt(7, new Point2d(scale1 * 0.0775370946373748, scale1 * 0.244560906431583), 0, 0, 0);
            poly178.AddVertexAt(8, new Point2d(scale1 * 0.0756668681006569, scale1 * 0.246945445265898), 0, 0, 0);
            poly178.AddVertexAt(9, new Point2d(scale1 * 0.0756668681006569, scale1 * 0.248628649148944), 0, 0, 0);
            poly178.AddVertexAt(10, new Point2d(scale1 * 0.0791735428570031, scale1 * 0.252275590895543), 0, 0, 0);
            poly178.AddVertexAt(11, new Point2d(scale1 * 0.0859531140526055, scale1 * 0.253818527788335), 0, 0, 0);
            poly178.AddVertexAt(12, new Point2d(scale1 * 0.0927326852482082, scale1 * 0.252275590895543), 0, 0, 0);
            poly178.Closed = true;
            poly178.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly178.Layer = "0";
            poly178.Color = color_GP;
            poly178.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly178);
            Polyline poly179 = new Polyline();
            poly179.AddVertexAt(0, new Point2d(scale1 * 0.062107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly179.AddVertexAt(1, new Point2d(scale1 * 0.0593023859043753, scale1 * 0.241895833616759), 0, 0, 0);
            poly179.AddVertexAt(2, new Point2d(scale1 * 0.0602374991727341, scale1 * 0.240072362743458), 0, 0, 0);
            poly179.AddVertexAt(3, new Point2d(scale1 * 0.0616401690752726, scale1 * 0.237968357889651), 0, 0, 0);
            poly179.AddVertexAt(4, new Point2d(scale1 * 0.0646792871974391, scale1 * 0.237687823909144), 0, 0, 0);
            poly179.AddVertexAt(5, new Point2d(scale1 * 0.0681859619537855, scale1 * 0.23937102779219), 0, 0, 0);
            poly179.AddVertexAt(6, new Point2d(scale1 * 0.0681859619537855, scale1 * 0.241615299636252), 0, 0, 0);
            poly179.AddVertexAt(7, new Point2d(scale1 * 0.065848178782888, scale1 * 0.244140105460821), 0, 0, 0);
            poly179.AddVertexAt(8, new Point2d(scale1 * 0.062107725709452, scale1 * 0.243438770509551), 0, 0, 0);
            poly179.Closed = true;
            poly179.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly179.Layer = "0";
            poly179.Color = color_GP;
            poly179.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly179);
            Polyline poly180 = new Polyline();
            poly180.AddVertexAt(0, new Point2d(scale1 * 0.0165209538769528, scale1 * 0.248628649148944), 0, 0, 0);
            poly180.AddVertexAt(1, new Point2d(scale1 * 0.0230667467554657, scale1 * 0.250452120022245), 0, 0, 0);
            poly180.AddVertexAt(2, new Point2d(scale1 * 0.0359245541954014, scale1 * 0.249750785070974), 0, 0, 0);
            poly180.AddVertexAt(3, new Point2d(scale1 * 0.0424703470739138, scale1 * 0.248628649148944), 0, 0, 0);
            poly180.AddVertexAt(4, new Point2d(scale1 * 0.0436392386593627, scale1 * 0.246805178275643), 0, 0, 0);
            poly180.AddVertexAt(5, new Point2d(scale1 * 0.0438730169764525, scale1 * 0.244981707402343), 0, 0, 0);
            poly180.AddVertexAt(6, new Point2d(scale1 * 0.0441067952935421, scale1 * 0.242737435558282), 0, 0, 0);
            poly180.AddVertexAt(7, new Point2d(scale1 * 0.0459770218302602, scale1 * 0.240773697694727), 0, 0, 0);
            poly180.AddVertexAt(8, new Point2d(scale1 * 0.0445743519277215, scale1 * 0.237407289928637), 0, 0, 0);
            poly180.AddVertexAt(9, new Point2d(scale1 * 0.0394312289517476, scale1 * 0.235583819055336), 0, 0, 0);
            poly180.AddVertexAt(10, new Point2d(scale1 * 0.0293787613168888, scale1 * 0.235303285074829), 0, 0, 0);
            poly180.AddVertexAt(11, new Point2d(scale1 * 0.024469416658004, scale1 * 0.236144887016352), 0, 0, 0);
            poly180.AddVertexAt(12, new Point2d(scale1 * 0.020962741901658, scale1 * 0.237828090899398), 0, 0, 0);
            poly180.AddVertexAt(13, new Point2d(scale1 * 0.0176898454624017, scale1 * 0.241054231675236), 0, 0, 0);
            poly180.AddVertexAt(14, new Point2d(scale1 * 0.0158196189256838, scale1 * 0.244420639441328), 0, 0, 0);
            poly180.Closed = true;
            poly180.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));
            poly180.Layer = "0";
            poly180.Color = color_GP;
            poly180.LineWeight = LineWeight.LineWeight000;
            bltrec1.AppendEntity(poly180);

            #endregion

            string nume_hatch = "ANSI31";
            double hatch_scale1 = 0.045 * scale1;
            double hatch_angle1 = 0;


            Polyline poly2 = new Polyline();
            poly2.AddVertexAt(0, new Point2d((poly1.GetPoint2dAt(0).X + poly1.GetPoint2dAt(1).X) / 2, (poly1.GetPoint2dAt(0).Y + poly1.GetPoint2dAt(1).Y) / 2), 0, 0, 0);
            poly2.AddVertexAt(1, poly1.GetPoint2dAt(1), 0, 0, 0);
            poly2.AddVertexAt(2, poly1.GetPoint2dAt(2), 0, 0, 0);
            poly2.AddVertexAt(3, new Point2d((poly1.GetPoint2dAt(2).X + poly1.GetPoint2dAt(3).X) / 2, (poly1.GetPoint2dAt(2).Y + poly1.GetPoint2dAt(3).Y) / 2), 0, 0, 0);
            poly2.Closed = true;


            BTrecord.AppendEntity(poly2);
            Trans1.AddNewlyCreatedDBObject(poly2, true);

            Hatch hatch1 = CreateHatch(poly2, nume_hatch, hatch_scale1, hatch_angle1 * Math.PI / 180);
            hatch1.Layer = "0";
            hatch1.LineWeight = LineWeight.LineWeight000;
            hatch1.Color = color_GP;
            bltrec1.AppendEntity(hatch1);



            poly2.Erase();

        }


        private void add_pattern_CL(BlockTableRecord bltrec1, double scale1, Polyline poly1, BlockTableRecord BTrecord, Autodesk.AutoCAD.DatabaseServices.Transaction Trans1)
        {
            string nume_hatch = "ANSI31";
            double hatch_scale = scale1 / 3;
            double hatch_angle = 0;
            Autodesk.AutoCAD.Colors.Color color1 = Autodesk.AutoCAD.Colors.Color.FromRgb(176, 71, 159);

            Polyline poly2 = new Polyline();
            poly2 = poly1.Clone() as Polyline;
            BTrecord.AppendEntity(poly2);
            Trans1.AddNewlyCreatedDBObject(poly2, true);

            Hatch hatch1 = CreateHatch(poly2, nume_hatch, hatch_scale, hatch_angle * Math.PI / 180);
            hatch1.Layer = "0";
            hatch1.LineWeight = LineWeight.LineWeight000;
            hatch1.Color = color1;
            bltrec1.AppendEntity(hatch1);
            poly2.Erase();

        }



        private void add_pattern_SC(BlockTableRecord bltrec1, double scale1, double graph_vexag, double stick_vexag, Polyline poly1, BlockTableRecord BTrecord, Autodesk.AutoCAD.DatabaseServices.Transaction Trans1)
        {
            Autodesk.AutoCAD.Colors.Color color1 = Autodesk.AutoCAD.Colors.Color.FromRgb(127, 95, 93);

            int nr_rows = 0;
            int nr_col = 0;

            double Xmax = 0.177187500001552 * scale1;
            double Ymax = 0.0645833333333333 * scale1;
            double r1 = 0.0015 * scale1;



            double x1 = poly1.GetPoint2dAt(3).X;
            double y1 = poly1.GetPoint2dAt(3).Y;
            double stick_width = poly1.GetPoint2dAt(1).X - poly1.GetPoint2dAt(0).X;
            double stick_height = poly1.GetPoint2dAt(1).Y - poly1.GetPoint2dAt(2).Y;

            if (stick_height < Ymax)
            {
                nr_rows = 1;
            }
            else
            {
                double nr2 = Math.Floor(stick_height / Ymax);
                nr_rows = Convert.ToInt32(nr2);
            }


            if (stick_width < Xmax)
            {
                nr_col = 1;
            }
            else
            {
                double nr2 = Math.Floor(stick_width / Xmax);
                nr_col = Convert.ToInt32(nr2);
            }








            for (int m = 0; m < nr_col; ++m)
            {
                for (int n = 0; n < nr_rows; ++n)
                {

                    double x2 = x1;
                    double y2 = y1;



                    Circle cerc1 = new Circle(new Point3d(0.0771875000015522 * scale1 + m * Xmax, 0.00208333333333333 * scale1 + n * Ymax, 0), Vector3d.ZAxis, r1);
                    cerc1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(x2, y2, 0))));
                    BTrecord.AppendEntity(cerc1);
                    Trans1.AddNewlyCreatedDBObject(cerc1, true);
                    Hatch hatch1 = CreateHatch(cerc1, "SOLID", 1, 0);
                    hatch1.Layer = "0";
                    hatch1.LineWeight = LineWeight.LineWeight000;
                    hatch1.Color = color1;
                    bltrec1.AppendEntity(hatch1);
                    cerc1.Erase();

                    cerc1 = new Circle(new Point3d(0.108497459879921 * scale1 + m * Xmax, 0.0111742424157759 * scale1 + n * Ymax, 0), Vector3d.ZAxis, r1);
                    cerc1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(x2, y2, 0))));
                    BTrecord.AppendEntity(cerc1);
                    Trans1.AddNewlyCreatedDBObject(cerc1, true);
                    hatch1 = CreateHatch(cerc1, "SOLID", 1, 0);
                    hatch1.Layer = "0";
                    hatch1.LineWeight = LineWeight.LineWeight000;
                    hatch1.Color = color1;
                    bltrec1.AppendEntity(hatch1);
                    cerc1.Erase();

                    cerc1 = new Circle(new Point3d(0.139687500001552 * scale1 + m * Xmax, 0.0395833333333333 * scale1 + n * Ymax, 0), Vector3d.ZAxis, r1);
                    cerc1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(x2, y2, 0))));
                    BTrecord.AppendEntity(cerc1);
                    Trans1.AddNewlyCreatedDBObject(cerc1, true);
                    hatch1 = CreateHatch(cerc1, "SOLID", 1, 0);
                    hatch1.Layer = "0";
                    hatch1.LineWeight = LineWeight.LineWeight000;
                    hatch1.Color = color1;
                    bltrec1.AppendEntity(hatch1);
                    cerc1.Erase();

                    cerc1 = new Circle(new Point3d(0.177187500001552 * scale1 + m * Xmax, 0.0520833333333333 * scale1 + n * Ymax, 0), Vector3d.ZAxis, r1);
                    cerc1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(x2, y2, 0))));
                    BTrecord.AppendEntity(cerc1);
                    Trans1.AddNewlyCreatedDBObject(cerc1, true);
                    hatch1 = CreateHatch(cerc1, "SOLID", 1, 0);
                    hatch1.Layer = "0";
                    hatch1.LineWeight = LineWeight.LineWeight000;
                    hatch1.Color = color1;
                    bltrec1.AppendEntity(hatch1);
                    cerc1.Erase();

                    cerc1 = new Circle(new Point3d(0.164687500001552 * scale1 + m * Xmax, 0.0145833333333333 * scale1 + n * Ymax, 0), Vector3d.ZAxis, r1);
                    cerc1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(x2, y2, 0))));
                    BTrecord.AppendEntity(cerc1);
                    Trans1.AddNewlyCreatedDBObject(cerc1, true);
                    hatch1 = CreateHatch(cerc1, "SOLID", 1, 0);
                    hatch1.Layer = "0";
                    hatch1.LineWeight = LineWeight.LineWeight000;
                    hatch1.Color = color1;
                    bltrec1.AppendEntity(hatch1);
                    cerc1.Erase();

                    cerc1 = new Circle(new Point3d(0.102187500001552 * scale1 + m * Xmax, 0.0395833333333333 * scale1 + n * Ymax, 0), Vector3d.ZAxis, r1);
                    cerc1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(x2, y2, 0))));
                    BTrecord.AppendEntity(cerc1);
                    Trans1.AddNewlyCreatedDBObject(cerc1, true);
                    hatch1 = CreateHatch(cerc1, "SOLID", 1, 0);
                    hatch1.Layer = "0";
                    hatch1.LineWeight = LineWeight.LineWeight000;
                    hatch1.Color = color1;
                    bltrec1.AppendEntity(hatch1);
                    cerc1.Erase();

                    cerc1 = new Circle(new Point3d(0.114687500001552 * scale1 + m * Xmax, 0.0645833333333333 * scale1 + n * Ymax, 0), Vector3d.ZAxis, r1);
                    cerc1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(x2, y2, 0))));
                    BTrecord.AppendEntity(cerc1);
                    Trans1.AddNewlyCreatedDBObject(cerc1, true);
                    hatch1 = CreateHatch(cerc1, "SOLID", 1, 0);
                    hatch1.Layer = "0";
                    hatch1.LineWeight = LineWeight.LineWeight000;
                    hatch1.Color = color1;
                    bltrec1.AppendEntity(hatch1);
                    cerc1.Erase();

                    cerc1 = new Circle(new Point3d(0.0857900086014221 * scale1 + m * Xmax, 0.0532196969725192 * scale1 + n * Ymax, 0), Vector3d.ZAxis, r1);
                    cerc1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(x2, y2, 0))));
                    BTrecord.AppendEntity(cerc1);
                    Trans1.AddNewlyCreatedDBObject(cerc1, true);
                    hatch1 = CreateHatch(cerc1, "SOLID", 1, 0);
                    hatch1.Layer = "0";
                    hatch1.LineWeight = LineWeight.LineWeight000;
                    hatch1.Color = color1;
                    bltrec1.AppendEntity(hatch1);
                    cerc1.Erase();

                    cerc1 = new Circle(new Point3d(0.0521875000015522 * scale1 + m * Xmax, 0.0270833333333333 * scale1 + n * Ymax, 0), Vector3d.ZAxis, r1);
                    cerc1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(x2, y2, 0))));
                    BTrecord.AppendEntity(cerc1);
                    Trans1.AddNewlyCreatedDBObject(cerc1, true);
                    hatch1 = CreateHatch(cerc1, "SOLID", 1, 0);
                    hatch1.Layer = "0";
                    hatch1.LineWeight = LineWeight.LineWeight000;
                    hatch1.Color = color1;
                    bltrec1.AppendEntity(hatch1);
                    cerc1.Erase();

                    cerc1 = new Circle(new Point3d(0.0146875000015522 * scale1 + m * Xmax, 0.0145833333333333 * scale1 + n * Ymax, 0), Vector3d.ZAxis, r1);
                    cerc1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(x2, y2, 0))));
                    BTrecord.AppendEntity(cerc1);
                    Trans1.AddNewlyCreatedDBObject(cerc1, true);
                    hatch1 = CreateHatch(cerc1, "SOLID", 1, 0);
                    hatch1.Layer = "0";
                    hatch1.LineWeight = LineWeight.LineWeight000;
                    hatch1.Color = color1;
                    bltrec1.AppendEntity(hatch1);
                    cerc1.Erase();

                    cerc1 = new Circle(new Point3d(0.0271875000015522 * scale1 + m * Xmax, 0.0395833333333333 * scale1 + n * Ymax, 0), Vector3d.ZAxis, r1);
                    cerc1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(x2, y2, 0))));
                    BTrecord.AppendEntity(cerc1);
                    Trans1.AddNewlyCreatedDBObject(cerc1, true);
                    hatch1 = CreateHatch(cerc1, "SOLID", 1, 0);
                    hatch1.Layer = "0";
                    hatch1.LineWeight = LineWeight.LineWeight000;
                    hatch1.Color = color1;
                    bltrec1.AppendEntity(hatch1);
                    cerc1.Erase();

                    cerc1 = new Circle(new Point3d(0.0396875000015522 * scale1 + m * Xmax, 0.0645833333333333 * scale1 + n * Ymax, 0), Vector3d.ZAxis, r1);
                    cerc1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(x2, y2, 0))));
                    BTrecord.AppendEntity(cerc1);
                    Trans1.AddNewlyCreatedDBObject(cerc1, true);
                    hatch1 = CreateHatch(cerc1, "SOLID", 1, 0);
                    hatch1.Layer = "0";
                    hatch1.LineWeight = LineWeight.LineWeight000;
                    hatch1.Color = color1;
                    bltrec1.AppendEntity(hatch1);
                    cerc1.Erase();

                    cerc1 = new Circle(new Point3d(0.0021875000015522 * scale1 + m * Xmax, 0.0520833333333333 * scale1 + n * Ymax, 0), Vector3d.ZAxis, r1);
                    cerc1.TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(new Point3d(x2, y2, 0))));
                    BTrecord.AppendEntity(cerc1);
                    Trans1.AddNewlyCreatedDBObject(cerc1, true);
                    hatch1 = CreateHatch(cerc1, "SOLID", 1, 0);
                    hatch1.Layer = "0";
                    hatch1.LineWeight = LineWeight.LineWeight000;
                    hatch1.Color = color1;
                    bltrec1.AppendEntity(hatch1);
                    cerc1.Erase();





                }
            }

            string nume_hatch = "ANSI31";
            double hatch_scale = scale1 / 2;
            double hatch_angle = 0;


            Polyline poly2 = new Polyline();
            poly2 = poly1.Clone() as Polyline;
            BTrecord.AppendEntity(poly2);
            Trans1.AddNewlyCreatedDBObject(poly2, true);

            Hatch hatch2 = CreateHatch(poly2, nume_hatch, hatch_scale, hatch_angle * Math.PI / 180);
            hatch2.Layer = "0";
            hatch2.LineWeight = LineWeight.LineWeight000;
            hatch2.Color = color1;
            bltrec1.AppendEntity(hatch2);
            poly2.Erase();


        }




        private Polyline get_poly_sym(double size1)
        {
            Polyline poly_sym = new Polyline();
            poly_sym.AddVertexAt(0, new Point2d(size1 * -0.38268343236509, size1 * 0.923879532511285), 0, 0, 0);
            poly_sym.AddVertexAt(1, new Point2d(size1 * -1.40363320911291, size1 * 1.40363320911291), 0, 0, 0);
            poly_sym.AddVertexAt(2, new Point2d(size1 * -0.923879532511285, size1 * 0.38268343236509), 0, 0, 0);
            poly_sym.AddVertexAt(3, new Point2d(size1 * -1.98503712092475, size1 * 0), 0, 0, 0);
            poly_sym.AddVertexAt(4, new Point2d(size1 * -0.923879532511285, size1 * -0.382683432365091), 0, 0, 0);
            poly_sym.AddVertexAt(5, new Point2d(size1 * -1.40363320911291, size1 * -1.40363320911291), 0, 0, 0);
            poly_sym.AddVertexAt(6, new Point2d(size1 * -0.38268343236509, size1 * -0.923879532511285), 0, 0, 0);
            poly_sym.AddVertexAt(7, new Point2d(size1 * 0, size1 * -1.98503712092474), 0, 0, 0);
            poly_sym.AddVertexAt(8, new Point2d(size1 * 0.38268343236509, size1 * -0.923879532511285), 0, 0, 0);
            poly_sym.AddVertexAt(9, new Point2d(size1 * 1.40363320911291, size1 * -1.40363320911291), 0, 0, 0);
            poly_sym.AddVertexAt(10, new Point2d(size1 * 0.923879532511285, size1 * -0.38268343236509), 0, 0, 0);
            poly_sym.AddVertexAt(11, new Point2d(size1 * 1.98503712092475, size1 * 0), 0, 0, 0);
            poly_sym.AddVertexAt(12, new Point2d(size1 * 0.923879532511285, size1 * 0.38268343236509), 0, 0, 0);
            poly_sym.AddVertexAt(13, new Point2d(size1 * 1.40363320911291, size1 * 1.40363320911291), 0, 0, 0);
            poly_sym.AddVertexAt(14, new Point2d(size1 * 0.38268343236509, size1 * 0.923879532511285), 0, 0, 0);
            poly_sym.AddVertexAt(15, new Point2d(size1 * 0, size1 * 1.98503712092475), 0, 0, 0);
            poly_sym.AddVertexAt(16, new Point2d(size1 * -0.38268343236509, size1 * 0.92387953251129), 0, 0, 0);

            return poly_sym;


        }

        private Polyline get_poly_gm1(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(scale1 * 0.00558759536749373, scale1 * 0.0133960036188364), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(scale1 * 0.00231399374315515, scale1 * 0.0204669860191643), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(scale1 * 0.00354159435179705, scale1 * 0.0293057140149176), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(scale1 * 0.0149992000311613, scale1 * 0.033135829474777), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(scale1 * 0.0375052111883027, scale1 * 0.0316627081483603), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(scale1 * 0.048962816867667, scale1 * 0.0293057140149176), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(scale1 * 0.0510088178833637, scale1 * 0.0254755985550582), 0, 0, 0);
            poly1.AddVertexAt(7, new Point2d(scale1 * 0.0514180180849507, scale1 * 0.0216454830765724), 0, 0, 0);
            poly1.AddVertexAt(8, new Point2d(scale1 * 0.0518272182865378, scale1 * 0.0169314948096871), 0, 0, 0);
            poly1.AddVertexAt(9, new Point2d(scale1 * 0.0551008199108765, scale1 * 0.0128067550808191), 0, 0, 0);
            poly1.AddVertexAt(10, new Point2d(scale1 * 0.0526456186935927, scale1 * 0.00573577268049121), 0, 0, 0);
            poly1.AddVertexAt(11, new Point2d(scale1 * 0.0436432142315122, scale1 * 0.00190565722063184), 0, 0, 0);
            poly1.AddVertexAt(12, new Point2d(scale1 * 0.0260476055089385, scale1 * 0.00131640868261456), 0, 0, 0);
            poly1.AddVertexAt(13, new Point2d(scale1 * 0.0174544012484452, scale1 * 0.00308415427803993), 0, 0, 0);
            poly1.AddVertexAt(14, new Point2d(scale1 * 0.0113163982052356, scale1 * 0.00661964548751712), 0, 0, 0);
            poly1.Closed = true;

            return poly1;


        }

        private Polyline get_poly_gm2(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(scale1 * 0.0784252314711922, scale1 * 0.0151637492142618), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(scale1 * 0.0800620322853017, scale1 * 0.0113336337544024), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(scale1 * 0.0825172335025855, scale1 * 0.00691426975652575), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(scale1 * 0.0878368361387402, scale1 * 0.00632502121850848), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(scale1 * 0.0939748391819497, scale1 * 0.00986051240935922), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(scale1 * 0.0939748391819497, scale1 * 0.0145745006762445), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(scale1 * 0.0898828371505563, scale1 * 0.0198777374811471), 0, 0, 0);
            poly1.AddVertexAt(7, new Point2d(scale1 * 0.0833356339057597, scale1 * 0.0184046161547303), 0, 0, 0);
            poly1.Closed = true;

            return poly1;
        }

        private Polyline get_poly_gm3(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(scale1 * 0.143897263926919, scale1 * 0.0181099918857217), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(scale1 * 0.141032862508049, scale1 * 0.0154583734832704), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(scale1 * 0.133667258856197, scale1 * 0.0125121308118105), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(scale1 * 0.126710855405933, scale1 * 0.0128067550808191), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(scale1 * 0.121391252769778, scale1 * 0.0154583734832704), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(scale1 * 0.115662449932036, scale1 * 0.0181099918857217), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(scale1 * 0.110342847292001, scale1 * 0.020761610288173), 0, 0, 0);
            poly1.AddVertexAt(7, new Point2d(scale1 * 0.107069245671543, scale1 * 0.0257702228054404), 0, 0, 0);
            poly1.AddVertexAt(8, new Point2d(scale1 * 0.107069245671543, scale1 * 0.0293057140149176), 0, 0, 0);
            poly1.AddVertexAt(9, new Point2d(scale1 * 0.113207248714752, scale1 * 0.0369659449532628), 0, 0, 0);
            poly1.AddVertexAt(10, new Point2d(scale1 * 0.125074054595704, scale1 * 0.0402068118751049), 0, 0, 0);
            poly1.AddVertexAt(11, new Point2d(scale1 * 0.136940860476655, scale1 * 0.0369659449532628), 0, 0, 0);
            poly1.AddVertexAt(12, new Point2d(scale1 * 0.143488063721452, scale1 * 0.0293057140149176), 0, 0, 0);
            poly1.Closed = true;

            return poly1;
        }

        private Polyline get_poly_gm4(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(scale1 * 0.0616480231556732, scale1 * 0.0561165222711861), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(scale1 * 0.0534640191006475, scale1 * 0.0581788921356201), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(scale1 * 0.0469168158558508, scale1 * 0.0634821289405227), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(scale1 * 0.0285028067262222, scale1 * 0.0720262326672673), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(scale1 * 0.018682001857087, scale1 * 0.0802757121436298), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(scale1 * 0.0211372030743708, scale1 * 0.0841058276034892), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(scale1 * 0.0358684103780737, scale1 * 0.0917660585418344), 0, 0, 0);
            poly1.AddVertexAt(7, new Point2d(scale1 * 0.0473260160574378, scale1 * 0.0932391798682511), 0, 0, 0);
            poly1.AddVertexAt(8, new Point2d(scale1 * 0.0546916197092893, scale1 * 0.0905875614657998), 0, 0, 0);
            poly1.AddVertexAt(9, new Point2d(scale1 * 0.064512424574544, scale1 * 0.0879359430633485), 0, 0, 0);
            poly1.AddVertexAt(10, new Point2d(scale1 * 0.0722872284279825, scale1 * 0.0864628217369318), 0, 0, 0);
            poly1.AddVertexAt(11, new Point2d(scale1 * 0.076788430660963, scale1 * 0.0835165790654719), 0, 0, 0);
            poly1.AddVertexAt(12, new Point2d(scale1 * 0.090292037356024, scale1 * 0.0799810878746212), 0, 0, 0);
            poly1.AddVertexAt(13, new Point2d(scale1 * 0.0960208401937658, scale1 * 0.070258487071842), 0, 0, 0);
            poly1.AddVertexAt(14, new Point2d(scale1 * 0.090292037356024, scale1 * 0.0608305105380714), 0, 0, 0);
            poly1.AddVertexAt(15, new Point2d(scale1 * 0.076788430660963, scale1 * 0.0572950193472207), 0, 0, 0);
            poly1.Closed = true;


            return poly1;
        }

        private Polyline get_poly_gm5(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(scale1 * 0.101947919845891, scale1 * 0.0574289394542575), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(scale1 * 0.0926560261752457, scale1 * 0.0526078150980175), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(scale1 * 0.0907976474418925, scale1 * 0.0461796492896974), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(scale1 * 0.100089541112538, scale1 * 0.0443047675862908), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(scale1 * 0.11161148926088, scale1 * 0.0464474895223975), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(scale1 * 0.118301652705607, scale1 * 0.0520721346139908), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(scale1 * 0.112726516501668, scale1 * 0.0571610992215574), 0, 0, 0);
            poly1.Closed = true;



            return poly1;
        }

        private Polyline get_poly_gm6(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(scale1 * 0.154892033373471, scale1 * 0.0600269898213446), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(scale1 * 0.155301233575059, scale1 * 0.0523667588829994), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(scale1 * 0.156938034389168, scale1 * 0.0464742735400796), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(scale1 * 0.164303638041019, scale1 * 0.0447065279446542), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(scale1 * 0.17166924168899, scale1 * 0.0458850250206888), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(scale1 * 0.1778072447322, scale1 * 0.0503043890185654), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(scale1 * 0.17166924168899, scale1 * 0.0606162383407354), 0, 0, 0);
            poly1.AddVertexAt(7, new Point2d(scale1 * 0.157699281862006, scale1 * 0.0673390284366906), 0, 0, 0);

            poly1.Closed = true;



            return poly1;
        }

        private Polyline get_poly_gm7(double scale1)
        {
            Polyline poly1 = new Polyline();

            poly1.AddVertexAt(0, new Point2d(scale1 * 0.122885093383957, scale1 * 0.113970014620572), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(scale1 * 0.119202291558031, scale1 * 0.103068916760385), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(scale1 * 0.119202291558031, scale1 * 0.0977656799554825), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(scale1 * 0.117974690949389, scale1 * 0.0862753335572779), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(scale1 * 0.124112693992599, scale1 * 0.0812667210400105), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(scale1 * 0.132705898253092, scale1 * 0.0794989754259586), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(scale1 * 0.152756708192949, scale1 * 0.0803828482329845), 0, 0, 0);
            poly1.AddVertexAt(7, new Point2d(scale1 * 0.158076310829104, scale1 * 0.0836237151734531), 0, 0, 0);
            poly1.AddVertexAt(8, new Point2d(scale1 * 0.159713111639333, scale1 * 0.0915785703621805), 0, 0, 0);
            poly1.AddVertexAt(9, new Point2d(scale1 * 0.154393509003178, scale1 * 0.0959979343600571), 0, 0, 0);
            poly1.AddVertexAt(10, new Point2d(scale1 * 0.149483106568611, scale1 * 0.101006546895951), 0, 0, 0);
            poly1.AddVertexAt(11, new Point2d(scale1 * 0.145391104541098, scale1 * 0.108372153565288), 0, 0, 0);
            poly1.AddVertexAt(12, new Point2d(scale1 * 0.14211750291676, scale1 * 0.113675390351564), 0, 0, 0);
            poly1.AddVertexAt(13, new Point2d(scale1 * 0.134751899268789, scale1 * 0.115737760215998), 0, 0, 0);
            poly1.AddVertexAt(14, new Point2d(scale1 * 0.128613896225579, scale1 * 0.117800130099058), 0, 0, 0);

            poly1.Closed = true;


            return poly1;
        }

        private Polyline get_poly_gm8(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(scale1 * 0.163227971566667, scale1 * 0.1062294316), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(scale1 * 0.165683172783333, scale1 * 0.1023993162), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(scale1 * 0.172639576166667, scale1 * 0.1009261948), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(scale1 * 0.177959178833333, scale1 * 0.1035778132), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(scale1 * 0.177959178833333, scale1 * 0.1091756743), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(scale1 * 0.167319973666667, scale1 * 0.1106487956), 0, 0, 0);
            poly1.Closed = true;

            return poly1;
        }

        private Polyline get_poly_gm9(double scale1)
        {
            Polyline poly1 = new Polyline();

            poly1.AddVertexAt(0, new Point2d(scale1 * 0.09029203735, scale1 * 0.09942628946), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(scale1 * 0.08988283715, scale1 * 0.09530154973), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(scale1 * 0.0943840393833333, scale1 * 0.09206068279), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(scale1 * 0.100112842216667, scale1 * 0.0929445556), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(scale1 * 0.101340442833333, scale1 * 0.0973639196), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(scale1 * 0.100522042433333, scale1 * 0.1008994108), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(scale1 * 0.0935656389833333, scale1 * 0.1026671564), 0, 0, 0);

            poly1.Closed = true;


            return poly1;
        }

        private Polyline get_poly_gm10(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(scale1 * 0.0690136268, scale1 * 0.1100327631), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(scale1 * 0.0767884306666667, scale1 * 0.1082650175), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(scale1 * 0.0788344316833333, scale1 * 0.1053187748), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(scale1 * 0.0767884306666667, scale1 * 0.1014886593), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(scale1 * 0.0735148290333333, scale1 * 0.1011940351), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(scale1 * 0.0694228270166667, scale1 * 0.1026671564), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(scale1 * 0.0661492253833333, scale1 * 0.1067918961), 0, 0, 0);
            poly1.Closed = true;




            return poly1;
        }

        private Polyline get_poly_gp1(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(78 * scale1 * 8.50048350868635E-05, 78 * scale1 * 0.000179175549322003), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(78 * scale1 * 0.000168925256606258, 78 * scale1 * 0.000202553381030987), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(78 * scale1 * 0.000333768941733638, 78 * scale1 * 0.000193561907296748), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(78 * scale1 * 0.00041768936325303, 78 * scale1 * 0.000179175549322003), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(78 * scale1 * 0.000432675152810065, 78 * scale1 * 0.00015579771761302), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(78 * scale1 * 0.000435672310721473, 78 * scale1 * 0.000132419885904037), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(78 * scale1 * 0.00043866946863288, 78 * scale1 * 0.000103647169954546), 0, 0, 0);
            poly1.AddVertexAt(7, new Point2d(78 * scale1 * 0.000462646731924135, 78 * scale1 * 7.84710434987101E-05), 0, 0, 0);
            poly1.AddVertexAt(8, new Point2d(78 * scale1 * 0.000444663784455693, 78 * scale1 * 3.53119695744745E-05), 0, 0, 0);
            poly1.AddVertexAt(9, new Point2d(78 * scale1 * 0.000378726310404742, 78 * scale1 * 1.19341378654911E-05), 0, 0, 0);
            poly1.AddVertexAt(10, new Point2d(78 * scale1 * 0.000249848520214245, 78 * scale1 * 8.33754837181126E-06), 0, 0, 0);
            poly1.AddVertexAt(11, new Point2d(78 * scale1 * 0.000186908204074698, 78 * scale1 * 1.91273168528765E-05), 0, 0, 0);
            poly1.AddVertexAt(12, new Point2d(78 * scale1 * 0.000141950835403595, 78 * scale1 * 4.07068538150071E-05), 0, 0, 0);
            poly1.AddVertexAt(13, new Point2d(78 * scale1 * 9.99906246438985E-05, 78 * scale1 * 8.20676329924156E-05), 0, 0, 0);
            poly1.AddVertexAt(14, new Point2d(78 * scale1 * 7.60133613526432E-05, 78 * scale1 * 0.000125226706916677), 0, 0, 0);

            poly1.Closed = true;

            return poly1;
        }


        private Polyline get_poly_gp2(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(78 * scale1 * 0.000669450627811212, 78 * scale1 * 0.000112638643688759), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(78 * scale1 * 0.00063348473287433, 78 * scale1 * 9.28574014734809E-05), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(78 * scale1 * 0.000645473364519957, 78 * scale1 * 6.94795697644975E-05), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(78 * scale1 * 0.000663456311988398, 78 * scale1 * 4.25051485618343E-05), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(78 * scale1 * 0.000702419364836687, 78 * scale1 * 3.89085590681544E-05), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(78 * scale1 * 0.000747376733507792, 78 * scale1 * 0.000060488096030285), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(78 * scale1 * 0.000747376733507792, 78 * scale1 * 8.92608119798009E-05), 0, 0, 0);
            poly1.AddVertexAt(7, new Point2d(78 * scale1 * 0.000717405154393722, 78 * scale1 * 0.000121630117422997), 0, 0, 0);
            poly1.AddVertexAt(8, new Point2d(78 * scale1 * 0.000669450627811212, 78 * scale1 * 0.000112638643688759), 0, 0, 0);

            poly1.Closed = true;
            return poly1;
        }

        private Polyline get_poly_gp3(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(78 * scale1 * 0.00111003284078803, 78 * scale1 * 0.000179175549322003), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(78 * scale1 * 0.00111302999869943, 78 * scale1 * 0.000110840348941906), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(78 * scale1 * 0.00109204989331959, 78 * scale1 * 9.46556962203336E-05), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(78 * scale1 * 0.00103810105091426, 78 * scale1 * 7.66727487518829E-05), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(78 * scale1 * 0.000987149366420343, 78 * scale1 * 7.84710434987101E-05), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(78 * scale1 * 0.000948186313572055, 78 * scale1 * 9.46556962203336E-05), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(78 * scale1 * 0.000906226102812358, 78 * scale1 * 0.000110840348941906), 0, 0, 0);
            poly1.AddVertexAt(7, new Point2d(78 * scale1 * 0.000867263049964068, 78 * scale1 * 0.000127025001663529), 0, 0, 0);
            poly1.AddVertexAt(8, new Point2d(78 * scale1 * 0.000843285786672813, 78 * scale1 * 0.000157596012359873), 0, 0, 0);
            poly1.AddVertexAt(9, new Point2d(78 * scale1 * 0.000843285786672813, 78 * scale1 * 0.000179175549322003), 0, 0, 0);
            poly1.AddVertexAt(10, new Point2d(78 * scale1 * 0.000888243155343917, 78 * scale1 * 0.000225931212739944), 0, 0, 0);
            poly1.AddVertexAt(11, new Point2d(78 * scale1 * 0.000975160734774717, 78 * scale1 * 0.000245712454955222), 0, 0, 0);
            poly1.AddVertexAt(12, new Point2d(78 * scale1 * 0.00106207831420552, 78 * scale1 * 0.000225931212739944), 0, 0, 0);

            poly1.Closed = true;

            return poly1;
        }


        private Polyline get_poly_gp4(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(78 * scale1 * 0.000884722066303067, 78 * scale1 * 0.000349196143568983), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(78 * scale1 * 0.000805774998294947, 78 * scale1 * 0.000350830956975229), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(78 * scale1 * 0.000737717181046563, 78 * scale1 * 0.000321404315663311), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(78 * scale1 * 0.000724105617596905, 78 * scale1 * 0.000282168793913944), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(78 * scale1 * 0.000792163434845288, 78 * scale1 * 0.000270725100070426), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(78 * scale1 * 0.000876555128233278, 78 * scale1 * 0.000283803607320164), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(78 * scale1 * 0.000925556756652102, 78 * scale1 * 0.000318134688850819), 0, 0, 0);
            poly1.AddVertexAt(7, new Point2d(78 * scale1 * 0.000884722066303067, 78 * scale1 * 0.000349196143568983), 0, 0, 0);

            poly1.Closed = true;
            return poly1;
        }

        private Polyline get_poly_gp5(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(78 * scale1 * 0.000621496101228702, 78 * scale1 * 0.000510061782741322), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(78 * scale1 * 0.00072040231230513, 78 * scale1 * 0.000488482245779191), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(78 * scale1 * 0.000762362523064825, 78 * scale1 * 0.000429138519133332), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(78 * scale1 * 0.00072040231230513, 78 * scale1 * 0.000371593087234326), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(78 * scale1 * 0.000621496101228702, 78 * scale1 * 0.000350013550272195), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(78 * scale1 * 0.000510601258506647, 78 * scale1 * 0.00034282037128481), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(78 * scale1 * 0.000450658100278508, 78 * scale1 * 0.000355408434512728), 0, 0, 0);
            poly1.AddVertexAt(7, new Point2d(78 * scale1 * 0.000402703573695997, 78 * scale1 * 0.000387777739955924), 0, 0, 0);
            poly1.AddVertexAt(8, new Point2d(78 * scale1 * 0.000267831467682685, 78 * scale1 * 0.000439928287614397), 0, 0, 0);
            poly1.AddVertexAt(9, new Point2d(78 * scale1 * 0.00019589967780892, 78 * scale1 * 0.000490280540526044), 0, 0, 0);
            poly1.AddVertexAt(10, new Point2d(78 * scale1 * 0.000213882625277362, 78 * scale1 * 0.000513658372235027), 0, 0, 0);
            poly1.AddVertexAt(11, new Point2d(78 * scale1 * 0.00032178031008801, 78 * scale1 * 0.000560414035652943), 0, 0, 0);
            poly1.AddVertexAt(12, new Point2d(78 * scale1 * 0.000405700731607403, 78 * scale1 * 0.000569405509387181), 0, 0, 0);
            poly1.AddVertexAt(13, new Point2d(78 * scale1 * 0.000459649574012728, 78 * scale1 * 0.000553220856665583), 0, 0, 0);
            poly1.AddVertexAt(14, new Point2d(78 * scale1 * 0.000531581363886495, 78 * scale1 * 0.000537036203943985), 0, 0, 0);
            poly1.AddVertexAt(15, new Point2d(78 * scale1 * 0.000588527364203225, 78 * scale1 * 0.000528044730209747), 0, 0, 0);

            poly1.Closed = true;
            return poly1;
        }

        private Polyline get_poly_gp6(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(78 * scale1 * 0.00100108805414107, 78 * scale1 * 0.000719317898737644), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(78 * scale1 * 0.000959127843381368, 78 * scale1 * 0.000695940067028686), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(78 * scale1 * 0.000932153422178707, 78 * scale1 * 0.000629403161395441), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(78 * scale1 * 0.000932153422178707, 78 * scale1 * 0.000597033855952245), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(78 * scale1 * 0.000923161948444487, 78 * scale1 * 0.000526900360825321), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(78 * scale1 * 0.00096811931711559, 78 * scale1 * 0.000496329350128978), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(78 * scale1 * 0.00103105963325514, 78 * scale1 * 0.000485539581647913), 0, 0, 0);
            poly1.AddVertexAt(7, new Point2d(78 * scale1 * 0.00117792037091407, 78 * scale1 * 0.000490934465888445), 0, 0, 0);
            poly1.AddVertexAt(8, new Point2d(78 * scale1 * 0.00121688342376236, 78 * scale1 * 0.000510715708103723), 0, 0, 0);
            poly1.AddVertexAt(9, new Point2d(78 * scale1 * 0.00122887205540799, 78 * scale1 * 0.000559269666268517), 0, 0, 0);
            poly1.AddVertexAt(10, new Point2d(78 * scale1 * 0.0011899090025597, 78 * scale1 * 0.00058624408747118), 0, 0, 0);
            poly1.AddVertexAt(11, new Point2d(78 * scale1 * 0.00115394310762282, 78 * scale1 * 0.000616815098167523), 0, 0, 0);
            poly1.AddVertexAt(12, new Point2d(78 * scale1 * 0.00112397152850875, 78 * scale1 * 0.000661772466838637), 0, 0, 0);
            poly1.AddVertexAt(13, new Point2d(78 * scale1 * 0.00109999426521749, 78 * scale1 * 0.000694141772281833), 0, 0, 0);
            poly1.AddVertexAt(14, new Point2d(78 * scale1 * 0.00104604542281217, 78 * scale1 * 0.000706729835509751), 0, 0, 0);
            poly1.AddVertexAt(15, new Point2d(78 * scale1 * 0.00100108805414107, 78 * scale1 * 0.000719317898737644), 0, 0, 0);
            poly1.Closed = true;
            return poly1;
        }

        private Polyline get_poly_gp7(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(78 * scale1 * 0.000015258737222475, 78 * scale1 * 0.000411319053005514), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(78 * scale1 * -5.30275554959253E-06, 78 * scale1 * 0.000366688647015613), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(78 * scale1 * -2.30559763818537E-06, 78 * scale1 * 0.000319932983597672), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(78 * scale1 * 9.68303400744238E-06, 78 * scale1 * 0.000283967088660771), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(78 * scale1 * 6.36318764127668E-05, 78 * scale1 * 0.000273177320179706), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(78 * scale1 * 0.000117580718818091, 78 * scale1 * 0.000280370499167091), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(78 * scale1 * 0.000162538087489195, 78 * scale1 * 0.000307344920369754), 0, 0, 0);
            poly1.AddVertexAt(7, new Point2d(78 * scale1 * 0.000117580718818091, 78 * scale1 * 0.000370285236509293), 0, 0, 0);
            poly1.AddVertexAt(8, new Point2d(78 * scale1 * 0.000015258737222475, 78 * scale1 * 0.000411319053005514), 0, 0, 0);
            poly1.Closed = true;
            return poly1;
        }

        private Polyline get_poly_gp8(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(78 * scale1 * 8.57248126521462E-05, 78 * scale1 * 0.000675668380791613), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(78 * scale1 * 0.000055753233538077, 78 * scale1 * 0.00064869395958895), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(78 * scale1 * 7.37361810065185E-05, 78 * scale1 * 0.000625316127879967), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(78 * scale1 * 0.000124687865500436, 78 * scale1 * 0.000616324654145754), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(78 * scale1 * 0.000163650918348726, 78 * scale1 * 0.000632509306867352), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(78 * scale1 * 0.000163650918348726, 78 * scale1 * 0.000666676907057375), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(78 * scale1 * 8.57248126521462E-05, 78 * scale1 * 0.000675668380791613), 0, 0, 0);
            poly1.Closed = true;
            return poly1;
        }

        private Polyline get_poly_gp9(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(78 * scale1 * 0.000744379575596385, 78 * scale1 * 0.000626950941286187), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(78 * scale1 * 0.00072040231230513, 78 * scale1 * 0.000607169699070909), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(78 * scale1 * 0.000717405154393722, 78 * scale1 * 0.000581993572615073), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(78 * scale1 * 0.000750373891419198, 78 * scale1 * 0.000562212330399796), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(78 * scale1 * 0.000792334102178895, 78 * scale1 * 0.000567607214640328), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(78 * scale1 * 0.000801325575913117, 78 * scale1 * 0.000594581635842991), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(78 * scale1 * 0.000795331260090303, 78 * scale1 * 0.000616161172805122), 0, 0, 0);
            poly1.AddVertexAt(7, new Point2d(78 * scale1 * 0.000744379575596385, 78 * scale1 * 0.000626950941286187), 0, 0, 0);
            poly1.Closed = true;
            return poly1;
        }

        private Polyline get_poly_gp10(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(78 * scale1 * 0.000621496101228702, 78 * scale1 * 0.000619757762298802), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(78 * scale1 * 0.000597518837937445, 78 * scale1 * 0.000617959467551975), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(78 * scale1 * 0.000567547258823377, 78 * scale1 * 0.000626950941286187), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(78 * scale1 * 0.000543569995532122, 78 * scale1 * 0.000652127067742023), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(78 * scale1 * 0.00056455010091197, 78 * scale1 * 0.000671908309957276), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(78 * scale1 * 0.000621496101228702, 78 * scale1 * 0.000661118541476236), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(78 * scale1 * 0.000636481890785737, 78 * scale1 * 0.000643135594007785), 0, 0, 0);

            poly1.Closed = true;
            return poly1;
        }


        private Polyline get_poly_mdst1(double scale1, double fac = 60)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(scale1 * 8.08414946636185 / fac, scale1 * 4 / fac), 0, 0.016667 * scale1, 0.016667 * scale1);
            poly1.AddVertexAt(1, new Point2d(scale1 * 9.18218045402318 / fac, scale1 * 7 / fac), 0, 0.016667 * scale1, 0.016667 * scale1);




            return poly1;
        }


        private Polyline get_poly_mdst2(double scale1, double fac = 60)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(scale1 * 9.90934027172625 / fac, scale1 * 0 / fac), 0, 0.016667 * scale1, 0.016667 * scale1);
            poly1.AddVertexAt(1, new Point2d(scale1 * 8.22549172234721 / fac, scale1 * 3.90670545771718 / fac), 0, 0.016667 * scale1, 0.016667 * scale1);
            poly1.AddVertexAt(2, new Point2d(scale1 * 5.04429773800075 / fac, scale1 * 6.00648824684322 / fac), 0, 0.016667 * scale1, 0.016667 * scale1);

            return poly1;
        }

        private Polyline get_poly_mdst3(double scale1, double fac = 60)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(scale1 * 0 / fac, scale1 * 0.39300299808383 / fac), 0, 0.016667 * scale1, 0.016667 * scale1);
            poly1.AddVertexAt(1, new Point2d(scale1 * 6.38135609426536 / fac, scale1 * 6.57110919244587 / fac), 0, 0.016667 * scale1, 0.016667 * scale1);


            return poly1;
        }

        private Polyline get_poly_gc1(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(50 * scale1 * 0.000560755223397328, 50 * scale1 * 0.000996363636363995), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(50 * scale1 * 0.001027421890064, 50 * scale1 * 0.00112636363636398), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(50 * scale1 * 0.00194408855673067, 50 * scale1 * 0.00107636363636395), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(50 * scale1 * 0.00241075522339733, 50 * scale1 * 0.000996363636363995), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(50 * scale1 * 0.00249408855673067, 50 * scale1 * 0.000866363636363943), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(50 * scale1 * 0.00251075522339733, 50 * scale1 * 0.000736363636363961), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(50 * scale1 * 0.002527421890064, 50 * scale1 * 0.000576363636363979), 0, 0, 0);
            poly1.AddVertexAt(7, new Point2d(50 * scale1 * 0.00266075522339733, 50 * scale1 * 0.00043636363636395), 0, 0, 0);
            poly1.AddVertexAt(8, new Point2d(50 * scale1 * 0.00256075522339733, 50 * scale1 * 0.000196363636363941), 0, 0, 0);
            poly1.AddVertexAt(9, new Point2d(50 * scale1 * 0.00219408855673067, 50 * scale1 * 6.63636363639597E-05), 0, 0, 0);
            poly1.AddVertexAt(10, new Point2d(50 * scale1 * 0.001477421890064, 50 * scale1 * 4.63636363639353E-05), 0, 0, 0);
            poly1.AddVertexAt(11, new Point2d(50 * scale1 * 0.00112742189006399, 50 * scale1 * 0.000106363636364009), 0, 0, 0);
            poly1.AddVertexAt(12, new Point2d(50 * scale1 * 0.000877421890063997, 50 * scale1 * 0.000226363636364013), 0, 0, 0);
            poly1.AddVertexAt(13, new Point2d(50 * scale1 * 0.000644088556730663, 50 * scale1 * 0.000456363636363974), 0, 0, 0);
            poly1.AddVertexAt(14, new Point2d(50 * scale1 * 0.00051075522339733, 50 * scale1 * 0.000696363636363984), 0, 0, 0);

            poly1.Closed = true;

            return poly1;
        }


        private Polyline get_poly_gc2(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(50 * scale1 * 0.00381075522339733, 50 * scale1 * 0.000626363636364005), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(50 * scale1 * 0.00361075522339733, 50 * scale1 * 0.000516363636363977), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(50 * scale1 * 0.003677421890064, 50 * scale1 * 0.000386363636363924), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(50 * scale1 * 0.003777421890064, 50 * scale1 * 0.000236363636363919), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(50 * scale1 * 0.00399408855673067, 50 * scale1 * 0.000216363636363965), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(50 * scale1 * 0.00424408855673067, 50 * scale1 * 0.000336363636363899), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(50 * scale1 * 0.00424408855673067, 50 * scale1 * 0.000496363636363952), 0, 0, 0);
            poly1.AddVertexAt(7, new Point2d(50 * scale1 * 0.004077421890064, 50 * scale1 * 0.000676363636363959), 0, 0, 0);
            poly1.AddVertexAt(8, new Point2d(50 * scale1 * 0.00381075522339733, 50 * scale1 * 0.000626363636364005), 0, 0, 0);

            poly1.Closed = true;
            return poly1;
        }

        private Polyline get_poly_gc3(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(50 * scale1 * 0.004777421890064, 50 * scale1 * 0.000996363636363995), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(50 * scale1 * 0.005027421890064, 50 * scale1 * 0.00125636363636403), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(50 * scale1 * 0.00551075522339733, 50 * scale1 * 0.00136636363636391), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(50 * scale1 * 0.00599408855673067, 50 * scale1 * 0.00125636363636403), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(50 * scale1 * 0.00626075522339733, 50 * scale1 * 0.000996363636363995), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(50 * scale1 * 0.006277421890064, 50 * scale1 * 0.000616363636363957), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(50 * scale1 * 0.00616075522339733, 50 * scale1 * 0.000526363636363953), 0, 0, 0);
            poly1.AddVertexAt(7, new Point2d(50 * scale1 * 0.00586075522339733, 50 * scale1 * 0.000426363636363973), 0, 0, 0);
            poly1.AddVertexAt(8, new Point2d(50 * scale1 * 0.005577421890064, 50 * scale1 * 0.00043636363636395), 0, 0, 0);
            poly1.AddVertexAt(9, new Point2d(50 * scale1 * 0.00536075522339733, 50 * scale1 * 0.000526363636363953), 0, 0, 0);
            poly1.AddVertexAt(10, new Point2d(50 * scale1 * 0.005127421890064, 50 * scale1 * 0.000616363636363957), 0, 0, 0);
            poly1.AddVertexAt(11, new Point2d(50 * scale1 * 0.00491075522339733, 50 * scale1 * 0.000706363636364031), 0, 0, 0);
            poly1.AddVertexAt(12, new Point2d(50 * scale1 * 0.004777421890064, 50 * scale1 * 0.00087636363636399), 0, 0, 0);


            poly1.Closed = true;

            return poly1;
        }


        private Polyline get_poly_gc4(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(50 * scale1 * 0.000172909471265833, 50 * scale1 * 0.00228727272727298), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(50 * scale1 * 5.85706354984968E-05, 50 * scale1 * 0.002039090909091), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(50 * scale1 * 7.52373021651653E-05, 50 * scale1 * 0.00177909090909097), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(50 * scale1 * 0.000141903968831832, 50 * scale1 * 0.00157909090909101), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(50 * scale1 * 0.000441903968831832, 50 * scale1 * 0.00151909090909101), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(50 * scale1 * 0.000741903968831833, 50 * scale1 * 0.00155909090909098), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(50 * scale1 * 0.000991903968831832, 50 * scale1 * 0.00170909090909099), 0, 0, 0);
            poly1.AddVertexAt(7, new Point2d(50 * scale1 * 0.000741903968831833, 50 * scale1 * 0.00205909090909095), 0, 0, 0);
            poly1.AddVertexAt(8, new Point2d(50 * scale1 * 0.000172909471265833, 50 * scale1 * 0.00228727272727298), 0, 0, 0);


            poly1.Closed = true;
            return poly1;
        }

        private Polyline get_poly_gc5(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(50 * scale1 * 0.0050078417347695, 50 * scale1 * 0.00194181818181896), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(50 * scale1 * 0.0045688310100445, 50 * scale1 * 0.00195090909090901), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(50 * scale1 * 0.00419037348872983, 50 * scale1 * 0.00178727272727301), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(50 * scale1 * 0.004114681984467, 50 * scale1 * 0.00156909090909096), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(50 * scale1 * 0.00449313950578167, 50 * scale1 * 0.00150545454545494), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(50 * scale1 * 0.00496242683221183, 50 * scale1 * 0.001578181818182), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(50 * scale1 * 0.00523491624755833, 50 * scale1 * 0.00176909090909099), 0, 0, 0);
            poly1.AddVertexAt(7, new Point2d(50 * scale1 * 0.0050078417347695, 50 * scale1 * 0.00194181818181896), 0, 0, 0);
            poly1.Closed = true;
            return poly1;
        }

        private Polyline get_poly_gc6(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(50 * scale1 * 0.00354408855673067, 50 * scale1 * 0.0028363636363639), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(50 * scale1 * 0.00409408855673067, 50 * scale1 * 0.00271636363636397), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(50 * scale1 * 0.004327421890064, 50 * scale1 * 0.00238636363636395), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(50 * scale1 * 0.00409408855673067, 50 * scale1 * 0.00206636363636392), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(50 * scale1 * 0.00354408855673067, 50 * scale1 * 0.00194636363636398), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(50 * scale1 * 0.002927421890064, 50 * scale1 * 0.00190636363636401), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(50 * scale1 * 0.00259408855673067, 50 * scale1 * 0.00197636363636398), 0, 0, 0);
            poly1.AddVertexAt(7, new Point2d(50 * scale1 * 0.002327421890064, 50 * scale1 * 0.00215636363636399), 0, 0, 0);
            poly1.AddVertexAt(8, new Point2d(50 * scale1 * 0.00157742189006399, 50 * scale1 * 0.00244636363636396), 0, 0, 0);
            poly1.AddVertexAt(9, new Point2d(50 * scale1 * 0.001177421890064, 50 * scale1 * 0.00272636363636401), 0, 0, 0);
            poly1.AddVertexAt(10, new Point2d(50 * scale1 * 0.001277421890064, 50 * scale1 * 0.00285636363636399), 0, 0, 0);
            poly1.AddVertexAt(11, new Point2d(50 * scale1 * 0.001877421890064, 50 * scale1 * 0.00311636363636396), 0, 0, 0);
            poly1.AddVertexAt(12, new Point2d(50 * scale1 * 0.00234408855673067, 50 * scale1 * 0.00316636363636398), 0, 0, 0);
            poly1.AddVertexAt(13, new Point2d(50 * scale1 * 0.00264408855673067, 50 * scale1 * 0.00307636363636398), 0, 0, 0);
            poly1.AddVertexAt(14, new Point2d(50 * scale1 * 0.00304408855673067, 50 * scale1 * 0.0029863636363639), 0, 0, 0);
            poly1.AddVertexAt(15, new Point2d(50 * scale1 * 0.00336075522339733, 50 * scale1 * 0.00293636363636395), 0, 0, 0);

            poly1.Closed = true;
            return poly1;
        }

        private Polyline get_poly_gc7(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(50 * scale1 * 0.0056549324747175, 50 * scale1 * 0.00399999999999999), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(50 * scale1 * 0.00542159914138417, 50 * scale1 * 0.00387), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(50 * scale1 * 0.00527159914138417, 50 * scale1 * 0.00350000000000001), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(50 * scale1 * 0.00527159914138417, 50 * scale1 * 0.00332000000000001), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(50 * scale1 * 0.00522159914138417, 50 * scale1 * 0.00292999999999999), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(50 * scale1 * 0.00547159914138417, 50 * scale1 * 0.00275999999999996), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(50 * scale1 * 0.00582159914138417, 50 * scale1 * 0.00269999999999996), 0, 0, 0);
            poly1.AddVertexAt(7, new Point2d(50 * scale1 * 0.00663826580805083, 50 * scale1 * 0.00272999999999996), 0, 0, 0);
            poly1.AddVertexAt(8, new Point2d(50 * scale1 * 0.0068549324747175, 50 * scale1 * 0.00283999999999999), 0, 0, 0);
            poly1.AddVertexAt(9, new Point2d(50 * scale1 * 0.00692159914138417, 50 * scale1 * 0.00311), 0, 0, 0);
            poly1.AddVertexAt(10, new Point2d(50 * scale1 * 0.0067049324747175, 50 * scale1 * 0.00326000000000001), 0, 0, 0);
            poly1.AddVertexAt(11, new Point2d(50 * scale1 * 0.0065049324747175, 50 * scale1 * 0.00342999999999996), 0, 0, 0);
            poly1.AddVertexAt(12, new Point2d(50 * scale1 * 0.00633826580805083, 50 * scale1 * 0.00367999999999995), 0, 0, 0);
            poly1.AddVertexAt(13, new Point2d(50 * scale1 * 0.0062049324747175, 50 * scale1 * 0.00385999999999996), 0, 0, 0);
            poly1.AddVertexAt(14, new Point2d(50 * scale1 * 0.0059049324747175, 50 * scale1 * 0.00393000000000001), 0, 0, 0);
            poly1.AddVertexAt(15, new Point2d(50 * scale1 * 0.0056549324747175, 50 * scale1 * 0.00399999999999999), 0, 0, 0);

            poly1.Closed = true;
            return poly1;
        }

        private Polyline get_poly_gc8(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(50 * scale1 * 0.004227421890064, 50 * scale1 * 0.00348636363636395), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(50 * scale1 * 0.00409408855673067, 50 * scale1 * 0.00337636363636399), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(50 * scale1 * 0.004077421890064, 50 * scale1 * 0.00323636363636396), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(50 * scale1 * 0.00426075522339733, 50 * scale1 * 0.003126363636364), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(50 * scale1 * 0.00449408855673067, 50 * scale1 * 0.00315636363636401), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(50 * scale1 * 0.00454408855673067, 50 * scale1 * 0.00330636363636401), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(50 * scale1 * 0.00451075522339733, 50 * scale1 * 0.00342636363636402), 0, 0, 0);
            poly1.AddVertexAt(7, new Point2d(50 * scale1 * 0.004227421890064, 50 * scale1 * 0.00348636363636395), 0, 0, 0);

            poly1.Closed = true;
            return poly1;
        }

        private Polyline get_poly_gc9(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(50 * scale1 * 0.00354408855673067, 50 * scale1 * 0.00344636363636397), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(50 * scale1 * 0.00341075522339733, 50 * scale1 * 0.00343636363636392), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(50 * scale1 * 0.00324408855673067, 50 * scale1 * 0.00348636363636395), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(50 * scale1 * 0.00311075522339733, 50 * scale1 * 0.00362636363636398), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(50 * scale1 * 0.003227421890064, 50 * scale1 * 0.00373636363636393), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(50 * scale1 * 0.00354408855673067, 50 * scale1 * 0.003676363636364), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(50 * scale1 * 0.003627421890064, 50 * scale1 * 0.00357636363636395), 0, 0, 0);

            poly1.Closed = true;
            return poly1;
        }

        private Polyline get_poly_gc10(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(50 * scale1 * 0.000564758891686667, 50 * scale1 * 0.00375727272727296), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(50 * scale1 * 0.00039809222502, 50 * scale1 * 0.00360727272727296), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(50 * scale1 * 0.00049809222502, 50 * scale1 * 0.00347727272727298), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(50 * scale1 * 0.000781425558353332, 50 * scale1 * 0.00342727272727295), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(50 * scale1 * 0.000998092225020002, 50 * scale1 * 0.00351727272727295), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(50 * scale1 * 0.000998092225020002, 50 * scale1 * 0.00370727272727294), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(50 * scale1 * 0.000564758891686667, 50 * scale1 * 0.00375727272727296), 0, 0, 0);

            poly1.Closed = true;
            return poly1;
        }


        private Polyline get_poly_gc1l(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(4 * scale1 + scale1 * 7.01350350027613, 0.051 * scale1 + scale1 * -2.118749611375), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(4 * scale1 + scale1 * 7.00778655848776, 0.051 * scale1 + scale1 * -2.1311587022841), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(4 * scale1 + scale1 * 7.0086198918211, 0.051 * scale1 + scale1 * -2.1441587022841), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(4 * scale1 + scale1 * 7.01195322515443, 0.051 * scale1 + scale1 * -2.1541587022841), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(4 * scale1 + scale1 * 7.02695322515443, 0.051 * scale1 + scale1 * -2.1571587022841), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(4 * scale1 + scale1 * 7.04195322515443, 0.051 * scale1 + scale1 * -2.1551587022841), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(4 * scale1 + scale1 * 7.05445322515443, 0.051 * scale1 + scale1 * -2.1476587022841), 0, 0, 0);
            poly1.AddVertexAt(7, new Point2d(4 * scale1 + scale1 * 7.04195322515443, 0.051 * scale1 + scale1 * -2.1301587022841), 0, 0, 0);
            poly1.AddVertexAt(8, new Point2d(4 * scale1 + scale1 * 7.01350350027613, 0.051 * scale1 + scale1 * -2.118749611375), 0, 0, 0);


            poly1.Closed = true;

            return poly1;
        }


        private Polyline get_poly_gc2l(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(4 * scale1 + scale1 * 7.108, 0.051 * scale1 + scale1 * -2.411), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(4 * scale1 + scale1 * 7.18, 0.051 * scale1 + scale1 * -2.411), 0, 0, 0);


            poly1.Closed = true;
            return poly1;
        }

        private Polyline get_poly_gc3l(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(4 * scale1 + scale1 * 7, 0.051 * scale1 + scale1 * -2.411), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(4 * scale1 + scale1 * 7.072, 0.051 * scale1 + scale1 * -2.411), 0, 0, 0);



            poly1.Closed = true;

            return poly1;
        }


        private Polyline get_poly_gc4l(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(4 * scale1 + scale1 * 7.02823794458433, 0.051 * scale1 + scale1 * -2.21213636363635), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(4 * scale1 + scale1 * 7.019904611251, 0.051 * scale1 + scale1 * -2.21963636363635), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(4 * scale1 + scale1 * 7.024904611251, 0.051 * scale1 + scale1 * -2.22613636363635), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(4 * scale1 + scale1 * 7.03907127791767, 0.051 * scale1 + scale1 * -2.22863636363635), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(4 * scale1 + scale1 * 7.049904611251, 0.051 * scale1 + scale1 * -2.22413636363635), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(4 * scale1 + scale1 * 7.049904611251, 0.051 * scale1 + scale1 * -2.21463636363635), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(4 * scale1 + scale1 * 7.02823794458433, 0.051 * scale1 + scale1 * -2.21213636363635), 0, 0, 0);

            poly1.Closed = true;
            return poly1;
        }

        private Polyline get_poly_gc5l(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(4 * scale1 + scale1 * 7.17720442783653, 0.051 * scale1 + scale1 * -2.2276818181818), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(4 * scale1 + scale1 * 7.17053776116987, 0.051 * scale1 + scale1 * -2.2281818181818), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(4 * scale1 + scale1 * 7.16220442783653, 0.051 * scale1 + scale1 * -2.2256818181818), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(4 * scale1 + scale1 * 7.15553776116987, 0.051 * scale1 + scale1 * -2.2186818181818), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(4 * scale1 + scale1 * 7.1613710945032, 0.051 * scale1 + scale1 * -2.2131818181818), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(4 * scale1 + scale1 * 7.17720442783653, 0.051 * scale1 + scale1 * -2.2161818181818), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(4 * scale1 + scale1 * 7.1813710945032, 0.051 * scale1 + scale1 * -2.2211818181818), 0, 0, 0);

            poly1.Closed = true;
            return poly1;
        }

        private Polyline get_poly_gc6l(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(4 * scale1 + scale1 * 7.15453342189391, 0.051 * scale1 + scale1 * -2.01230210919678), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(4 * scale1 + scale1 * 7.14786675522725, 0.051 * scale1 + scale1 * -2.01780210919678), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(4 * scale1 + scale1 * 7.14703342189391, 0.051 * scale1 + scale1 * -2.02480210919678), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(4 * scale1 + scale1 * 7.15620008856058, 0.051 * scale1 + scale1 * -2.03030210919678), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(4 * scale1 + scale1 * 7.16786675522725, 0.051 * scale1 + scale1 * -2.02880210919678), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(4 * scale1 + scale1 * 7.17036675522725, 0.051 * scale1 + scale1 * -2.02130210919678), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(4 * scale1 + scale1 * 7.16870008856058, 0.051 * scale1 + scale1 * -2.01530210919678), 0, 0, 0);
            poly1.AddVertexAt(7, new Point2d(4 * scale1 + scale1 * 7.15453342189391, 0.051 * scale1 + scale1 * -2.01230210919678), 0, 0, 0);


            poly1.Closed = true;
            return poly1;
        }

        private Polyline get_poly_gc7l(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(4 * scale1 + scale1 * 7.10967107669487, 0.051 * scale1 + scale1 * -2.12733965901626), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(4 * scale1 + scale1 * 7.0980044100282, 0.051 * scale1 + scale1 * -2.13383965901626), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(4 * scale1 + scale1 * 7.0905044100282, 0.051 * scale1 + scale1 * -2.15233965901626), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(4 * scale1 + scale1 * 7.0905044100282, 0.051 * scale1 + scale1 * -2.16133965901626), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(4 * scale1 + scale1 * 7.0880044100282, 0.051 * scale1 + scale1 * -2.18083965901626), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(4 * scale1 + scale1 * 7.1005044100282, 0.051 * scale1 + scale1 * -2.18933965901626), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(4 * scale1 + scale1 * 7.1180044100282, 0.051 * scale1 + scale1 * -2.19233965901626), 0, 0, 0);
            poly1.AddVertexAt(7, new Point2d(4 * scale1 + scale1 * 7.15883774336154, 0.051 * scale1 + scale1 * -2.19083965901626), 0, 0, 0);
            poly1.AddVertexAt(8, new Point2d(4 * scale1 + scale1 * 7.16967107669487, 0.051 * scale1 + scale1 * -2.18533965901626), 0, 0, 0);
            poly1.AddVertexAt(9, new Point2d(4 * scale1 + scale1 * 7.1730044100282, 0.051 * scale1 + scale1 * -2.17183965901626), 0, 0, 0);
            poly1.AddVertexAt(10, new Point2d(4 * scale1 + scale1 * 7.16217107669487, 0.051 * scale1 + scale1 * -2.16433965901626), 0, 0, 0);
            poly1.AddVertexAt(11, new Point2d(4 * scale1 + scale1 * 7.15217107669487, 0.051 * scale1 + scale1 * -2.15583965901626), 0, 0, 0);
            poly1.AddVertexAt(12, new Point2d(4 * scale1 + scale1 * 7.14383774336154, 0.051 * scale1 + scale1 * -2.14333965901626), 0, 0, 0);
            poly1.AddVertexAt(13, new Point2d(4 * scale1 + scale1 * 7.13717107669487, 0.051 * scale1 + scale1 * -2.13433965901626), 0, 0, 0);
            poly1.AddVertexAt(14, new Point2d(4 * scale1 + scale1 * 7.12217107669487, 0.051 * scale1 + scale1 * -2.13083965901626), 0, 0, 0);
            poly1.AddVertexAt(15, new Point2d(4 * scale1 + scale1 * 7.10967107669487, 0.051 * scale1 + scale1 * -2.12733965901626), 0, 0, 0);

            poly1.Closed = true;
            return poly1;
        }

        private Polyline get_poly_gc8l(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(4 * scale1 + scale1 * 7.18, 0.051 * scale1 + scale1 * -2.26065035536885), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(4 * scale1 + scale1 * 7.18, 0.051 * scale1 + scale1 * -2.30237344279885), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(4 * scale1 + scale1 * 7.1463710945032, 0.051 * scale1 + scale1 * -2.3046818181818), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(4 * scale1 + scale1 * 7.12970442783653, 0.051 * scale1 + scale1 * -2.3011818181818), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(4 * scale1 + scale1 * 7.1163710945032, 0.051 * scale1 + scale1 * -2.2921818181818), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(4 * scale1 + scale1 * 7.0788710945032, 0.051 * scale1 + scale1 * -2.2776818181818), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(4 * scale1 + scale1 * 7.0588710945032, 0.051 * scale1 + scale1 * -2.2636818181818), 0, 0, 0);
            poly1.AddVertexAt(7, new Point2d(4 * scale1 + scale1 * 7.0638710945032, 0.051 * scale1 + scale1 * -2.2571818181818), 0, 0, 0);
            poly1.AddVertexAt(8, new Point2d(4 * scale1 + scale1 * 7.0938710945032, 0.051 * scale1 + scale1 * -2.2441818181818), 0, 0, 0);
            poly1.AddVertexAt(9, new Point2d(4 * scale1 + scale1 * 7.11720442783653, 0.051 * scale1 + scale1 * -2.2416818181818), 0, 0, 0);
            poly1.AddVertexAt(10, new Point2d(4 * scale1 + scale1 * 7.13220442783653, 0.051 * scale1 + scale1 * -2.2461818181818), 0, 0, 0);
            poly1.AddVertexAt(11, new Point2d(4 * scale1 + scale1 * 7.15220442783653, 0.051 * scale1 + scale1 * -2.25068181818181), 0, 0, 0);
            poly1.AddVertexAt(12, new Point2d(4 * scale1 + scale1 * 7.16803776116987, 0.051 * scale1 + scale1 * -2.2531818181818), 0, 0, 0);


            poly1.Closed = true;
            return poly1;
        }

        private Polyline get_poly_gc9l(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(4 * scale1 + scale1 * 7.08615443613922, 0.051 * scale1 + scale1 * -2.00751053067765), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(4 * scale1 + scale1 * 7.06420389990297, 0.051 * scale1 + scale1 * -2.00705598522315), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(4 * scale1 + scale1 * 7.03653280297294, 0.051 * scale1 + scale1 * -2.0130776129663), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(4 * scale1 + scale1 * 7.03471794025972, 0.051 * scale1 + scale1 * -2.02637950889766), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(4 * scale1 + scale1 * 7.04149644862409, 0.051 * scale1 + scale1 * -2.03222413732922), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(4 * scale1 + scale1 * 7.06041932468982, 0.051 * scale1 + scale1 * -2.02932871249585), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(4 * scale1 + scale1 * 7.08388369101133, 0.051 * scale1 + scale1 * -2.0256923488595), 0, 0, 0);
            poly1.AddVertexAt(7, new Point2d(4 * scale1 + scale1 * 7.09750816177866, 0.051 * scale1 + scale1 * -2.01614689431405), 0, 0, 0);
            poly1.AddVertexAt(8, new Point2d(4 * scale1 + scale1 * 7.08615443613922, 0.051 * scale1 + scale1 * -2.00751053067765), 0, 0, 0);


            poly1.Closed = true;
            return poly1;
        }

        private Polyline get_poly_gc10l(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(4 * scale1 + scale1 * 7.00864547356329, 0.051 * scale1 + scale1 * -2.28563636363635), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(4 * scale1 + scale1 * 7.00292853177493, 0.051 * scale1 + scale1 * -2.29804545454545), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(4 * scale1 + scale1 * 7.00376186510826, 0.051 * scale1 + scale1 * -2.31104545454545), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(4 * scale1 + scale1 * 7.00709519844159, 0.051 * scale1 + scale1 * -2.32104545454545), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(4 * scale1 + scale1 * 7.02209519844159, 0.051 * scale1 + scale1 * -2.32404545454545), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(4 * scale1 + scale1 * 7.03709519844159, 0.051 * scale1 + scale1 * -2.32204545454545), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(4 * scale1 + scale1 * 7.04959519844159, 0.051 * scale1 + scale1 * -2.31454545454545), 0, 0, 0);
            poly1.AddVertexAt(7, new Point2d(4 * scale1 + scale1 * 7.03709519844159, 0.051 * scale1 + scale1 * -2.29704545454545), 0, 0, 0);
            poly1.AddVertexAt(8, new Point2d(4 * scale1 + scale1 * 7.00864547356329, 0.051 * scale1 + scale1 * -2.28563636363635), 0, 0, 0);


            poly1.Closed = true;
            return poly1;
        }


        private Polyline get_poly_gc11l(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(4 * scale1 + scale1 * 7.08175410082965, 0.051 * scale1 + scale1 * -2.06997894553156), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(4 * scale1 + scale1 * 7.09425410082965, 0.051 * scale1 + scale1 * -2.05697894553156), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(4 * scale1 + scale1 * 7.11842076749631, 0.051 * scale1 + scale1 * -2.05147894553156), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(4 * scale1 + scale1 * 7.14258743416298, 0.051 * scale1 + scale1 * -2.05697894553156), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(4 * scale1 + scale1 * 7.15592076749631, 0.051 * scale1 + scale1 * -2.06997894553156), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(4 * scale1 + scale1 * 7.15675410082965, 0.051 * scale1 + scale1 * -2.08897894553156), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(4 * scale1 + scale1 * 7.15092076749631, 0.051 * scale1 + scale1 * -2.09347894553156), 0, 0, 0);
            poly1.AddVertexAt(7, new Point2d(4 * scale1 + scale1 * 7.13592076749631, 0.051 * scale1 + scale1 * -2.09847894553156), 0, 0, 0);
            poly1.AddVertexAt(8, new Point2d(4 * scale1 + scale1 * 7.12175410082965, 0.051 * scale1 + scale1 * -2.09797894553156), 0, 0, 0);
            poly1.AddVertexAt(9, new Point2d(4 * scale1 + scale1 * 7.11092076749631, 0.051 * scale1 + scale1 * -2.09347894553156), 0, 0, 0);
            poly1.AddVertexAt(10, new Point2d(4 * scale1 + scale1 * 7.09925410082965, 0.051 * scale1 + scale1 * -2.08897894553156), 0, 0, 0);
            poly1.AddVertexAt(11, new Point2d(4 * scale1 + scale1 * 7.08842076749631, 0.051 * scale1 + scale1 * -2.08447894553156), 0, 0, 0);
            poly1.AddVertexAt(12, new Point2d(4 * scale1 + scale1 * 7.08175410082965, 0.051 * scale1 + scale1 * -2.07597894553156), 0, 0, 0);


            poly1.Closed = true;
            return poly1;
        }

        private Polyline get_poly_gc12l(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(4 * scale1 + scale1 * 7.02630011057061, 0.051 * scale1 + scale1 * -2.0732832579504), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(4 * scale1 + scale1 * 7.01630011057061, 0.051 * scale1 + scale1 * -2.0787832579504), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(4 * scale1 + scale1 * 7.01963344390394, 0.051 * scale1 + scale1 * -2.0852832579504), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(4 * scale1 + scale1 * 7.02463344390394, 0.051 * scale1 + scale1 * -2.0927832579504), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(4 * scale1 + scale1 * 7.03546677723727, 0.051 * scale1 + scale1 * -2.0937832579504), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(4 * scale1 + scale1 * 7.04796677723727, 0.051 * scale1 + scale1 * -2.0877832579504), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(4 * scale1 + scale1 * 7.04796677723727, 0.051 * scale1 + scale1 * -2.0797832579504), 0, 0, 0);
            poly1.AddVertexAt(7, new Point2d(4 * scale1 + scale1 * 7.03963344390394, 0.051 * scale1 + scale1 * -2.0707832579504), 0, 0, 0);
            poly1.AddVertexAt(8, new Point2d(4 * scale1 + scale1 * 7.02630011057061, 0.051 * scale1 + scale1 * -2.0732832579504), 0, 0, 0);


            poly1.Closed = true;
            return poly1;
        }

        private Polyline get_poly_gc13l(double scale1)
        {
            Polyline poly1 = new Polyline();
            poly1.AddVertexAt(0, new Point2d(4 * scale1 + scale1 * 7.02803776116987, 0.051 * scale1 + scale1 * -2.3501818181818), 0, 0, 0);
            poly1.AddVertexAt(1, new Point2d(4 * scale1 + scale1 * 7.0513710945032, 0.051 * scale1 + scale1 * -2.3436818181818), 0, 0, 0);
            poly1.AddVertexAt(2, new Point2d(4 * scale1 + scale1 * 7.09720442783653, 0.051 * scale1 + scale1 * -2.3461818181818), 0, 0, 0);
            poly1.AddVertexAt(3, new Point2d(4 * scale1 + scale1 * 7.12053776116987, 0.051 * scale1 + scale1 * -2.3501818181818), 0, 0, 0);
            poly1.AddVertexAt(4, new Point2d(4 * scale1 + scale1 * 7.12470442783653, 0.051 * scale1 + scale1 * -2.3566818181818), 0, 0, 0);
            poly1.AddVertexAt(5, new Point2d(4 * scale1 + scale1 * 7.12553776116987, 0.051 * scale1 + scale1 * -2.3631818181818), 0, 0, 0);
            poly1.AddVertexAt(6, new Point2d(4 * scale1 + scale1 * 7.1263710945032, 0.051 * scale1 + scale1 * -2.3711818181818), 0, 0, 0);
            poly1.AddVertexAt(7, new Point2d(4 * scale1 + scale1 * 7.13303776116987, 0.051 * scale1 + scale1 * -2.3781818181818), 0, 0, 0);
            poly1.AddVertexAt(8, new Point2d(4 * scale1 + scale1 * 7.12803776116987, 0.051 * scale1 + scale1 * -2.3901818181818), 0, 0, 0);
            poly1.AddVertexAt(9, new Point2d(4 * scale1 + scale1 * 7.10970442783653, 0.051 * scale1 + scale1 * -2.3966818181818), 0, 0, 0);
            poly1.AddVertexAt(10, new Point2d(4 * scale1 + scale1 * 7.0738710945032, 0.051 * scale1 + scale1 * -2.3976818181818), 0, 0, 0);
            poly1.AddVertexAt(11, new Point2d(4 * scale1 + scale1 * 7.0563710945032, 0.051 * scale1 + scale1 * -2.3946818181818), 0, 0, 0);
            poly1.AddVertexAt(12, new Point2d(4 * scale1 + scale1 * 7.0438710945032, 0.051 * scale1 + scale1 * -2.3886818181818), 0, 0, 0);
            poly1.AddVertexAt(13, new Point2d(4 * scale1 + scale1 * 7.03220442783653, 0.051 * scale1 + scale1 * -2.3771818181818), 0, 0, 0);
            poly1.AddVertexAt(14, new Point2d(4 * scale1 + scale1 * 7.02553776116987, 0.051 * scale1 + scale1 * -2.3651818181818), 0, 0, 0);


            poly1.Closed = true;
            return poly1;
        }

        private void textBox_pozitive_KeyPress(object sender, KeyPressEventArgs e)
        {
            Functions.textbox_input_only_pozitive_doubles_at_keypress(sender, e);
        }
        private static Hatch CreateHatch(Curve poly1, string hatchname, double scale1, double angle1)
        {


            ObjectIdCollection col1 = new ObjectIdCollection();
            col1.Add(poly1.ObjectId);
            Hatch hatch1 = new Hatch();
            hatch1.SetHatchPattern(HatchPatternType.PreDefined, hatchname);
            hatch1.Normal = Vector3d.ZAxis;
            hatch1.PatternScale = scale1;
            hatch1.PatternAngle = angle1;
            hatch1.AppendLoop(HatchLoopTypes.Default, col1);
            hatch1.EvaluateHatch(true);

            return hatch1;
        }

        private static bool DoesAssemblyMatchLoad(System.Reflection.Assembly reference1)
        {
            try
            {
                var loadedAssembly = System.Reflection.Assembly.Load(reference1.FullName);
                return reference1.CodeBase == loadedAssembly.CodeBase;
            }
            catch (FileNotFoundException)
            {
                return false;
            }

        }

    }
}
