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
                                        boe.Color = color_labels;
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
                                                        add_pattern_GM(bltrec1, 1, 1, 1, poly1, BTrecord, Trans1);
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
                                                        add_pattern_GP(bltrec1, 1, 1, 1, poly1, BTrecord, Trans1);
                                                    }
                                                    if (lista_legend[i] == "GC")
                                                    {
                                                        add_pattern_GC(bltrec1, 1, 1, 1, poly1, BTrecord, Trans1);
                                                    }
                                                    if (lista_legend[i] == "MUDSTONE")
                                                    {
                                                        add_pattern_MUDSTONE_legend(bltrec1, 1, 1, 1, poly1, BTrecord, Trans1);
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
