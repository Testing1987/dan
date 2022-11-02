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
    public partial class Wgen_feature : Form
    {
        bool clickdragdown;
        Point lastLocation;
        public static System.Data.DataTable dt_display = null;

        string col_pmc_as_cc = "USE PMC AS CC";


        public Wgen_feature()
        {
            InitializeComponent();
            dt_display = Wgen_main_form.dt_feature_codes.Copy();

            if (Wgen_main_form.client_name == "xxx") Wgen_main_form.client_name = Wgen_main_form.lista_clienti[0];

            for (int i = dt_display.Rows.Count - 1; i >= 0; --i)
            {
                if (Convert.ToString(dt_display.Rows[i][0]) != Wgen_main_form.client_name)
                {
                    dt_display.Rows[i].Delete();
                }
            }

            dt_display.Columns.RemoveAt(0);

            dataGridView_feature_codes.DataSource = dt_display;
            dataGridView_feature_codes.Columns[0].Width = 300;
            dataGridView_feature_codes.Columns[1].Width = 100;
            dataGridView_feature_codes.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_feature_codes.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView_feature_codes.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_feature_codes.DefaultCellStyle.ForeColor = Color.White;
            dataGridView_feature_codes.EnableHeadersVisualStyles = false;

            label_client.Text = Wgen_main_form.client_name;
        }



        private void clickmove_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            clickdragdown = true;
            lastLocation = e.Location;
        }
        private void clickmove_MouseMove(object sender, MouseEventArgs e)
        {
            if (clickdragdown == true)
            {
                this.Location = new Point(
                  (this.Location.X - lastLocation.X) + e.X, (this.Location.Y - lastLocation.Y) + e.Y);

                this.Update();
            }
        }
        private void clickmove_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            clickdragdown = false;
        }
        private void button_Exit_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < Wgen_main_form.dt_feature_codes.Rows.Count; ++i)
            {
                if (Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[i][0]) == Wgen_main_form.client_name)
                {
                    string string1 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[i][1]);

                    for (int j = 0; j < dt_display.Rows.Count; ++j)
                    {

                        string string2 = Convert.ToString(dt_display.Rows[j][0]);

                        if (string1 == string2)
                        {
                            if (dt_display.Rows[j][1] != DBNull.Value)
                            {
                                string string3 = Convert.ToString(dt_display.Rows[j][1]);
                            }

                            Wgen_main_form.dt_feature_codes.Rows[i][2] = dt_display.Rows[j][1];
                            j = dt_display.Rows.Count;
                        }
                    }
                }
            }

            Wgen_main_form.lista_feature_code.Clear();
            for (int i = 0; i < Wgen_main_form.dt_feature_codes.Rows.Count; ++i)
            {
                string client = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[i][0]);
                if (client == Wgen_main_form.client_name)
                {
                    if ((bool)Wgen_main_form.dt_feature_codes.Rows[i][2] == true)
                    {
                        Wgen_main_form.lista_feature_code.Add(Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[i][1]));
                    }
                }
            }

            if (Wgen_main_form.lista_feature_code.Contains("WELD") == false) Wgen_main_form.lista_feature_code.Add("WELD");
            if (Wgen_main_form.lista_feature_code.Contains("BEND") == false) Wgen_main_form.lista_feature_code.Add("BEND");
            if (Wgen_main_form.lista_feature_code.Contains("NATURAL_GROUND") == false) Wgen_main_form.lista_feature_code.Add("NATURAL_GROUND");

            string file1 = Wgen_main_form.WGEN_folder + "wgen_feature_codes.xlsx";
            bool visible1 = false;
            if (Functions.is_dan_popescu() == true) visible1 = true;

            int nr_excel_open = Functions.Get_no_of_workbooks_from_Excel();

            if (System.IO.File.Exists(file1) == true)
            {
                Microsoft.Office.Interop.Excel.Application Excel1 = null;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;
                try
                {
                    try
                    {
                        Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    }
                    catch (System.Exception ex)
                    {
                        Excel1 = new Microsoft.Office.Interop.Excel.Application();
                    }
                    Excel1.Visible = visible1;
                    Workbook1 = Excel1.Workbooks.Open(file1);
                    foreach (Microsoft.Office.Interop.Excel.Worksheet W1 in Workbook1.Worksheets)
                    {

                        System.Data.DataTable dt1 = new System.Data.DataTable();

                        dt1.Columns.Add("FEATURE CODE", typeof(string));
                        dt1.Columns.Add("INCLUDED (YES/NO)", typeof(string));
                        dt1.Columns.Add("DESCRIPTION", typeof(string));
                        dt1.Columns.Add("CHECK 1", typeof(string));
                        dt1.Columns.Add("CHECK 2", typeof(string));
                        dt1.Columns.Add("CHECK 3", typeof(string));
                        dt1.Columns.Add("CHECK 4", typeof(string));
                        dt1.Columns.Add("CHECK 5", typeof(string));
                        dt1.Columns.Add("CHECK 6", typeof(string));
                        dt1.Columns.Add("CHECK 7", typeof(string));
                        dt1.Columns.Add("CHECK 8", typeof(string));
                        dt1.Columns.Add("CHECK 9", typeof(string));
                        dt1.Columns.Add("CHECK 10", typeof(string));
                        dt1.Columns.Add("CHECK XRAY", typeof(string));
                        dt1.Columns.Add("BEND TYPE", typeof(string));
                        dt1.Columns.Add("BEND DEFLECTION TYPE", typeof(string));
                        dt1.Columns.Add("BEND POSITION", typeof(string));
                        dt1.Columns.Add("BEND HORIZONTAL DEFLECTION", typeof(string));
                        dt1.Columns.Add("BEND VERTICAL DEFLECTION", typeof(string));
                        dt1.Columns.Add("WELD MM BACK", typeof(string));
                        dt1.Columns.Add("WELD MM AHEAD", typeof(string));
                        dt1.Columns.Add("CHECK DJ AGAINST XRAY", typeof(string));
                        dt1.Columns.Add("DJ TOTAL LENGTH", typeof(string));
                        dt1.Columns.Add(col_pmc_as_cc, typeof(string));


                        string nume1 = W1.Name;
                        for (int j = 0; j < Wgen_main_form.dt_feature_codes.Rows.Count; ++j)
                        {
                            if (Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][0]) == nume1)
                            {
                                dt1.Rows.Add();
                                string fc = "";

                                if (Wgen_main_form.dt_feature_codes.Rows[j][1] != DBNull.Value)
                                {
                                    string val1 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][1]);
                                    dt1.Rows[dt1.Rows.Count - 1][0] = val1;
                                    fc = val1;
                                }

                                if (Wgen_main_form.dt_feature_codes.Rows[j][1] != DBNull.Value)
                                {
                                    string val2 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][2]);
                                    if (val2 == "True")
                                    {
                                        dt1.Rows[dt1.Rows.Count - 1][1] = "YES";
                                    }
                                    else
                                    {
                                        dt1.Rows[dt1.Rows.Count - 1][1] = "NO";
                                    }
                                }
                                if (Wgen_main_form.dt_feature_codes.Rows[j][1] != DBNull.Value)
                                {
                                    string val3 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][3]);
                                    dt1.Rows[dt1.Rows.Count - 1][2] = val3;

                                    string val4 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][4]);
                                    dt1.Rows[dt1.Rows.Count - 1][3] = val4;

                                    string val5 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][5]);
                                    dt1.Rows[dt1.Rows.Count - 1][4] = val5;

                                    string val6 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][6]);
                                    dt1.Rows[dt1.Rows.Count - 1][5] = val6;

                                    string val7 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][7]);
                                    dt1.Rows[dt1.Rows.Count - 1][6] = val7;

                                    string val8 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][8]);
                                    dt1.Rows[dt1.Rows.Count - 1][7] = val8;

                                    string val9 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][9]);
                                    dt1.Rows[dt1.Rows.Count - 1][8] = val9;

                                    string val10 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][10]);
                                    dt1.Rows[dt1.Rows.Count - 1][9] = val10;

                                    string val11 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][11]);
                                    dt1.Rows[dt1.Rows.Count - 1][10] = val11;

                                    string val12 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][12]);
                                    dt1.Rows[dt1.Rows.Count - 1][11] = val12;

                                    string val13 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][13]);
                                    dt1.Rows[dt1.Rows.Count - 1][12] = val13;

                                    string val14 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][14]);
                                    dt1.Rows[dt1.Rows.Count - 1][13] = val14;

                                    string val15 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][15]);
                                    dt1.Rows[dt1.Rows.Count - 1][14] = val15;

                                    string val16 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][16]);
                                    dt1.Rows[dt1.Rows.Count - 1][15] = val16;

                                    string val17 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][17]);
                                    dt1.Rows[dt1.Rows.Count - 1][16] = val17;

                                    string val18 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][18]);
                                    dt1.Rows[dt1.Rows.Count - 1][17] = val18;

                                    string val19 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][19]);
                                    dt1.Rows[dt1.Rows.Count - 1][18] = val19;

                                    string val20 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][20]);
                                    dt1.Rows[dt1.Rows.Count - 1][19] = val20;

                                    string val21 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][21]);
                                    dt1.Rows[dt1.Rows.Count - 1][20] = val21;

                                    string val22 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][22]);
                                    dt1.Rows[dt1.Rows.Count - 1][21] = val22;

                                    string val23 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j]["DJ TOTAL LENGTH"]);
                                    if (val23 == "True")
                                    {
                                        val23 = "YES";
                                    }
                                    else
                                    {
                                        val23 = "NO";
                                    }

                                    dt1.Rows[dt1.Rows.Count - 1][22] = val23;

                                    if (fc == "PIPE_MATERIAL_CHANGE")
                                    {

                                        string val24 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][col_pmc_as_cc]);
                                        if (val24 == "True")
                                        {
                                            val24 = "YES";
                                        }
                                        else
                                        {
                                            val24 = "NO";
                                        }

                                        dt1.Rows[dt1.Rows.Count - 1][23] = val24;
                                    }


                                }
                            }
                        }



                        if (dt1.Rows.Count > 0)
                        {
                            if (W1.Name.ToLower() == Wgen_main_form.client_name.ToLower())
                            {
                                if (Wgen_main_form.check_dj_against_xray == true)
                                {
                                    dt1.Rows[0][21] = "YES";
                                }
                                else
                                {
                                    dt1.Rows[0][21] = "NO";
                                }
                            }


                            int NrR = dt1.Rows.Count;
                            int NrC = dt1.Columns.Count;

                            object[,] values = new object[NrR, NrC];
                            for (int i = 0; i < NrR; ++i)
                            {
                                for (int j = 0; j < NrC; ++j)
                                {
                                    if (dt1.Rows[i][j] != DBNull.Value)
                                    {
                                        values[i, j] = dt1.Rows[i][j];
                                    }
                                }
                            }
                            Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[2, 1], W1.Cells[NrR + 1, NrC]];
                            range1.Value2 = values;
                            W1.Range["W1"].Value2 = "PIPE MANIFEST IS DISPLAYING FOR DOUBLE JOINTS TOTAL LENGTH";
                            W1.Range["V1"].Value2 = "INCOMPLETE PIPE MANIFEST";
                            W1.Range["X1"].Value2 = "USE PMC AS CC";
                            W1.Range["A1:X1"].Font.Bold = true;
                        }
                    }

                    foreach (Microsoft.Office.Interop.Excel.Worksheet W1 in Workbook1.Worksheets)
                    {
                        if (W1.Name.ToLower() == Wgen_main_form.client_name.ToLower())
                        {
                            W1.Move(Before: Workbook1.Worksheets[1]);
                            //W1.Move(After: Workbook1.Worksheets[1]);
                        }
                    }
                    Wgen_main_form.tpage_weldmap.set_button_gen_wmR2();
                    Workbook1.Save();
                    Workbook1.Close();
                    if (nr_excel_open == 0) Excel1.Quit();
                }
                catch (System.Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);
                }
                finally
                {
                    if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                    if (nr_excel_open == 0 && Excel1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }
            }

            maximize_Wgen();
            this.Close();
        }
        private void button_minimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void maximize_Wgen()
        {
            foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
            {
                if (Forma1 is Alignment_mdi.Wgen_main_form)
                {
                    Forma1.Focus();
                    Forma1.WindowState = System.Windows.Forms.FormWindowState.Normal;
                }
            }
        }

        private void button_plus1_Click(object sender, EventArgs e)
        {

            if (Wgen_main_form.lista_clienti.Count > 1)
            {
                dt_display = Wgen_main_form.dt_feature_codes.Copy();

                int curent_index = Wgen_main_form.lista_clienti.IndexOf(Wgen_main_form.client_name);

                if (curent_index < Wgen_main_form.lista_clienti.Count - 1)
                {
                    Wgen_main_form.client_name = Wgen_main_form.lista_clienti[curent_index + 1];
                }
                else
                {
                    Wgen_main_form.client_name = Wgen_main_form.lista_clienti[0];
                }

                for (int i = dt_display.Rows.Count - 1; i >= 0; --i)
                {
                    if (Convert.ToString(dt_display.Rows[i][0]) != Wgen_main_form.client_name)
                    {
                        dt_display.Rows[i].Delete();
                    }
                }

                dt_display.Columns.RemoveAt(0);



                label_client.Text = Wgen_main_form.client_name;

                for (int i = 0; i < dt_display.Rows.Count; ++i)
                {
                    if ((bool)dt_display.Rows[i][1] == true)
                    {
                        Wgen_main_form.lista_feature_code.Add(Convert.ToString(dt_display.Rows[i][0]));
                    }
                }

                if (dt_display.Rows[0]["CHECK DJ AGAINST XRAY"] != DBNull.Value && Convert.ToString(dt_display.Rows[0]["CHECK DJ AGAINST XRAY"]) == "YES")
                {
                    Wgen_main_form.check_dj_against_xray = true;
                }
                else
                {
                    Wgen_main_form.check_dj_against_xray = false;
                }


                dataGridView_feature_codes.DataSource = dt_display;
                dataGridView_feature_codes.Columns[0].Width = 300;
                dataGridView_feature_codes.Columns[1].Width = 100;
                dataGridView_feature_codes.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_feature_codes.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                dataGridView_feature_codes.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_feature_codes.DefaultCellStyle.ForeColor = Color.White;
                dataGridView_feature_codes.EnableHeadersVisualStyles = false;

                Wgen_main_form.tpage_allpts.set_label_client(Wgen_main_form.client_name);
                Wgen_main_form.tpage_pipe_manifest.set_label_client(Wgen_main_form.client_name);

                if (dt_display.Rows[0]["DJ TOTAL LENGTH"] != DBNull.Value && Convert.ToBoolean(dt_display.Rows[0]["DJ TOTAL LENGTH"]) == true)
                {

                    checkBox_use_total_length_for_dj.Checked = true;
                }
                else
                {

                    checkBox_use_total_length_for_dj.Checked = false;
                }


                Wgen_main_form.use_pmc_as_cc = false;
                set_checkbox_pmc_as_cc(false);
                for (int i = 0; i < dt_display.Rows.Count; ++i)
                {
                    if (dt_display.Rows[i]["FEATURE CODE"] != DBNull.Value && dt_display.Rows[i][col_pmc_as_cc] != DBNull.Value)
                    {

                        if (Convert.ToString(dt_display.Rows[i]["FEATURE CODE"]) == "PIPE_MATERIAL_CHANGE")
                        {
                            if (dt_display.Rows[i][col_pmc_as_cc] != DBNull.Value && Convert.ToBoolean(dt_display.Rows[i][col_pmc_as_cc]) == true)
                            {
                                Wgen_main_form.use_pmc_as_cc = true;
                                set_checkbox_pmc_as_cc(true);

                            }
                        }
                    }
                }
            }

        }




        private void button_minus1_Click(object sender, EventArgs e)
        {
            if (Wgen_main_form.lista_clienti.Count > 1)
            {
                dt_display = Wgen_main_form.dt_feature_codes.Copy();

                int curent_index = Wgen_main_form.lista_clienti.IndexOf(Wgen_main_form.client_name);

                if (curent_index > 0)
                {
                    Wgen_main_form.client_name = Wgen_main_form.lista_clienti[curent_index - 1];
                }
                else
                {
                    Wgen_main_form.client_name = Wgen_main_form.lista_clienti[Wgen_main_form.lista_clienti.Count - 1];
                }

                for (int i = dt_display.Rows.Count - 1; i >= 0; --i)
                {
                    if (Convert.ToString(dt_display.Rows[i][0]) != Wgen_main_form.client_name)
                    {
                        dt_display.Rows[i].Delete();
                    }

                }

                dt_display.Columns.RemoveAt(0);


                label_client.Text = Wgen_main_form.client_name;

                for (int i = 0; i < dt_display.Rows.Count; ++i)
                {
                    if ((bool)dt_display.Rows[i][1] == true)
                    {
                        Wgen_main_form.lista_feature_code.Add(Convert.ToString(dt_display.Rows[i][0]));
                    }
                }


                if (dt_display.Rows[0]["CHECK DJ AGAINST XRAY"] != DBNull.Value && Convert.ToString(dt_display.Rows[0]["CHECK DJ AGAINST XRAY"]) == "YES")
                {
                    Wgen_main_form.check_dj_against_xray = true;
                }
                else
                {
                    Wgen_main_form.check_dj_against_xray = false;
                }


                dataGridView_feature_codes.DataSource = dt_display;
                dataGridView_feature_codes.Columns[0].Width = 300;
                dataGridView_feature_codes.Columns[1].Width = 100;
                dataGridView_feature_codes.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_feature_codes.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                dataGridView_feature_codes.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_feature_codes.DefaultCellStyle.ForeColor = Color.White;
                dataGridView_feature_codes.EnableHeadersVisualStyles = false;

                Wgen_main_form.tpage_allpts.set_label_client(Wgen_main_form.client_name);
                Wgen_main_form.tpage_pipe_manifest.set_label_client(Wgen_main_form.client_name);

                if (dt_display.Rows[0]["DJ TOTAL LENGTH"] != DBNull.Value && Convert.ToBoolean(dt_display.Rows[0]["DJ TOTAL LENGTH"]) == true)
                {

                    checkBox_use_total_length_for_dj.Checked = true;
                }
                else
                {

                    checkBox_use_total_length_for_dj.Checked = false;
                }

                Wgen_main_form.use_pmc_as_cc = false;
                set_checkbox_pmc_as_cc(false);
                for (int i = 0; i < dt_display.Rows.Count; ++i)
                {
                    if (dt_display.Rows[i]["FEATURE CODE"] != DBNull.Value && dt_display.Rows[i][col_pmc_as_cc] != DBNull.Value)
                    {

                        if (Convert.ToString(dt_display.Rows[i]["FEATURE CODE"]) == "PIPE_MATERIAL_CHANGE")
                        {
                            if (dt_display.Rows[i][col_pmc_as_cc] != DBNull.Value && Convert.ToBoolean(dt_display.Rows[i][col_pmc_as_cc]) == true)
                            {
                                Wgen_main_form.use_pmc_as_cc = true;
                                set_checkbox_pmc_as_cc(true);

                            }
                        }
                    }
                }

            }


        }

        public void set_checkbox_total_length(bool check_on)
        {
            checkBox_use_total_length_for_dj.Checked = check_on;
        }


        public void set_checkbox_pmc_as_cc(bool check_on)
        {
            checkBox_pmc_as_cc.Checked = check_on;
        }

        private void checkBox_use_total_length_for_dj_CheckedChanged(object sender, EventArgs e)
        {




            if (checkBox_use_total_length_for_dj.Checked == true)
            {
                Wgen_main_form.dj_total_length = true;
            }
            else
            {
                Wgen_main_form.dj_total_length = false;
            }


            for (int i = 0; i < Wgen_main_form.dt_feature_codes.Rows.Count; ++i)
            {
                if (Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[i][0]) == Wgen_main_form.client_name)
                {
                    Wgen_main_form.dt_feature_codes.Rows[i]["DJ TOTAL LENGTH"] = Wgen_main_form.dj_total_length;
                }

            }

        }

        private void checkBox_pmc_as_cc_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_pmc_as_cc.Checked == true)
            {
                Wgen_main_form.use_pmc_as_cc = true;
            }
            else
            {
                Wgen_main_form.use_pmc_as_cc = false;
            }


            for (int i = 0; i < Wgen_main_form.dt_feature_codes.Rows.Count; ++i)
            {
                if (Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[i][0]) == Wgen_main_form.client_name && Wgen_main_form.dt_feature_codes.Rows[i]["FEATURE CODE"] != DBNull.Value)
                {

                    if (Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[i]["FEATURE CODE"]) == "PIPE_MATERIAL_CHANGE")
                    {
                        Wgen_main_form.dt_feature_codes.Rows[i][col_pmc_as_cc] = Wgen_main_form.use_pmc_as_cc;
                    }
                }

            }
        }
    }
}
