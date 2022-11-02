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
using Autodesk.AutoCAD.EditorInput;


namespace Alignment_mdi
{
    public partial class Wgen_main_form : Form
    {
        bool clickdragdown;
        Point lastLocation;

        static public List<string> col_labels_zoom;

        static public List<string> col_station_labels;

        public static Wgen_pipemanifest tpage_pipe_manifest = null;
        public static Wgen_pipetally tpage_pipe_tally = null;
        public static Wgen_Blank_form tpage_blank = null;
        public static Wgen_weldmap tpage_weldmap = null;
        public static Wgen_all_pts tpage_allpts = null;
        public static Wgen_pipe_tally tpage_build_pipe_tally = null;
        public static Wgen_duplicates tpage_duplicates = null;
        public static Wgen_feature tpage_features = null;

        public static System.Data.DataTable dt_ground_tally = null;

        public static System.Data.DataTable dt_double_joint = null;


        public static System.Data.DataTable dt_pipe_list = null;
        public static System.Data.DataTable dt_weld_map = null;
        public static System.Data.DataTable dt_all_points = null;
        public static System.Data.DataTable dt_pt_keep = null;
        public static System.Data.DataTable dt_pt_move = null;
        public static System.Data.DataTable dt_pt_resolved = null;
        public static System.Data.DataTable dt_pt_processed = null;
        public static List<string> lista_feature_code = null;
        public static List<string> lista_feature_code_exception = null;

        public static double pipe_diam = 0;


        public static System.Data.DataTable dt_feature_codes;

        public static List<string> lista_clienti = null;
        public static string client_name = "xxx";
        public static bool incomplete_pipe_manifest = false;

        public static string WGEN_folder = Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments) + "\\WGEN\\";

        public Wgen_main_form()
        {
            InitializeComponent();
            col_labels_zoom = new List<string>();
            col_station_labels = new List<string>();

            int nr_excel_open = Functions.Get_no_of_workbooks_from_Excel();


            //sets the mdi background color at runtime
            foreach (Control ctrl in this.Controls)
            {
                if (ctrl is MdiClient)
                {
                    ctrl.BackColor = Color.FromArgb(62, 62, 66);
                }
            }

            // if (Functions.is_dan_popescu() == false) treeView_inquiry.Nodes[4].ForeColor = Color.Gray;

            lista_feature_code = new List<string>();

            string file1 = WGEN_folder+ "wgen_feature_codes.xlsx";
            bool visible1 = false;
            if (Functions.is_dan_popescu() == true) visible1 = true;
            lista_clienti = new List<string>();
            int index_sheet = 1;

            dt_feature_codes = new System.Data.DataTable();
            dt_feature_codes.Columns.Add("CLIENT", typeof(string));
            dt_feature_codes.Columns.Add("FEATURE CODE", typeof(string));
            dt_feature_codes.Columns.Add("INCLUDED (YES/NO)", typeof(bool));
            dt_feature_codes.Columns.Add("DESCRIPTION", typeof(string));
            dt_feature_codes.Columns.Add("CHECK 1", typeof(string));
            dt_feature_codes.Columns.Add("CHECK 2", typeof(string));
            dt_feature_codes.Columns.Add("CHECK 3", typeof(string));
            dt_feature_codes.Columns.Add("CHECK 4", typeof(string));
            dt_feature_codes.Columns.Add("CHECK 5", typeof(string));
            dt_feature_codes.Columns.Add("CHECK 6", typeof(string));
            dt_feature_codes.Columns.Add("CHECK 7", typeof(string));
            dt_feature_codes.Columns.Add("CHECK 8", typeof(string));
            dt_feature_codes.Columns.Add("CHECK 9", typeof(string));
            dt_feature_codes.Columns.Add("CHECK 10", typeof(string));
            dt_feature_codes.Columns.Add("CHECK XRAY", typeof(string));
            dt_feature_codes.Columns.Add("BEND TYPE", typeof(string));
            dt_feature_codes.Columns.Add("BEND DEFLECTION TYPE", typeof(string));
            dt_feature_codes.Columns.Add("BEND POSITION", typeof(string));
            dt_feature_codes.Columns.Add("BEND HORIZONTAL DEFLECTION", typeof(string));
            dt_feature_codes.Columns.Add("BEND VERTICAL DEFLECTION", typeof(string));
            dt_feature_codes.Columns.Add("WELD MM BACK", typeof(string));
            dt_feature_codes.Columns.Add("WELD MM AHEAD", typeof(string));

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
                    if (Excel1 == null)
                    {
                        MessageBox.Show("PROBLEM WITH EXCEL!");
                        return;
                    }
                    Excel1.Visible = visible1;
                    Workbook1 = Excel1.Workbooks.Open(file1);
                    foreach (Microsoft.Office.Interop.Excel.Worksheet W1 in Workbook1.Worksheets)
                    {
                        string nume1 = W1.Name;
                        if (lista_clienti.Contains(W1.Name) == false) lista_clienti.Add(nume1);

                        Microsoft.Office.Interop.Excel.Range range2 = W1.Range["A2: U30000"];
                        object[,] values2 = new object[30000, 21];
                        values2 = range2.Value2;
                        for (int i = 1; i <= values2.Length / 3; ++i)
                        {
                            object valA = values2[i, 1];
                            object valB = values2[i, 2];
                            object valC = values2[i, 3];
                            object valD = values2[i, 4];
                            object valE = values2[i, 5];
                            object valF = values2[i, 6];
                            object valG = values2[i, 7];
                            object valH = values2[i, 8];
                            object valI = values2[i, 9];
                            object valJ = values2[i, 10];
                            object valK = values2[i, 11];
                            object valL = values2[i, 12];
                            object valM = values2[i, 13];
                            object valN = values2[i, 14];
                            object valO = values2[i, 15];
                            object valP = values2[i, 16];
                            object valQ = values2[i, 17];
                            object valR = values2[i, 18];
                            object valS = values2[i, 19];
                            object valT = values2[i, 20];
                            object valU = values2[i, 21];

                            if (valA != null)
                            {
                                dt_feature_codes.Rows.Add();
                                dt_feature_codes.Rows[dt_feature_codes.Rows.Count - 1]["CLIENT"] = nume1;
                                dt_feature_codes.Rows[dt_feature_codes.Rows.Count - 1]["FEATURE CODE"] = Convert.ToString(valA);
                                if (valB != null && Convert.ToString(valB).ToUpper() == "YES")
                                {
                                    dt_feature_codes.Rows[dt_feature_codes.Rows.Count - 1]["INCLUDED (YES/NO)"] = true;
                                }
                                else
                                {
                                    dt_feature_codes.Rows[dt_feature_codes.Rows.Count - 1]["INCLUDED (YES/NO)"] = false;
                                }
                                if (valC != null)
                                {
                                    dt_feature_codes.Rows[dt_feature_codes.Rows.Count - 1]["DESCRIPTION"] = Convert.ToString(valC);
                                }
                                else
                                {
                                    dt_feature_codes.Rows[dt_feature_codes.Rows.Count - 1]["DESCRIPTION"] = "{I}";
                                }

                                if (valD != null)
                                {
                                    dt_feature_codes.Rows[dt_feature_codes.Rows.Count - 1]["CHECK 1"] = Convert.ToString(valD);
                                }

                                if (valE != null)
                                {
                                    dt_feature_codes.Rows[dt_feature_codes.Rows.Count - 1]["CHECK 2"] = Convert.ToString(valE);
                                }

                                if (valF != null)
                                {
                                    dt_feature_codes.Rows[dt_feature_codes.Rows.Count - 1]["CHECK 3"] = Convert.ToString(valF);
                                }

                                if (valG != null)
                                {
                                    dt_feature_codes.Rows[dt_feature_codes.Rows.Count - 1]["CHECK 4"] = Convert.ToString(valG);
                                }

                                if (valH != null)
                                {
                                    dt_feature_codes.Rows[dt_feature_codes.Rows.Count - 1]["CHECK 5"] = Convert.ToString(valH);
                                }

                                if (valI != null)
                                {
                                    dt_feature_codes.Rows[dt_feature_codes.Rows.Count - 1]["CHECK 6"] = Convert.ToString(valI);
                                }

                                if (valJ != null)
                                {
                                    dt_feature_codes.Rows[dt_feature_codes.Rows.Count - 1]["CHECK 7"] = Convert.ToString(valJ);
                                }

                                if (valK != null)
                                {
                                    dt_feature_codes.Rows[dt_feature_codes.Rows.Count - 1]["CHECK 8"] = Convert.ToString(valK);
                                }

                                if (valL != null)
                                {
                                    dt_feature_codes.Rows[dt_feature_codes.Rows.Count - 1]["CHECK 9"] = Convert.ToString(valL);
                                }

                                if (valM != null)
                                {
                                    dt_feature_codes.Rows[dt_feature_codes.Rows.Count - 1]["CHECK 10"] = Convert.ToString(valM);
                                }

                                if (valN != null)
                                {
                                    dt_feature_codes.Rows[dt_feature_codes.Rows.Count - 1]["CHECK XRAY"] = Convert.ToString(valN);
                                }

                                if (valO != null)
                                {
                                    dt_feature_codes.Rows[dt_feature_codes.Rows.Count - 1]["BEND TYPE"] = Convert.ToString(valO);
                                }

                                if (valP != null)
                                {
                                    dt_feature_codes.Rows[dt_feature_codes.Rows.Count - 1]["BEND DEFLECTION TYPE"] = Convert.ToString(valP);
                                }

                                if (valQ != null)
                                {
                                    dt_feature_codes.Rows[dt_feature_codes.Rows.Count - 1]["BEND POSITION"] = Convert.ToString(valQ);
                                }

                                if (valR != null)
                                {
                                    dt_feature_codes.Rows[dt_feature_codes.Rows.Count - 1]["BEND HORIZONTAL DEFLECTION"] = Convert.ToString(valR);
                                }

                                if (valS != null)
                                {
                                    dt_feature_codes.Rows[dt_feature_codes.Rows.Count - 1]["BEND VERTICAL DEFLECTION"] = Convert.ToString(valS);
                                }

                                if (valT != null)
                                {
                                    dt_feature_codes.Rows[dt_feature_codes.Rows.Count - 1]["WELD MM BACK"] = Convert.ToString(valT);
                                }

                                if (valU != null)
                                {
                                    dt_feature_codes.Rows[dt_feature_codes.Rows.Count - 1]["WELD MM AHEAD"] = Convert.ToString(valU);
                                }


                            }
                            else
                            {
                                i = values2.Length + 1;
                            }
                        }
                        ++index_sheet;
                    }
                    Workbook1.Close();
                    if (nr_excel_open == 0) Excel1.Quit();
                    // MessageBox.Show("debug");
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
            else
            {
                MessageBox.Show("[wgen_feature_codes.xlsx] not found.");
                return;

            }
            if (client_name == "xxx") client_name = lista_clienti[0];


            lista_feature_code.Clear();
            for (int i = 0; i < dt_feature_codes.Rows.Count; ++i)
            {
                string client = Convert.ToString(dt_feature_codes.Rows[i][0]);
                if (client == client_name)
                {
                    if ((bool)dt_feature_codes.Rows[i][2] == true)
                    {
                        lista_feature_code.Add(Convert.ToString(dt_feature_codes.Rows[i][1]));
                    }
                }
            }


            if (Wgen_main_form.lista_feature_code.Contains("WELD") == false)  Wgen_main_form.lista_feature_code.Add("WELD");
            if (Wgen_main_form.lista_feature_code.Contains("BEND") == false)  Wgen_main_form.lista_feature_code.Add("BEND");
            if (Wgen_main_form.lista_feature_code.Contains("NATURAL_GROUND") == false)  Wgen_main_form.lista_feature_code.Add("NATURAL_GROUND");

            tpage_pipe_manifest = new Wgen_pipemanifest();
            tpage_pipe_manifest.MdiParent = this;
            tpage_pipe_manifest.Dock = DockStyle.Fill;
            tpage_pipe_manifest.Hide();

            tpage_pipe_tally = new Wgen_pipetally();
            tpage_pipe_tally.MdiParent = this;
            tpage_pipe_tally.Dock = DockStyle.Fill;
            tpage_pipe_tally.Hide();

            tpage_allpts = new Wgen_all_pts();
            tpage_allpts.MdiParent = this;
            tpage_allpts.Dock = DockStyle.Fill;
            tpage_allpts.Hide();

            tpage_weldmap = new Wgen_weldmap();
            tpage_weldmap.MdiParent = this;
            tpage_weldmap.Dock = DockStyle.Fill;
            tpage_weldmap.Hide();

            tpage_build_pipe_tally = new Wgen_pipe_tally();
            tpage_build_pipe_tally.MdiParent = this;
            tpage_build_pipe_tally.Dock = DockStyle.Fill;
            tpage_build_pipe_tally.Hide();

            tpage_blank = new Wgen_Blank_form();
            tpage_blank.MdiParent = this;
            tpage_blank.Dock = DockStyle.Fill;
            tpage_blank.Show();

            tpage_duplicates = new Wgen_duplicates();
            tpage_duplicates.MdiParent = this;
            tpage_duplicates.Dock = DockStyle.Fill;
            tpage_duplicates.Hide();

            //Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt_feature_codes);
        }

        private void treeView1_BeforeSelect(object sender, TreeViewCancelEventArgs e)
        {
            if (Color.Gray == e.Node.ForeColor)
                e.Cancel = true;
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
            dt_pipe_list = null;
            dt_ground_tally = null;
            dt_all_points = null;
            dt_weld_map = null;
            dt_pt_keep = null;
            dt_pt_move = null;
            client_name = "xxx";
            this.Close();
        }
        private void button_minimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void treeView_inquiry_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (e.Node.Text == "Pipe Manifest")
            {
                tpage_pipe_manifest.Show();
                tpage_pipe_tally.Hide();
                tpage_weldmap.Hide();
                tpage_blank.Hide();
                tpage_allpts.Hide();
                tpage_duplicates.Hide();
                tpage_build_pipe_tally.Hide();


            }
            if (e.Node.Text == "Pipe Tally")
            {
                tpage_pipe_manifest.Hide();
                tpage_pipe_tally.Show();
                tpage_weldmap.Hide();
                tpage_blank.Hide();
                tpage_allpts.Hide();
                tpage_duplicates.Hide();
                tpage_build_pipe_tally.Hide();

            }
            if (e.Node.Text == "Build Pipe Tally")
            {
                tpage_pipe_manifest.Hide();
                tpage_pipe_tally.Hide();
                tpage_weldmap.Hide();
                tpage_blank.Hide();
                tpage_allpts.Hide();
                tpage_duplicates.Hide();
                tpage_build_pipe_tally.Show();

            }
            if (e.Node.Text == "Weld Map")
            {
                tpage_pipe_manifest.Hide();
                tpage_pipe_tally.Hide();
                tpage_weldmap.Show();
                tpage_blank.Hide();
                tpage_allpts.Hide();
                tpage_duplicates.Hide();
                tpage_build_pipe_tally.Hide();

            }
            if (e.Node.Text == "All Points")
            {
                tpage_pipe_manifest.Hide();
                tpage_pipe_tally.Hide();
                tpage_weldmap.Hide();
                tpage_blank.Hide();
                tpage_allpts.Show();
                tpage_build_pipe_tally.Hide();
                tpage_duplicates.Hide();
            }
            if (e.Node.Text == "Duplicates")
            {
                tpage_pipe_manifest.Hide();
                tpage_pipe_tally.Hide();
                tpage_weldmap.Hide();
                tpage_blank.Hide();
                tpage_allpts.Hide();
                tpage_build_pipe_tally.Hide();
                tpage_duplicates.Show();
            }
        }

        private void label_iq_treeviewnav_Click(object sender, EventArgs e)
        {
            tpage_pipe_manifest.Hide();

            tpage_blank.Show();
        }
    }


}
