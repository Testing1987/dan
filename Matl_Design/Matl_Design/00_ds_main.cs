using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Collections.Generic;

namespace Alignment_mdi
{
    public partial class ds_main : Form
    {
        private bool clickdragdown;
        private Point lastLocation;
        public static Mat_Design_form tpage_mat_design = null;
        public static Centerline_form tpage_centerline = null;

        public static Export_form tpage_export = null;
        public static Blank_Form tpage_blank = null;

        public static System.Data.DataTable dt_centerline = null;
        public static System.Data.DataTable dt_pipe = null;
        public static System.Data.DataTable dt_extra = null;
        public static System.Data.DataTable dt_points = null;
        public static ds_main tpage_main = null;

        public static bool is3D = false;

        public static string col_Cat = "Category";

        public static string config_xls = "";
        public static string centerline_xls = "";


        string col_elbow = "ELBOW";
        string col_mat_elbow = "Elbow Item No";
        string col_MSblock = "MS Block";
        string col_2dsta = "2DSta";
        string col_3dsta = "3DSta";
        string col_x = "X";
        string col_y = "Y";
        string col_item_no = "ItemNo";
        string col_category = "Category";
        string col_type = "Type";

        int nr_max = 150000;

        public static string client1 = "";
        public static string project1 = "";
        public static string segment1 = "";
        public static string diam1 = "";

        private void make_variables_null()
        {
            dt_centerline = null;
            dt_points = null;
            dt_extra = null;
            is3D = false;
            tpage_mat_design.dt_mat_library = null;
            tpage_mat_design.dt_filter = null;
            ds_main.tpage_mat_design.ct_list = null;
            ds_main.tpage_centerline.dt_cl_display = null;
            ds_main.tpage_centerline.dt_cl_display_filtered = null;
            ds_main.dt_centerline = null;
            tpage_mat_design.dt_ct = null;
          
            config_xls = "";
        }

        public ds_main()
        {
            InitializeComponent();
            tpage_main = this;
            textBox_client_name.Focus();

            tpage_mat_design = new Mat_Design_form();
            tpage_mat_design.MdiParent = this;
            tpage_mat_design.Dock = DockStyle.Fill;
            tpage_mat_design.Hide();

            tpage_centerline = new Centerline_form();
            tpage_centerline.MdiParent = this;
            tpage_centerline.Dock = DockStyle.Fill;
            tpage_centerline.Hide();

            tpage_export = new Export_form();
            tpage_export.MdiParent = this;
            tpage_export.Dock = DockStyle.Fill;
            tpage_export.Hide();


            tpage_blank = new Blank_Form();
            tpage_blank.MdiParent = this;
            tpage_blank.Dock = DockStyle.Fill;
            tpage_blank.Show();


            //sets the mdi background color at runtime
            foreach (Control ctrl in this.Controls)
            {
                if (ctrl is MdiClient)
                {
                    ctrl.BackColor = Color.FromArgb(37, 37, 38);
                }
            }

            treeView1.ShowPlusMinus = false;
        }

        [CommandMethod("MD")]
        public void ShowForm()
        {
            if (Functions.isSECURE() == true)
            {
                foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
                {
                    if (Forma1 is Alignment_mdi.ds_main)
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
                    Alignment_mdi.ds_main forma2 = new Alignment_mdi.ds_main();
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


        protected override void OnLoad(EventArgs e)
        {
            // Hides the ugly border around the mdi container (main form)
            var mdiclient = this.Controls.OfType<MdiClient>().Single();
            this.SuspendLayout();
            mdiclient.SuspendLayout();
            var hdiff = mdiclient.Size.Width - mdiclient.ClientSize.Width;
            var vdiff = mdiclient.Size.Height - mdiclient.ClientSize.Height;
            var size = new Size(mdiclient.Width + hdiff, mdiclient.Height + vdiff);
            var location = new Point(mdiclient.Left - (hdiff / 2), mdiclient.Top - (vdiff / 2));
            mdiclient.Dock = DockStyle.None;
            mdiclient.Size = size;
            mdiclient.Location = location;
            mdiclient.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom;
            mdiclient.ResumeLayout(true);
            this.ResumeLayout(true);
            base.OnLoad(e);
        }


        private void clickmove_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            clickdragdown = true;
            lastLocation = e.Location;
        }

        private void clickmove_MouseMove(object sender, MouseEventArgs e)
        {
            if (clickdragdown)
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


            make_variables_null();


            this.Close();
        }
        private void button_minimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }


        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            switch (e.Node.Text)
            {
                case "Material Design":
                    tpage_mat_design.Show();
                    tpage_centerline.Hide();
                    tpage_export.Hide();
                    tpage_blank.Hide();
                    break;
                case "Route Design":
                    tpage_mat_design.Hide();
                    tpage_centerline.Show();
                    tpage_export.Hide();
                    tpage_blank.Hide();
                    break;
                case "Export":
                    tpage_mat_design.Hide();
                    tpage_centerline.Hide();
                    tpage_export.Show();
                    tpage_blank.Hide();
                    break;
                default:
                    tpage_mat_design.Hide();
                    tpage_centerline.Hide();
                    tpage_export.Hide();
                    tpage_blank.Show();
                    break;
            }


        }




        public string get_textbox_client_name()
        {
            return textBox_client_name.Text;
        }

        public string get_textbox_pipe_diam()
        {
            return textBox_pipe_diam.Text;
        }

        public string get_textbox_project()
        {
            return textBox_project.Text;
        }
        public string get_textbox_segment()
        {
            return textBox_segment.Text;
        }

        public void set_textbox_client_name(string continut)
        {
            textBox_client_name.Text = continut;
        }

        public void set_textbox_pipe_diam(string continut)
        {
            textBox_pipe_diam.Text = continut;
        }

        public void set_textbox_project(string continut)
        {
            textBox_project.Text = continut;
        }
        public void set_textbox_segment(string continut)
        {
            textBox_segment.Text = continut;
        }

        private void radioButton_load_design_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton_load_design.Checked == true)
            {
                button_load_config.Visible = true;
                button_save_config.Visible = false;
            }
            else
            {
                button_load_config.Visible = false;
                button_save_config.Visible = true;
            }
        }

        private void button_load_config_Click(object sender, EventArgs e)
        {
            using (System.Windows.Forms.OpenFileDialog fbd = new System.Windows.Forms.OpenFileDialog())
            {
                fbd.Multiselect = false;
                fbd.Filter = "Excel files (*.xlsx)|*.xlsx";

                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    config_xls = fbd.FileName;
                }
                else
                {

                    return;
                }
            }

            if (System.IO.File.Exists(config_xls) == true)
            {
                Load_centerline_and_mat_library();
                populate_client_project_segment_pipe_diam(client1, project1, segment1, diam1);
            }

        }

        public void Load_centerline_and_mat_library()
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application Excel1 = null;
                Microsoft.Office.Interop.Excel.Workbook Workbook_cfg = null;
                Microsoft.Office.Interop.Excel.Worksheet W_cfg = null;
                Microsoft.Office.Interop.Excel.Workbook Workbook_mat_lib = null;
                Microsoft.Office.Interop.Excel.Worksheet W_cl = null;

                Microsoft.Office.Interop.Excel.Worksheet W_m_desc = null;
                Microsoft.Office.Interop.Excel.Worksheet W_m_pipe = null;
                Microsoft.Office.Interop.Excel.Worksheet W_m_pts = null;
                Microsoft.Office.Interop.Excel.Worksheet W_m_oth = null;
                Microsoft.Office.Interop.Excel.Workbook Workbook_cl = null;

                bool is_opened = false;
                bool is_opened_cl = false;
                try
                {
                    Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    foreach (Microsoft.Office.Interop.Excel.Workbook Workbookx in Excel1.Workbooks)
                    {
                        foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workbookx.Worksheets)
                        {
                            if (Workbookx.FullName == config_xls)
                            {
                                Workbook_cfg = Workbookx;
                                if (Wx.Name == "MDConfig")
                                {
                                    W_cfg = Wx;
                                }
                                if (Wx.Name == "MatDesc")
                                {
                                    W_m_desc = Wx;
                                }
                                if (Wx.Name == "MatPipe")
                                {
                                    W_m_pipe = Wx;
                                }
                                if (Wx.Name == "MatPoints")
                                {
                                    W_m_pts = Wx;
                                }
                                if (Wx.Name == "MatOther")
                                {
                                    W_m_oth = Wx;
                                }

                                is_opened = true;
                            }
                        }
                    }
                }
                catch (System.Exception)
                {
                    Excel1 = new Microsoft.Office.Interop.Excel.Application();
                    Excel1.Visible = true;
                }


                bool save_file = false;
                if (is_opened == false)
                {
                    Workbook_cfg = Excel1.Workbooks.Open(config_xls);
                    foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workbook_cfg.Worksheets)
                    {
                        if (Wx.Name == "MDConfig")
                        {
                            W_cfg = Wx;
                        }

                        if (Wx.Name == "MatDesc")
                        {
                            W_m_desc = Wx;
                        }
                        if (Wx.Name == "MatPipe")
                        {
                            W_m_pipe = Wx;
                        }
                        if (Wx.Name == "MatPoints")
                        {
                            W_m_pts = Wx;
                        }
                        if (Wx.Name == "MatOther")
                        {
                            W_m_oth = Wx;
                        }


                    }
                }
                if (W_cfg != null)
                {

                    client1 = Convert.ToString(W_cfg.Range["B1"].Value2);
                    project1 = Convert.ToString(W_cfg.Range["B2"].Value2);
                    segment1 = Convert.ToString(W_cfg.Range["B3"].Value2);
                    diam1 = Convert.ToString(W_cfg.Range["B4"].Value2);
                    centerline_xls = Convert.ToString(W_cfg.Range["B5"].Value2);


                    #region centerline

                    if (System.IO.File.Exists(centerline_xls) == true)
                    {

                        foreach (Microsoft.Office.Interop.Excel.Workbook Workbookx in Excel1.Workbooks)
                        {
                            if (Workbookx.FullName == centerline_xls)
                            {
                                foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workbookx.Worksheets)
                                {
                                    Workbook_cl = Workbookx;
                                    if (Wx.Name == "CenterLine")
                                    {
                                        W_cl = Wx;
                                    }
                                    is_opened_cl = true;
                                }
                            }
                        }
                        if (is_opened_cl == false)
                        {
                            Workbook_cl = Excel1.Workbooks.Open(centerline_xls);
                            foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workbook_cl.Worksheets)
                            {
                                if (Wx.Name == "CenterLine")
                                {
                                    W_cl = Wx;
                                }
                            }
                        }
                        if (W_cl != null)
                        {
                            ds_main.dt_centerline = tpage_centerline.Build_Data_table_centerline_from_excel(W_cl, 10);
                            tpage_centerline.dt_cl_display = null;
                            tpage_centerline.dt_cl_display_filtered = null;
                            tpage_centerline.set_label_contents(centerline_xls);
                            tpage_centerline.populate_datagridview_cl();
                        }

                        if (is_opened_cl == false)
                        {
                            Workbook_cl.Close();
                        }

                    }

                    #endregion




                    if (W_m_desc != null)
                    {
                        tpage_mat_design.dt_mat_library = tpage_mat_design.Creaza_mat_library_structure();

                        ds_main.dt_pipe = null;
                        ds_main.dt_points = null;
                        ds_main.dt_extra = null;

                        if (ds_main.tpage_centerline.dt_cl_display != null && ds_main.tpage_centerline.dt_cl_display.Rows.Count > 0)
                        {
                            for (int i = 0; i < ds_main.tpage_centerline.dt_cl_display.Rows.Count; ++i)
                            {
                                ds_main.tpage_centerline.dt_cl_display.Rows[i][col_elbow] = DBNull.Value;
                                ds_main.tpage_centerline.dt_cl_display.Rows[i][col_mat_elbow] = DBNull.Value;
                            }
                        }

                        ds_main.tpage_centerline.dt_cl_display_filtered = ds_main.tpage_centerline.dt_cl_display;

                        tpage_mat_design.load_bom(config_xls, W_m_desc, W_m_pipe, W_m_pts, W_m_oth, true);

                        if (tpage_mat_design.dt_mat_library == null || tpage_mat_design.dt_mat_library.Rows.Count == 0)
                        {
                            tpage_mat_design.set_textBox_library_to_red();

                        }

                        List<string> lista1 = Functions.get_blocks_from_current_drawing();

                        for (int i = 0; i < tpage_mat_design.dt_mat_library.Rows.Count; ++i)
                        {
                            if (tpage_mat_design.dt_mat_library.Rows[i][col_MSblock] != DBNull.Value)
                            {
                                string bn = Convert.ToString(tpage_mat_design.dt_mat_library.Rows[i][col_MSblock]);

                                if (lista1.Contains(bn) == false)
                                {
                                    MessageBox.Show("the block " + bn + " not present in current drawing", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }

                            }
                        }


                        try
                        {
                            tpage_mat_design.ct_list = tpage_mat_design.build_category_and_type_list_and_dt_ct();

                            tpage_mat_design.add_tab_pages();

                            tpage_mat_design.populate_datagridview_pipe();
                            ds_main.tpage_centerline.populate_datagridview_cl();
                            ds_main.tpage_centerline.add_elbows_mat_to_combobox(tpage_mat_design.dt_mat_library);

                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                            tpage_mat_design.set_textBox_library_to_red();
                        }


                    }


                }
                try
                {
                    if (is_opened == false)
                    {
                        if (save_file == true) Workbook_cfg.Save();
                        Workbook_cfg.Close();

                        if (Excel1.Workbooks.Count == 0)
                        {
                            Excel1.Quit();
                        }
                        else
                        {
                            Excel1.Visible = true;
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);
                }
                finally
                {
                    if (W_cfg != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W_cfg);
                    if (W_cl != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W_cl);
                    if (W_m_desc != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W_m_desc);
                    if (W_m_pipe != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W_m_pipe);
                    if (W_m_pts != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W_m_pts);
                    if (W_m_oth != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W_m_oth);
                    if (Workbook_cfg != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook_cfg);
                    if (Workbook_mat_lib != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook_mat_lib);
                    if (Workbook_cl != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook_cl);
                    if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        public void populate_client_project_segment_pipe_diam(string client, string project, string segment, string pipe_diam)
        {
            textBox_client_name.Text = client;
            textBox_pipe_diam.Text = pipe_diam;
            textBox_segment.Text = segment;
            textBox_project.Text = project;
        }

        public string get_textBox_client_name()
        {
            return textBox_client_name.Text;
        }
        public string get_textBox_pipe_diam()
        {
            return textBox_pipe_diam.Text;
        }

        public string get_textBox_segment()
        {
            return textBox_segment.Text;
        }

        public string get_textBox_project()
        {
            return textBox_project.Text;
        }


        private void button_save_config_Click(object sender, EventArgs e)
        {
            bool is_excel_opened = true;

            try
            {
                Microsoft.Office.Interop.Excel.Application Excel1 = null;

                try
                {
                    Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    if (Excel1.Workbooks.Count == 0)
                    {
                        try
                        {
                            Excel1.Quit();
                        }
                        catch (System.Exception)
                        {

                        }
                        finally
                        {
                            if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                        }
                        Excel1 = new Microsoft.Office.Interop.Excel.Application();
                        is_excel_opened = false;
                    }

                }
                catch (System.Exception)
                {
                    Excel1 = new Microsoft.Office.Interop.Excel.Application();
                    is_excel_opened = false;
                }



                Microsoft.Office.Interop.Excel.Workbook Workbook_cfg = null;
                Microsoft.Office.Interop.Excel._Worksheet W_cfg = null;


                try
                {
                    SaveFileDialog Save_dlg = new SaveFileDialog();
                    Save_dlg.Filter = "Excel file|*.xlsx";

                    get_client_project_segment_pipe_diam();

                    if (Save_dlg.ShowDialog() == DialogResult.OK)
                    {
                        config_xls = Save_dlg.FileName;


                        if (System.IO.File.Exists(config_xls) == true)
                        {
                            foreach (Microsoft.Office.Interop.Excel.Workbook Workbook2 in Excel1.Workbooks)
                            {
                                if (Workbook2.FullName == config_xls)
                                {
                                    Workbook_cfg = Workbook2;
                                    is_excel_opened = true;
                                    foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workbook_cfg.Worksheets)
                                    {
                                        if (Wx.Name == "MDConfig")
                                        {
                                            W_cfg = Wx;
                                        }
                                    }
                                }
                            }

                            if (Workbook_cfg == null)
                            {
                                Workbook_cfg = Excel1.Workbooks.Open(config_xls);
                                foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workbook_cfg.Worksheets)
                                {
                                    if (Wx.Name == "MDConfig")
                                    {
                                        W_cfg = Wx;
                                    }
                                }
                            }

                            if (W_cfg == null)
                            {
                                W_cfg = Workbook_cfg.Worksheets.Add(System.Reflection.Missing.Value, Workbook_cfg.Worksheets[1], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                                W_cfg.Name = "MDConfig";
                            }

                        }
                        else
                        {
                            Workbook_cfg = Excel1.Workbooks.Add();
                            W_cfg = Workbook_cfg.Worksheets[1];
                            W_cfg.Name = "MDConfig";
                        }

                        if (W_cfg != null)
                        {

                            W_cfg.Range["A:A"].ColumnWidth = 28;
                            W_cfg.Range["B:B"].ColumnWidth = 100;
                            W_cfg.Range["A1"].Value2 = "Client Name";
                            W_cfg.Range["A2"].Value2 = "Project Name";
                            W_cfg.Range["A3"].Value2 = "Segment Name";
                            W_cfg.Range["A4"].Value2 = "Pipe Diameter";
                            W_cfg.Range["A5"].Value2 = "Centerline File Location";
                            W_cfg.Range["A6"].Value2 = "Material Library File Location";

                            W_cfg.Range["B1"].Value2 = client1;
                            W_cfg.Range["B2"].Value2 = project1;
                            W_cfg.Range["B3"].Value2 = segment1;
                            W_cfg.Range["B4"].Value2 = diam1;
                            W_cfg.Range["B5"].Value2 = centerline_xls;
                            W_cfg.Range["B6"].Value2 = config_xls;

                            if (tpage_mat_design.ct_list != null && tpage_mat_design.ct_list.Count > 0)
                            {
                                if (ds_main.dt_points != null && ds_main.dt_points.Rows.Count > 0)
                                {
                                    string ct1 = "ELL_POINT";

                                    if (tpage_mat_design.ct_list.Contains(ct1) == true)
                                    {
                                        System.Data.DataTable dt1 = tpage_mat_design.dt_ct[tpage_mat_design.ct_list.IndexOf(ct1)];
                                        List<string> lista1 = new List<string>();

                                        for (int j = 0; j < tpage_mat_design.dt_mat_library.Rows.Count; ++j)
                                        {
                                            if (tpage_mat_design.dt_mat_library.Rows[j][col_category] != DBNull.Value && tpage_mat_design.dt_mat_library.Rows[j][col_type] != DBNull.Value && tpage_mat_design.dt_mat_library.Rows[j][col_item_no] != DBNull.Value)
                                            {
                                                if ((Convert.ToString(tpage_mat_design.dt_mat_library.Rows[j][col_category]) + "_" + Convert.ToString(tpage_mat_design.dt_mat_library.Rows[j][col_type])).ToUpper().Replace(" ", "") == ct1)
                                                {
                                                    string mat1 = Convert.ToString(tpage_mat_design.dt_mat_library.Rows[j][col_item_no]).ToUpper().Replace(" ", "");
                                                    lista1.Add(mat1);
                                                }
                                            }
                                        }

                                        for (int i = ds_main.dt_points.Rows.Count - 1; i >= 0; --i)
                                        {
                                            if (ds_main.dt_points.Rows[i][col_item_no] != DBNull.Value)
                                            {
                                                string mat1 = Convert.ToString(ds_main.dt_points.Rows[i][col_item_no]).ToUpper().Replace(" ", "");
                                                if (lista1.Contains(mat1) == true)
                                                {
                                                    ds_main.dt_points.Rows[i].Delete();
                                                }
                                            }
                                        }

                                        for (int j = 0; j < dt1.Rows.Count; ++j)
                                        {
                                            ds_main.dt_points.Rows.Add();
                                            for (int k = 0; k < dt1.Columns.Count; ++k)
                                            {
                                                ds_main.dt_points.Rows[ds_main.dt_points.Rows.Count - 1][k] = dt1.Rows[j][k];
                                            }
                                        }

                                        if (ds_main.dt_points.Rows.Count > 0)
                                        {
                                            string col1 = col_2dsta;
                                            if (ds_main.dt_points.Rows[0][col_3dsta] != DBNull.Value) col1 = col_3dsta;

                                            ds_main.dt_points = Functions.Sort_data_table(ds_main.dt_points, col1);
                                        }
                                    }

                                    if (ds_main.dt_centerline != null && ds_main.dt_centerline.Rows.Count > 1)
                                    {
                                        Polyline Poly2D = Functions.Build_2d_poly_for_scanning(ds_main.dt_centerline);
                                        for (int i = 0; i < ds_main.dt_points.Rows.Count; ++i)
                                        {
                                            if (ds_main.dt_points.Rows[i][col_2dsta] != DBNull.Value && (ds_main.dt_points.Rows[i][col_x] == DBNull.Value || ds_main.dt_points.Rows[i][col_y] == DBNull.Value))
                                            {
                                                double sta1 = Convert.ToDouble(ds_main.dt_points.Rows[i][col_2dsta]);
                                                if (sta1 < 0) sta1 = 0;
                                                if (sta1 >= Poly2D.Length) sta1 = Poly2D.Length - 0.001;

                                                Point3d pt_on_poly = Poly2D.GetPointAtDist(sta1);
                                                ds_main.dt_points.Rows[i][col_x] = pt_on_poly.X;
                                                ds_main.dt_points.Rows[i][col_y] = pt_on_poly.Y;
                                            }
                                        }
                                    }


                                    tpage_mat_design.set_textBox_library_content();
                                }
                            }

                            Transfer_all_datatables_to_config(Excel1, Workbook_cfg);

                            if (System.IO.File.Exists(config_xls) == true)
                            {
                                Workbook_cfg.Save();
                            }
                            else
                            {
                                Workbook_cfg.SaveAs(config_xls);
                            }

                            if (is_excel_opened == false) Workbook_cfg.Close();
                            if (Excel1.Workbooks.Count == 0)
                            {
                                Excel1.Quit();
                            }
                            else
                            {
                                Excel1.Visible = true;
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);

                }
                finally
                {
                    if (W_cfg != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W_cfg);
                    if (Workbook_cfg != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook_cfg);
                    if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }

        }
        public void get_client_project_segment_pipe_diam()
        {
            ds_main.client1 = ds_main.tpage_main.get_textBox_client_name();
            ds_main.diam1 = ds_main.tpage_main.get_textBox_pipe_diam();
            ds_main.segment1 = ds_main.tpage_main.get_textBox_segment();
            ds_main.project1 = ds_main.tpage_main.get_textBox_project();
        }

        public void Transfer_all_datatables_to_config(Microsoft.Office.Interop.Excel.Application Excel1, Microsoft.Office.Interop.Excel.Workbook Workbook1)
        {

            Microsoft.Office.Interop.Excel.Worksheet W_mat_lib = null;
            Microsoft.Office.Interop.Excel.Worksheet W_mat_pipe = null;
            Microsoft.Office.Interop.Excel.Worksheet W_mat_pts = null;
            Microsoft.Office.Interop.Excel.Worksheet W_mat_oth = null;



            try
            {

                if (System.IO.File.Exists(ds_main.config_xls) == true)
                {




                    foreach (Microsoft.Office.Interop.Excel.Workbook Workbook2 in Excel1.Workbooks)
                    {
                        if (Workbook2.FullName == ds_main.config_xls)
                        {
                            foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workbook2.Worksheets)
                            {

                                Workbook1 = Workbook2;
                                if (Wx.Name == "MatDesc")
                                {
                                    W_mat_lib = Wx;
                                }
                                if (Wx.Name == "MatPipe")
                                {
                                    W_mat_pipe = Wx;
                                }
                                if (Wx.Name == "MatPoints")
                                {
                                    W_mat_pts = Wx;
                                }
                                if (Wx.Name == "MatOther")
                                {
                                    W_mat_oth = Wx;
                                }


                            }


                        }
                    }



                    if (W_mat_lib == null)
                    {
                        W_mat_lib = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[1], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                        W_mat_lib.Name = "MatDesc";
                    }
                    if (W_mat_pipe == null)
                    {
                        W_mat_pipe = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[2], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                        W_mat_pipe.Name = "MatPipe";
                    }
                    if (W_mat_pts == null)
                    {
                        W_mat_pts = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[3], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                        W_mat_pts.Name = "MatPoints";
                    }

                    if (W_mat_oth == null)
                    {
                        W_mat_oth = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[4], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                        W_mat_oth.Name = "MatOther";
                    }


                }

                if (W_mat_lib != null || W_mat_pipe != null || W_mat_pts != null || W_mat_oth != null)
                {
                    if (tpage_mat_design.dt_mat_library != null && tpage_mat_design.dt_mat_library.Rows.Count > 0)
                    {
                        Mat_Design_form.Create_header_material_library(W_mat_lib, ds_main.client1, ds_main.project1, ds_main.segment1, tpage_mat_design.dt_mat_library);

                        W_mat_lib.Cells.NumberFormat = "General";
                        int maxRows = tpage_mat_design.dt_mat_library.Rows.Count;
                        int maxCols = tpage_mat_design.dt_mat_library.Columns.Count;
                        W_mat_lib.Range["A14:G1000"].ClearContents();
                        W_mat_lib.Range["A14:G1000"].ClearFormats();

                        Microsoft.Office.Interop.Excel.Range range1 = W_mat_lib.Range["A14:G" + (14 + maxRows - 1).ToString()];
                        object[,] values1 = new object[maxRows, maxCols];

                        for (int i = 0; i < maxRows; ++i)
                        {
                            for (int j = 0; j < maxCols; ++j)
                            {
                                if (tpage_mat_design.dt_mat_library.Rows[i][j] != DBNull.Value && j > 0)// i did not want to save mmid value
                                {
                                    values1[i, j] = tpage_mat_design.dt_mat_library.Rows[i][j];
                                }
                            }
                        }
                        range1.Value2 = values1;

                    }


                    if (ds_main.dt_pipe != null && ds_main.dt_pipe.Rows.Count > 0)
                    {
                        Mat_Design_form.Create_header_material_linear_file(W_mat_pipe, ds_main.client1, ds_main.project1, ds_main.segment1, ds_main.dt_pipe);

                        int last_row = nr_max + 14;
                        W_mat_pipe.Cells.NumberFormat = "General";
                        int maxRows = ds_main.dt_pipe.Rows.Count;
                        int maxCols = ds_main.dt_pipe.Columns.Count;
                        W_mat_pipe.Range["A14:V" + last_row.ToString()].ClearContents();
                        W_mat_pipe.Range["A14:V" + last_row.ToString()].ClearFormats();

                        Microsoft.Office.Interop.Excel.Range range1 = W_mat_pipe.Range["A14:V" + (14 + maxRows - 1).ToString()];
                        object[,] values1 = new object[maxRows, maxCols];

                        for (int i = 0; i < maxRows; ++i)
                        {
                            for (int j = 0; j < maxCols; ++j)
                            {
                                if (ds_main.dt_pipe.Rows[i][j] != DBNull.Value)
                                {
                                    values1[i, j] = ds_main.dt_pipe.Rows[i][j];
                                }
                            }
                        }
                        range1.Value2 = values1;

                    }



                    if (ds_main.dt_points != null && ds_main.dt_points.Rows.Count > 0)
                    {


                        Mat_Design_form.Create_header_material_points_file(W_mat_pts, ds_main.client1, ds_main.project1, ds_main.segment1, ds_main.dt_points);
                        int last_row = nr_max + 13;
                        W_mat_pts.Cells.NumberFormat = "General";
                        int maxRows = ds_main.dt_points.Rows.Count;
                        int maxCols = ds_main.dt_points.Columns.Count;
                        W_mat_pts.Range["A13:Q" + last_row.ToString()].ClearContents();
                        W_mat_pts.Range["A13:Q" + last_row.ToString()].ClearFormats();

                        Microsoft.Office.Interop.Excel.Range range1 = W_mat_pts.Range["A13:Q" + (13 + maxRows - 1).ToString()];
                        object[,] values1 = new object[maxRows, maxCols];

                        for (int i = 0; i < maxRows; ++i)
                        {
                            for (int j = 0; j < maxCols; ++j)
                            {
                                if (ds_main.dt_points.Rows[i][j] != DBNull.Value)
                                {
                                    values1[i, j] = ds_main.dt_points.Rows[i][j];
                                }
                            }
                        }
                        range1.Value2 = values1;
                    }

                    if (ds_main.dt_extra != null && ds_main.dt_extra.Rows.Count > 0)
                    {
                        Mat_Design_form.Create_header_material_linear_file(W_mat_oth, ds_main.client1, ds_main.project1, ds_main.segment1, ds_main.dt_extra, "Material Linear Other");

                        int last_row = nr_max + 14;
                        W_mat_oth.Cells.NumberFormat = "General";
                        int maxRows = ds_main.dt_extra.Rows.Count;
                        int maxCols = ds_main.dt_extra.Columns.Count;
                        W_mat_oth.Range["A14:V" + last_row.ToString()].ClearContents();
                        W_mat_oth.Range["A14:V" + last_row.ToString()].ClearFormats();

                        Microsoft.Office.Interop.Excel.Range range1 = W_mat_oth.Range["A14:V" + (14 + maxRows - 1).ToString()];
                        object[,] values1 = new object[maxRows, maxCols];

                        for (int i = 0; i < maxRows; ++i)
                        {
                            for (int j = 0; j < maxCols; ++j)
                            {
                                if (ds_main.dt_extra.Rows[i][j] != DBNull.Value)
                                {
                                    values1[i, j] = ds_main.dt_extra.Rows[i][j];
                                }
                            }
                        }
                        range1.Value2 = values1;

                    }





                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }



        }

    }
}
