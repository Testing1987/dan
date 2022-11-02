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
using Font = System.Drawing.Font;
using Autodesk.Civil.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;

namespace Alignment_mdi
{
    public partial class tcpl_tools_form : Form
    {
        string Col_x = "X";
        string Col_y = "Y";
        string Col_z = "Z";
        string Col_3DSta = "3DSta";
        string Col_BackSta = "BackSta";
        string Col_AheadSta = "AheadSta";
        string segment_current = "";
        string ProjFolder_main = "";

        System.Data.DataTable dt_centerline = null;
        Polyline poly_centerline = null;

        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();

            lista_butoane.Add(button_load_agen_project);
            lista_butoane.Add(button_read_chainage_from_profile_view);
            lista_butoane.Add(comboBox_segments);
            lista_butoane.Add(button_place_points_on_profile);
            lista_butoane.Add(comboBox_surface);
            lista_butoane.Add(button_load_surf);
            lista_butoane.Add(button_place_blocks_on_profile);
            lista_butoane.Add(button_load_blocks);
            lista_butoane.Add(comboBox_blocks);
            lista_butoane.Add(button_project_poly_on_surface);
            lista_butoane.Add(button_calcZ);
            lista_butoane.Add(comboBox_blocks);
            lista_butoane.Add(button_calc_staZ_of_block);
            lista_butoane.Add(comboBox_round_elev);
            lista_butoane.Add(comboBox_round_sta);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = false;
            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();

            lista_butoane.Add(button_load_agen_project);
            lista_butoane.Add(button_read_chainage_from_profile_view);
            lista_butoane.Add(comboBox_segments);
            lista_butoane.Add(button_place_points_on_profile);
            lista_butoane.Add(comboBox_surface);
            lista_butoane.Add(button_load_surf);
            lista_butoane.Add(button_place_blocks_on_profile);
            lista_butoane.Add(button_load_blocks);
            lista_butoane.Add(comboBox_blocks);
            lista_butoane.Add(button_project_poly_on_surface);
            lista_butoane.Add(button_calcZ);
            lista_butoane.Add(comboBox_blocks);
            lista_butoane.Add(button_calc_staZ_of_block);
            lista_butoane.Add(comboBox_round_elev);
            lista_butoane.Add(comboBox_round_sta);


            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }

        public tcpl_tools_form()
        {
            InitializeComponent();

            comboBox_round_sta.SelectedIndex = 1;
            comboBox_round_elev.SelectedIndex = 1;

            var toolStripMenuItem2 = new ToolStripMenuItem { Text = "tool strip menu" };
            //toolStripMenuItem2.Click += go_to_excel_point;


            //ContextMenuStrip_go_to_error = new ContextMenuStrip();
            //ContextMenuStrip_go_to_error.Items.AddRange(new ToolStripItem[] { toolStripMenuItem2 });


        }

        private void button_project_poly_on_surface_Click(object sender, EventArgs e)
        {
            if (comboBox_surface.Text == "")
            {
                MessageBox.Show("no surface selected");
                return;
            }
            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.Civil.ApplicationServices.CivilDocument CivilDrawing = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument;
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
                        ObjectIdCollection col_surf = CivilDrawing.GetSurfaceIds();
                        for (int j = 0; j < col_surf.Count; ++j)
                        {
                            Autodesk.Civil.DatabaseServices.Surface surf1 = Trans1.GetObject(col_surf[j], OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Surface;
                            if (surf1 != null)
                            {
                                if (surf1.Name == comboBox_surface.Text)
                                {
                                    this.MdiParent.WindowState = FormWindowState.Minimized;

                                    Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                                    Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                                    Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect polyline:");
                                    Prompt_centerline.SetRejectMessage("\nSelect a polyline!");
                                    Prompt_centerline.AllowNone = true;
                                    Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                                    Rezultat_centerline = ThisDrawing.Editor.GetEntity(Prompt_centerline);

                                    if (Rezultat_centerline.Status != PromptStatus.OK)
                                    {
                                        Editor1.SetImpliedSelection(Empty_array);
                                        Editor1.WriteMessage("\nCommand:");
                                        this.MdiParent.WindowState = FormWindowState.Normal;

                                        set_enable_true();
                                        return;
                                    }



                                    Polyline poly1 = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForRead) as Polyline;

                                    if (poly1 != null)
                                    {


                                        System.Data.DataTable dt1 = new System.Data.DataTable();
                                        dt1.Columns.Add("Station", typeof(double));
                                        dt1.Columns.Add("X", typeof(double));
                                        dt1.Columns.Add("Y", typeof(double));
                                        dt1.Columns.Add("Z on surface", typeof(double));
                                        dt1.Columns.Add("Surface Name", typeof(string));
                                        dt1.Columns.Add("USER", typeof(string));
                                        dt1.Columns.Add("Station 3D", typeof(double));



                                        for (int i = 0; i < poly1.NumberOfVertices; i++)
                                        {
                                            Point3d pt1 = poly1.GetPointAtParameter(i);

                                            double x = pt1.X;
                                            double y = pt1.Y;
                                            double sta = poly1.GetDistanceAtParameter(i);

                                            try
                                            {
                                                double z = surf1.FindElevationAtXY(x, y);

                                                dt1.Rows.Add();
                                                dt1.Rows[dt1.Rows.Count - 1][0] = sta;
                                                dt1.Rows[dt1.Rows.Count - 1][1] = x;
                                                dt1.Rows[dt1.Rows.Count - 1][2] = y;
                                                dt1.Rows[dt1.Rows.Count - 1][3] = z;
                                                dt1.Rows[dt1.Rows.Count - 1][4] = comboBox_surface.Text;
                                                dt1.Rows[dt1.Rows.Count - 1][5] = Environment.UserName.ToUpper();
                                            }
                                            catch (System.Exception)
                                            {


                                            }
                                        }

                                        double spacing = 0.1;

                                        if (Functions.IsNumeric(textBox_interval.Text) == true)
                                        {
                                            spacing = Convert.ToDouble(textBox_interval.Text);
                                        }

                                        int no_of_spacing = Convert.ToInt32(Math.Floor(poly1.Length / spacing));


                                        for (int i = 1; i <= no_of_spacing; i++)
                                        {
                                            Point3d pt1 = poly1.GetPointAtDist(i * spacing);

                                            double x = pt1.X;
                                            double y = pt1.Y;



                                            try
                                            {
                                                double z = surf1.FindElevationAtXY(x, y);
                                                dt1.Rows.Add();

                                                dt1.Rows[dt1.Rows.Count - 1][0] = i * spacing;
                                                dt1.Rows[dt1.Rows.Count - 1][1] = x;
                                                dt1.Rows[dt1.Rows.Count - 1][2] = y;
                                                dt1.Rows[dt1.Rows.Count - 1][3] = z;
                                                dt1.Rows[dt1.Rows.Count - 1][4] = comboBox_surface.Text;
                                                dt1.Rows[dt1.Rows.Count - 1][5] = Environment.UserName.ToUpper();
                                            }
                                            catch (System.Exception)
                                            {


                                            }
                                        }

                                        dt1 = Functions.Sort_data_table(dt1, "Station");

                                        dt1.Rows[0][6] = 0;

                                        double sta1 = 0;
                                        for (int i = 1; i < dt1.Rows.Count; i++)
                                        {
                                            if (dt1.Rows[i - 1][1] != DBNull.Value && dt1.Rows[i][1] != DBNull.Value &&
                                                dt1.Rows[i - 1][2] != DBNull.Value && dt1.Rows[i][2] != DBNull.Value &&
                                                dt1.Rows[i - 1][3] != DBNull.Value && dt1.Rows[i][3] != DBNull.Value &&
                                                dt1.Rows[i - 1][6] != DBNull.Value)
                                            {
                                                double x1 = Convert.ToDouble(dt1.Rows[i - 1][1]);
                                                double y1 = Convert.ToDouble(dt1.Rows[i - 1][2]);
                                                double z1 = Convert.ToDouble(dt1.Rows[i - 1][3]);
                                                sta1 = Convert.ToDouble(dt1.Rows[i - 1][6]);

                                                double x2 = Convert.ToDouble(dt1.Rows[i][1]);
                                                double y2 = Convert.ToDouble(dt1.Rows[i][2]);
                                                double z2 = Convert.ToDouble(dt1.Rows[i][3]);

                                                dt1.Rows[i][6] = sta1 + Math.Pow(Math.Pow(x1 - x2, 2) + Math.Pow(y1 - y2, 2) + Math.Pow(z1 - z2, 2), 0.5);
                                            }


                                        }

                                        string name1 = System.DateTime.Now.Year + "-" + System.DateTime.Now.Month + "-" + System.DateTime.Now.Day + "_" + System.DateTime.Now.Hour + "h" + System.DateTime.Now.Minute + "m";

                                        Functions.Transfer_datatable_to_new_excel_spreadsheet(dt1, name1);
                                        dt1 = null;

                                    }



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
            this.MdiParent.WindowState = FormWindowState.Normal;

            set_enable_true();
        }

        private void button_calcZ_Click(object sender, EventArgs e)
        {
            if (comboBox_surface.Text == "")
            {
                MessageBox.Show("no surface selected");
                return;
            }

            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.Civil.ApplicationServices.CivilDocument CivilDrawing = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument;
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
                        ObjectIdCollection col_surf = CivilDrawing.GetSurfaceIds();
                        for (int j = 0; j < col_surf.Count; ++j)
                        {
                            Autodesk.Civil.DatabaseServices.Surface surf1 = Trans1.GetObject(col_surf[j], OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Surface;
                            if (surf1 != null)
                            {
                                if (surf1.Name == comboBox_surface.Text)
                                {

                                    System.Data.DataTable dt1 = new System.Data.DataTable();
                                    dt1.Columns.Add("X", typeof(double));
                                    dt1.Columns.Add("Y", typeof(double));
                                    dt1.Columns.Add("Z picked", typeof(double));
                                    dt1.Columns.Add("Z on surface", typeof(double));
                                    dt1.Columns.Add("cover", typeof(double));
                                    dt1.Columns.Add("Surface", typeof(string));
                                    dt1.Columns.Add("USER", typeof(string));

                                    bool run1 = true;

                                    do
                                    {
                                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                        PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify point");
                                        PP1.AllowNone = false;
                                        this.MdiParent.WindowState = FormWindowState.Minimized;

                                        Point_res1 = Editor1.GetPoint(PP1);

                                        if (Point_res1.Status != PromptStatus.OK)
                                        {
                                            run1 = false;
                                        }
                                        else
                                        {
                                            Point3d ptins = Point_res1.Value;

                                            double x = Math.Round(ptins.X, 3);
                                            double y = Math.Round(ptins.Y, 3);
                                            double z = Math.Round(ptins.Z, 3);

                                            double elev1 = surf1.FindElevationAtXY(ptins.X, ptins.Y);
                                            ThisDrawing.Editor.WriteMessage("\nX=" + Convert.ToString(x) + ", Y=" + Convert.ToString(y) +
                                                                             ", Z=" + Convert.ToString(z) + "(" + Convert.ToString(Math.Round(elev1, 3)) + " on surface)");
                                            dt1.Rows.Add();
                                            dt1.Rows[dt1.Rows.Count - 1][0] = x;
                                            dt1.Rows[dt1.Rows.Count - 1][1] = y;
                                            dt1.Rows[dt1.Rows.Count - 1][2] = z;
                                            dt1.Rows[dt1.Rows.Count - 1][3] = Math.Round(elev1, 3);
                                            dt1.Rows[dt1.Rows.Count - 1][4] = Math.Round(elev1, 3) - z;
                                            dt1.Rows[dt1.Rows.Count - 1][5] = comboBox_surface.Text;
                                            dt1.Rows[dt1.Rows.Count - 1][6] = Environment.UserName.ToUpper();


                                        }
                                    } while (run1 == true);

                                    if (checkBox_excel.Checked == true)
                                    {
                                        string name1 = System.DateTime.Now.Year + "-" + System.DateTime.Now.Month + "-" + System.DateTime.Now.Day + "_" + System.DateTime.Now.Hour + "h" + System.DateTime.Now.Minute + "m";

                                        Functions.Transfer_datatable_to_new_excel_spreadsheet(dt1, name1);
                                    }
                                    dt1 = null;
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
            this.MdiParent.WindowState = FormWindowState.Normal;

            set_enable_true();
        }

        private void button_load_agen_project_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog fbd = new OpenFileDialog())
            {
                fbd.Multiselect = false;
                fbd.Filter = "Excel files (*.xlsx)|*.xlsx";

                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {

                    string File1 = fbd.FileName;



                    #region Load_config_method
                    {
                        set_enable_false();
                        Load_existing_config_file(File1);
                        set_enable_true();
                    }
                    #endregion


                }
            }
        }


        private void Load_existing_config_file(string File1)
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

                if (Excel1.Workbooks.Count == 0) Excel1.Visible = false;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(File1);

                string ProjFolder = "";
                string cl_excel_name = "centerline.xlsx";

                try
                {
                    int no_worksheets = Workbook1.Worksheets.Count;
                    foreach (Microsoft.Office.Interop.Excel.Worksheet W1 in Workbook1.Worksheets)
                    {
                        try
                        {
                            #region main_worksheet
                            if (W1.Name == "main_cfg")
                            {
                                string b1 = Convert.ToString(W1.Range["B1"].Value2);
                                string b2 = Convert.ToString(W1.Range["B2"].Value);
                                string b3 = Convert.ToString(W1.Range["B3"].Value);
                                string b4 = Convert.ToString(W1.Range["B4"].Value);
                                string b5 = Convert.ToString(W1.Range["B5"].Value);
                                string b6 = Convert.ToString(W1.Range["B6"].Value);
                                string b7 = Convert.ToString(W1.Range["B7"].Value);
                                string b8 = Convert.ToString(W1.Range["B8"].Value);
                                string b9 = Convert.ToString(W1.Range["B9"].Value);
                                string b10 = Convert.ToString(W1.Range["B10"].Value);
                                string b11 = Convert.ToString(W1.Range["B11"].Value);
                                string b12 = Convert.ToString(W1.Range["B12"].Value);
                                string b13 = Convert.ToString(W1.Range["B13"].Value);
                                string b14 = Convert.ToString(W1.Range["B14"].Value);
                                string b15 = Convert.ToString(W1.Range["B15"].Value);
                                string b16 = Convert.ToString(W1.Range["B16"].Value);
                                string b17 = Convert.ToString(W1.Range["B17"].Value);
                                string b18 = Convert.ToString(W1.Range["B18"].Value);
                                string b19 = Convert.ToString(W1.Range["B19"].Value);
                                string b20 = Convert.ToString(W1.Range["B20"].Value);
                                string b21 = Convert.ToString(W1.Range["B21"].Value);
                                string b22 = Convert.ToString(W1.Range["B22"].Value);
                                string b23 = Convert.ToString(W1.Range["B23"].Value);
                                string b24 = Convert.ToString(W1.Range["B24"].Value);
                                string b25 = Convert.ToString(W1.Range["B25"].Value);
                                string b26 = Convert.ToString(W1.Range["B26"].Value);
                                string b27 = Convert.ToString(W1.Range["B27"].Value);
                                string b28 = Convert.ToString(W1.Range["B28"].Value);
                                string b29 = Convert.ToString(W1.Range["B29"].Value);
                                string b30 = Convert.ToString(W1.Range["B30"].Value);
                                string b31 = Convert.ToString(W1.Range["B31"].Value);
                                string b32 = Convert.ToString(W1.Range["B32"].Value);
                                string b33 = Convert.ToString(W1.Range["B33"].Value);
                                string b34 = Convert.ToString(W1.Range["B34"].Value);
                                string b35 = Convert.ToString(W1.Range["B35"].Value);
                                string b36 = Convert.ToString(W1.Range["B36"].Value);
                                string b37 = Convert.ToString(W1.Range["B37"].Value);
                                string b38 = Convert.ToString(W1.Range["B38"].Value);
                                string b39 = Convert.ToString(W1.Range["B39"].Value);
                                string b40 = Convert.ToString(W1.Range["B40"].Value);
                                string b41 = Convert.ToString(W1.Range["B41"].Value);
                                string b42 = Convert.ToString(W1.Range["B42"].Value);
                                string b43 = Convert.ToString(W1.Range["B43"].Value);
                                string b44 = Convert.ToString(W1.Range["B44"].Value);
                                string b45 = Convert.ToString(W1.Range["B45"].Value);
                                string b46 = Convert.ToString(W1.Range["B46"].Value);

                                string b47 = Convert.ToString(W1.Range["B47"].Value);
                                if (b47 != null && b47.Replace(" ", "") != "")
                                {
                                    segment_current = b47;
                                }

                                string b48 = Convert.ToString(W1.Range["B48"].Value);
                                string b49 = Convert.ToString(W1.Range["B49"].Value);
                                string b50 = Convert.ToString(W1.Range["B50"].Value);
                                string b51 = Convert.ToString(W1.Range["B51"].Value);
                                string b52 = Convert.ToString(W1.Range["B52"].Value);
                                string b53 = Convert.ToString(W1.Range["B53"].Value);
                                string b54 = Convert.ToString(W1.Range["B54"].Value);
                                string b55 = Convert.ToString(W1.Range["B55"].Value);
                                string b56 = Convert.ToString(W1.Range["B56"].Value);
                                string b57 = Convert.ToString(W1.Range["B57"].Value);
                                string b58 = Convert.ToString(W1.Range["B58"].Value);
                                string b59 = Convert.ToString(W1.Range["B59"].Value);
                                string b60 = Convert.ToString(W1.Range["B60"].Value);
                                string b61 = Convert.ToString(W1.Range["B61"].Value);
                                string b62 = Convert.ToString(W1.Range["B62"].Value);
                                string b63 = Convert.ToString(W1.Range["B63"].Value);
                                string b64 = Convert.ToString(W1.Range["B64"].Value);
                                string b65 = Convert.ToString(W1.Range["B65"].Value);
                                string b66 = Convert.ToString(W1.Range["B66"].Value);
                                string b67 = Convert.ToString(W1.Range["B67"].Value);
                                string b68 = Convert.ToString(W1.Range["B68"].Value);
                                string b69 = Convert.ToString(W1.Range["B69"].Value);
                                string b70 = Convert.ToString(W1.Range["B70"].Value);
                                string b71 = Convert.ToString(W1.Range["B71"].Value);
                                string b72 = Convert.ToString(W1.Range["B72"].Value);
                                string b73 = Convert.ToString(W1.Range["B73"].Value);
                                string b74 = Convert.ToString(W1.Range["B74"].Value);
                                string b75 = Convert.ToString(W1.Range["B75"].Value);
                                string b76 = Convert.ToString(W1.Range["B76"].Value);
                                string b77 = Convert.ToString(W1.Range["B77"].Value);
                                string b78 = Convert.ToString(W1.Range["B78"].Value);
                                string b79 = Convert.ToString(W1.Range["B79"].Value);
                                string b80 = Convert.ToString(W1.Range["B80"].Value);
                                string b81 = Convert.ToString(W1.Range["B81"].Value);
                                string b82 = Convert.ToString(W1.Range["B82"].Value);
                                string b83 = Convert.ToString(W1.Range["B83"].Value);
                                string b84 = Convert.ToString(W1.Range["B84"].Value);
                                string b85 = Convert.ToString(W1.Range["B85"].Value);
                                string b86 = Convert.ToString(W1.Range["B86"].Value);
                                string b87 = Convert.ToString(W1.Range["B87"].Value);
                                string b88 = Convert.ToString(W1.Range["B88"].Value);
                                string b89 = Convert.ToString(W1.Range["B89"].Value);
                                string b90 = Convert.ToString(W1.Range["B90"].Value);
                                string b91 = Convert.ToString(W1.Range["B91"].Value);
                                string b92 = Convert.ToString(W1.Range["B92"].Value);
                                string b93 = Convert.ToString(W1.Range["B93"].Value);
                                string b94 = Convert.ToString(W1.Range["B94"].Value);
                                string b95 = Convert.ToString(W1.Range["B95"].Value);
                                string b96 = Convert.ToString(W1.Range["B96"].Value);
                                string b97 = Convert.ToString(W1.Range["B97"].Value);
                                string b98 = Convert.ToString(W1.Range["B98"].Value);
                                string b99 = Convert.ToString(W1.Range["B99"].Value);

                                if (b16 != null)
                                {
                                    ProjFolder = b16.ToString();
                                    if (System.IO.Directory.Exists(ProjFolder) == true)
                                    {
                                        if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                                        {
                                            ProjFolder = ProjFolder + "\\";
                                        }
                                        ProjFolder_main = ProjFolder;

                                        ProjFolder = ProjFolder + segment_current;

                                        if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                                        {
                                            ProjFolder = ProjFolder + "\\";
                                        }

                                        Microsoft.Office.Interop.Excel.Workbook Workbook2 = null;
                                        Microsoft.Office.Interop.Excel.Worksheet W2 = null;
                                        try
                                        {
                                            string fisier_cl = ProjFolder + cl_excel_name;
                                            if (System.IO.File.Exists(fisier_cl) == true)
                                            {
                                                List<string> lista_segm = new List<string>();

                                                if (b47 != "")
                                                {
                                                    lista_segm.Add(b47);

                                                    if (b48 != "")
                                                    {
                                                        lista_segm.Add(b48);

                                                        if (b49 != "")
                                                        {
                                                            lista_segm.Add(b49);

                                                            if (b50 != "")
                                                            {
                                                                lista_segm.Add(b50);

                                                                if (b51 != "")
                                                                {
                                                                    lista_segm.Add(b51);

                                                                    if (b52 != "")
                                                                    {
                                                                        lista_segm.Add(b52);

                                                                        if (b53 != "")
                                                                        {
                                                                            lista_segm.Add(b53);
                                                                            if (b54 != "")
                                                                            {
                                                                                lista_segm.Add(b54);

                                                                                if (b55 != "")
                                                                                {
                                                                                    lista_segm.Add(b55);

                                                                                    if (b56 != "")
                                                                                    {
                                                                                        lista_segm.Add(b56);

                                                                                        if (b57 != "")
                                                                                        {
                                                                                            lista_segm.Add(b57);

                                                                                            if (b58 != "")
                                                                                            {
                                                                                                lista_segm.Add(b58);

                                                                                                if (b59 != "")
                                                                                                {
                                                                                                    lista_segm.Add(b59);

                                                                                                    if (b60 != "")
                                                                                                    {
                                                                                                        lista_segm.Add(b61);

                                                                                                        if (b62 != "")
                                                                                                        {
                                                                                                            lista_segm.Add(b62);

                                                                                                            if (b63 != "")
                                                                                                            {
                                                                                                                lista_segm.Add(b63);

                                                                                                                if (b64 != "")
                                                                                                                {
                                                                                                                    lista_segm.Add(b64);

                                                                                                                    if (b65 != "")
                                                                                                                    {
                                                                                                                        lista_segm.Add(b65);

                                                                                                                        if (b66 != "")
                                                                                                                        {
                                                                                                                            lista_segm.Add(b66);

                                                                                                                            if (b67 != "")
                                                                                                                            {
                                                                                                                                lista_segm.Add(b67);

                                                                                                                                if (b68 != "")
                                                                                                                                {
                                                                                                                                    lista_segm.Add(b68);

                                                                                                                                    if (b69 != "")
                                                                                                                                    {
                                                                                                                                        lista_segm.Add(b69);
                                                                                                                                        if (b70 != "")
                                                                                                                                        {
                                                                                                                                            lista_segm.Add(b70);

                                                                                                                                            if (b71 != "")
                                                                                                                                            {
                                                                                                                                                lista_segm.Add(b71);

                                                                                                                                                if (b72 != "")
                                                                                                                                                {
                                                                                                                                                    lista_segm.Add(b72);

                                                                                                                                                    if (b73 != "")
                                                                                                                                                    {
                                                                                                                                                        lista_segm.Add(b73);

                                                                                                                                                        if (b74 != "")
                                                                                                                                                        {
                                                                                                                                                            lista_segm.Add(b74);

                                                                                                                                                            if (b75 != "")
                                                                                                                                                            {
                                                                                                                                                                lista_segm.Add(b75);

                                                                                                                                                                if (b76 != "")
                                                                                                                                                                {
                                                                                                                                                                    lista_segm.Add(b76);
                                                                                                                                                                    if (b77 != "")
                                                                                                                                                                    {
                                                                                                                                                                        lista_segm.Add(b77);

                                                                                                                                                                        if (b78 != "")
                                                                                                                                                                        {
                                                                                                                                                                            lista_segm.Add(b78);

                                                                                                                                                                            if (b79 != "")
                                                                                                                                                                            {
                                                                                                                                                                                lista_segm.Add(b79);

                                                                                                                                                                                if (b80 != "")
                                                                                                                                                                                {
                                                                                                                                                                                    lista_segm.Add(b80);
                                                                                                                                                                                    if (b81 != "")
                                                                                                                                                                                    {
                                                                                                                                                                                        lista_segm.Add(b81);
                                                                                                                                                                                        if (b82 != "")
                                                                                                                                                                                        {
                                                                                                                                                                                            lista_segm.Add(b82);

                                                                                                                                                                                            if (b83 != "")
                                                                                                                                                                                            {
                                                                                                                                                                                                lista_segm.Add(b83);

                                                                                                                                                                                                if (b84 != "")
                                                                                                                                                                                                {
                                                                                                                                                                                                    lista_segm.Add(b84);

                                                                                                                                                                                                    if (b85 != "")
                                                                                                                                                                                                    {
                                                                                                                                                                                                        lista_segm.Add(b85);

                                                                                                                                                                                                        if (b86 != "")
                                                                                                                                                                                                        {
                                                                                                                                                                                                            lista_segm.Add(b86);

                                                                                                                                                                                                            if (b87 != "")
                                                                                                                                                                                                            {
                                                                                                                                                                                                                lista_segm.Add(b87);

                                                                                                                                                                                                                if (b88 != "")
                                                                                                                                                                                                                {
                                                                                                                                                                                                                    lista_segm.Add(b88);

                                                                                                                                                                                                                    if (b89 != "")
                                                                                                                                                                                                                    {
                                                                                                                                                                                                                        lista_segm.Add(b89);
                                                                                                                                                                                                                        if (b90 != "")
                                                                                                                                                                                                                        {
                                                                                                                                                                                                                            lista_segm.Add(b90);

                                                                                                                                                                                                                            if (b91 != "")
                                                                                                                                                                                                                            {
                                                                                                                                                                                                                                lista_segm.Add(b91);

                                                                                                                                                                                                                                if (b92 != "")
                                                                                                                                                                                                                                {
                                                                                                                                                                                                                                    lista_segm.Add(b92);

                                                                                                                                                                                                                                    if (b93 != "")
                                                                                                                                                                                                                                    {
                                                                                                                                                                                                                                        lista_segm.Add(b93);

                                                                                                                                                                                                                                        if (b94 != "")
                                                                                                                                                                                                                                        {
                                                                                                                                                                                                                                            lista_segm.Add(b94);

                                                                                                                                                                                                                                            if (b95 != "")
                                                                                                                                                                                                                                            {
                                                                                                                                                                                                                                                lista_segm.Add(b95);

                                                                                                                                                                                                                                                if (b96 != "")
                                                                                                                                                                                                                                                {
                                                                                                                                                                                                                                                    lista_segm.Add(b96);
                                                                                                                                                                                                                                                    if (b97 != "")
                                                                                                                                                                                                                                                    {
                                                                                                                                                                                                                                                        lista_segm.Add(b97);
                                                                                                                                                                                                                                                        if (b98 != "")
                                                                                                                                                                                                                                                        {
                                                                                                                                                                                                                                                            lista_segm.Add(b98);
                                                                                                                                                                                                                                                            if (b98 != "")
                                                                                                                                                                                                                                                            {
                                                                                                                                                                                                                                                                lista_segm.Add(b98);
                                                                                                                                                                                                                                                                if (b99 != "")
                                                                                                                                                                                                                                                                {
                                                                                                                                                                                                                                                                    lista_segm.Add(b99);
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

                                                        }

                                                    }

                                                }

                                                comboBox_segments.DataSource = lista_segm;

                                                dt_centerline = Creaza_centerline_datatable_structure();
                                                Workbook2 = Excel1.Workbooks.Open(fisier_cl);
                                                W2 = Workbook2.Worksheets[1];

                                                Microsoft.Office.Interop.Excel.Range range1 = W2.Range["C1:J300000"];

                                                object[,] values = new object[300000, 8];

                                                values = range1.Value2;

                                                for (int i = 10; i <= 300000; ++i)
                                                {
                                                    object valX = values[i, 1];
                                                    object valY = values[i, 2];
                                                    object valZ = values[i, 3];
                                                    object valSta = values[i, 5];
                                                    object valBackSta = values[i, 7];
                                                    object valAheadSta = values[i, 8];
                                                    if (valX != null && valY != null && valZ != null && valSta != null
                                                        && Functions.IsNumeric(Convert.ToString(valX)) == true && Functions.IsNumeric(Convert.ToString(valY)) == true
                                                        && Functions.IsNumeric(Convert.ToString(valZ)) == true && Functions.IsNumeric(Convert.ToString(valSta)) == true)
                                                    {
                                                        dt_centerline.Rows.Add();
                                                        dt_centerline.Rows[dt_centerline.Rows.Count - 1][Col_3DSta] = Convert.ToDouble(valSta);
                                                        dt_centerline.Rows[dt_centerline.Rows.Count - 1][Col_x] = Convert.ToDouble(valX);
                                                        dt_centerline.Rows[dt_centerline.Rows.Count - 1][Col_y] = Convert.ToDouble(valY);
                                                        dt_centerline.Rows[dt_centerline.Rows.Count - 1][Col_z] = Convert.ToDouble(valZ);
                                                        if (valBackSta != null && valAheadSta != null)
                                                        {
                                                            dt_centerline.Rows[dt_centerline.Rows.Count - 1][Col_BackSta] = Convert.ToDouble(valBackSta);
                                                            dt_centerline.Rows[dt_centerline.Rows.Count - 1][Col_AheadSta] = Convert.ToDouble(valAheadSta);
                                                        }
                                                    }
                                                    else
                                                    {
                                                        i = 300001;
                                                    }
                                                }
                                                Workbook2.Close();
                                                poly_centerline = Functions.Build_2d_poly_from_datatable(dt_centerline);
                                            }
                                            else
                                            {
                                                dt_centerline = null;
                                                poly_centerline = null;
                                            }
                                        }
                                        catch (System.Exception ex)
                                        {
                                            System.Windows.Forms.MessageBox.Show(ex.Message);
                                        }
                                        finally
                                        {
                                            if (W2 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W2);
                                            if (Workbook2 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook2);
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show("the Project database folder location is not specified\r\n" + ProjFolder + "\r\n operation aborted");
                                        return;
                                    }
                                }
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
                    if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);

                }


                if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);

            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }


            //Functions.Transfer_datatable_to_new_excel_spreadsheet(Data_table_centerline);

        }


        public System.Data.DataTable Creaza_centerline_datatable_structure()
        {

            string Col_MMid = "MMID";
            string Col_Type = "Type";

            string Col_2DSta = "2DSta";

            string Col_EqSta = "EqSta";

            string Col_DeflAng = "DeflAng";
            string Col_DeflAngDMS = "DeflAngDMS";
            string Col_Bearing = "Bearing";
            string Col_Distance = "Distance";
            string Col_DisplaySta = "DisplaySta";
            string Col_DisplayPI = "DisplayPI";
            string Col_DisplayProf = "DisplayProf";
            string Col_Symbol = "Symbol";

            System.Type type_MMid = typeof(string);
            System.Type type_Type = typeof(string);
            System.Type type_x = typeof(double);
            System.Type type_y = typeof(double);
            System.Type type_z = typeof(double);
            System.Type type_2DSta = typeof(double);
            System.Type type_3DSta = typeof(double);
            System.Type type_EqSta = typeof(double);
            System.Type type_BackSta = typeof(double);
            System.Type type_AheadSta = typeof(double);
            System.Type type_DeflAng = typeof(double);
            System.Type type_DeflAngDMS = typeof(string);
            System.Type type_Bearing = typeof(string);
            System.Type type_Distance = typeof(double);
            System.Type type_DisplaySta = typeof(double);
            System.Type type_DisplayPI = typeof(int);
            System.Type type_DisplayProf = typeof(int);
            System.Type type_Symbol = typeof(string);


            List<string> Lista1 = new List<string>();
            List<System.Type> Lista2 = new List<System.Type>();

            Lista1.Add(Col_MMid);
            Lista1.Add(Col_Type);
            Lista1.Add(Col_x);
            Lista1.Add(Col_y);
            Lista1.Add(Col_z);
            Lista1.Add(Col_2DSta);
            Lista1.Add(Col_3DSta);
            Lista1.Add(Col_EqSta);
            Lista1.Add(Col_BackSta);
            Lista1.Add(Col_AheadSta);
            Lista1.Add(Col_DeflAng);
            Lista1.Add(Col_DeflAngDMS);
            Lista1.Add(Col_Bearing);
            Lista1.Add(Col_Distance);
            Lista1.Add(Col_DisplaySta);
            Lista1.Add(Col_DisplayPI);
            Lista1.Add(Col_DisplayProf);
            Lista1.Add(Col_Symbol);

            Lista2.Add(type_MMid);
            Lista2.Add(type_Type);
            Lista2.Add(type_x);
            Lista2.Add(type_y);
            Lista2.Add(type_z);
            Lista2.Add(type_2DSta);
            Lista2.Add(type_3DSta);
            Lista2.Add(type_EqSta);
            Lista2.Add(type_BackSta);
            Lista2.Add(type_AheadSta);
            Lista2.Add(type_DeflAng);
            Lista2.Add(type_DeflAngDMS);
            Lista2.Add(type_Bearing);
            Lista2.Add(type_Distance);
            Lista2.Add(type_DisplaySta);
            Lista2.Add(type_DisplayPI);
            Lista2.Add(type_DisplayProf);
            Lista2.Add(type_Symbol);


            System.Data.DataTable Data_table_centerline = new System.Data.DataTable();

            for (int i = 0; i < Lista1.Count; ++i)
            {
                Data_table_centerline.Columns.Add(Lista1[i], Lista2[i]);
            }
            return Data_table_centerline;
        }

        private void comboBox_segments_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (segment_current != "")
            {
                if (comboBox_segments.Text != segment_current)
                {
                    segment_current = comboBox_segments.Text;
                    Microsoft.Office.Interop.Excel.Application Excel1 = null;
                    Microsoft.Office.Interop.Excel.Worksheet W1 = null;
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

                        if (Excel1.Workbooks.Count == 0) Excel1.Visible = false;

                        string ProjFolder = ProjFolder_main + segment_current;

                        if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                        {
                            ProjFolder = ProjFolder + "\\";
                        }


                        string cl_excel_name = "centerline.xlsx";
                        string fisier_cl = ProjFolder + cl_excel_name;


                        if (System.IO.File.Exists(fisier_cl) == true)
                        {
                            set_enable_false();
                            dt_centerline = Creaza_centerline_datatable_structure();
                            Workbook1 = Excel1.Workbooks.Open(fisier_cl);
                            W1 = Workbook1.Worksheets[1];

                            Microsoft.Office.Interop.Excel.Range range1 = W1.Range["C1:J300000"];

                            object[,] values = new object[300000, 8];

                            values = range1.Value2;

                            for (int i = 10; i <= 300000; ++i)
                            {
                                object valX = values[i, 1];
                                object valY = values[i, 2];
                                object valZ = values[i, 3];
                                object valSta = values[i, 5];
                                object valBackSta = values[i, 7];
                                object valAheadSta = values[i, 8];

                                if (valX != null && valY != null && valZ != null && valSta != null
                                    && Functions.IsNumeric(Convert.ToString(valX)) == true && Functions.IsNumeric(Convert.ToString(valY)) == true
                                    && Functions.IsNumeric(Convert.ToString(valZ)) == true && Functions.IsNumeric(Convert.ToString(valSta)) == true)
                                {
                                    dt_centerline.Rows.Add();
                                    dt_centerline.Rows[dt_centerline.Rows.Count - 1][Col_3DSta] = Convert.ToDouble(valSta);
                                    dt_centerline.Rows[dt_centerline.Rows.Count - 1][Col_x] = Convert.ToDouble(valX);
                                    dt_centerline.Rows[dt_centerline.Rows.Count - 1][Col_y] = Convert.ToDouble(valY);
                                    dt_centerline.Rows[dt_centerline.Rows.Count - 1][Col_z] = Convert.ToDouble(valZ);
                                    if (valBackSta != null && valAheadSta != null)
                                    {
                                        dt_centerline.Rows[dt_centerline.Rows.Count - 1][Col_BackSta] = Convert.ToDouble(valBackSta);
                                        dt_centerline.Rows[dt_centerline.Rows.Count - 1][Col_AheadSta] = Convert.ToDouble(valAheadSta);
                                    }
                                }
                                else
                                {
                                    i = 300001;
                                }
                            }

                            Workbook1.Close();
                            poly_centerline = Functions.Build_2d_poly_from_datatable(dt_centerline);
                        }
                        else
                        {
                            dt_centerline = null;
                            poly_centerline = null;
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

                    // Functions.Transfer_datatable_to_new_excel_spreadsheet(Data_table_centerline);

                }
            }
            set_enable_true();

        }

        private void button_read_chainage_from_profile_view_Click(object sender, EventArgs e)
        {
            if (dt_centerline == null || dt_centerline.Rows.Count <= 1 || poly_centerline == null)
            {
                MessageBox.Show("No Agen project is loaded\r\nOperation aborted.");
                return;
            }

            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.Civil.ApplicationServices.CivilDocument CivilDrawing = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument;

            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {
                set_enable_false();
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    List<ObjectId> lista_objid = new List<ObjectId>();
                    List<string> lista_chain = new List<string>();


                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                        this.MdiParent.WindowState = FormWindowState.Minimized;

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_profile;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions prompt_prof;
                        prompt_prof = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect profile (GROUND):");
                        prompt_prof.SetRejectMessage("\nSelect a profile!");
                        prompt_prof.AllowNone = true;
                        prompt_prof.AddAllowedClass(typeof(Autodesk.Civil.DatabaseServices.Profile), false);
                        Rezultat_profile = ThisDrawing.Editor.GetEntity(prompt_prof);

                        if (Rezultat_profile.Status != PromptStatus.OK)
                        {
                            Editor1.SetImpliedSelection(Empty_array);
                            Editor1.WriteMessage("\nCommand:");
                            set_enable_true();
                            this.MdiParent.WindowState = FormWindowState.Normal;
                            return;
                        }

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_profileview;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions prompt_profview;
                        prompt_profview = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect Profile View (GRID):");
                        prompt_profview.SetRejectMessage("\nSelect a profile View!");
                        prompt_profview.AllowNone = true;
                        prompt_profview.AddAllowedClass(typeof(Autodesk.Civil.DatabaseServices.ProfileView), false);
                        Rezultat_profileview = ThisDrawing.Editor.GetEntity(prompt_profview);

                        if (Rezultat_profileview.Status != PromptStatus.OK)
                        {
                            Editor1.SetImpliedSelection(Empty_array);
                            Editor1.WriteMessage("\nCommand:");
                            set_enable_true();
                            this.MdiParent.WindowState = FormWindowState.Normal;
                            return;
                        }

                        Autodesk.Civil.DatabaseServices.Profile prof1 = Trans1.GetObject(Rezultat_profile.ObjectId, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Profile;
                        Autodesk.Civil.DatabaseServices.ProfileView profview1 = Trans1.GetObject(Rezultat_profileview.ObjectId, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.ProfileView;
                        if (prof1 != null && profview1 != null)
                        {
                            ObjectId align_id = prof1.AlignmentId;
                            Autodesk.Civil.DatabaseServices.Alignment align1 = Trans1.GetObject(align_id, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Alignment;
                            if (align1 != null)
                            {
                                Functions.Creaza_layer("NO PLOT", 40, false);
                                Create_mleader_object_data_table();
                                bool run1 = true;
                                do
                                {
                                    Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                    Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                    PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify point:");
                                    PP1.AllowNone = false;
                                    Point_res1 = Editor1.GetPoint(PP1);

                                    if (Point_res1.Status != PromptStatus.OK)
                                    {
                                        run1 = false;
                                    }

                                    if (run1 == true)
                                    {
                                        Point3d pickpt = Point_res1.Value;
                                        double station_prof = -1.234;
                                        double elev1 = 0;
                                        profview1.FindStationAndElevationAtXY(pickpt.X, pickpt.Y, ref station_prof, ref elev1);
                                        double x = -1.234;
                                        double y = -1.234;
                                        align1.PointLocation(station_prof, 0, ref x, ref y);

                                        Point3d pt_on_poly = poly_centerline.GetClosestPointTo(new Point3d(x, y, 0), Vector3d.ZAxis, false);
                                        int idx1 = Convert.ToInt32(Math.Floor(poly_centerline.GetParameterAtPoint(pt_on_poly)));
                                        if (idx1 < dt_centerline.Rows.Count - 1)
                                        {
                                            double x1 = Convert.ToDouble(dt_centerline.Rows[idx1][Col_x]);
                                            double y1 = Convert.ToDouble(dt_centerline.Rows[idx1][Col_y]);

                                            double x2 = Convert.ToDouble(dt_centerline.Rows[idx1 + 1][Col_x]);
                                            double y2 = Convert.ToDouble(dt_centerline.Rows[idx1 + 1][Col_y]);

                                            double sta1 = Convert.ToDouble(dt_centerline.Rows[idx1][Col_3DSta]);
                                            double sta2 = Convert.ToDouble(dt_centerline.Rows[idx1 + 1][Col_3DSta]);

                                            if (dt_centerline.Rows[idx1][Col_AheadSta] != DBNull.Value)
                                            {
                                                sta1 = Convert.ToDouble(dt_centerline.Rows[idx1][Col_AheadSta]);
                                            }

                                            if (dt_centerline.Rows[idx1 + 1][Col_BackSta] != DBNull.Value)
                                            {
                                                sta2 = Convert.ToDouble(dt_centerline.Rows[idx1 + 1][Col_BackSta]);
                                            }

                                            double d1 = Math.Pow(Math.Pow(x - x1, 2) + Math.Pow(y - y1, 2), 0.5);
                                            double d = Math.Pow(Math.Pow(x2 - x1, 2) + Math.Pow(y2 - y1, 2), 0.5);

                                            double stax = sta1 + ((sta2 - sta1) * d1) / d;

                                            string continut = Functions.Get_chainage_from_double(stax, "m", Convert.ToInt32(comboBox_round_sta.Text));
                                            if (checkBox_sta.Checked == true)
                                            {
                                                continut = continut + "\r\nSTA " + Functions.Get_chainage_from_double(station_prof, "m", Convert.ToInt32(comboBox_round_sta.Text));
                                            }
                                            if (checkBox_elev.Checked == true)
                                            {
                                                continut = continut + "\r\nEL. " + Functions.Get_String_Rounded(elev1, Convert.ToInt32(comboBox_round_elev.Text));
                                            }

                                            MLeader ml1 = Functions.creaza_mleader(pickpt, continut, 0.5, 1.5, -3.5, 0.5, 0.5, 0.5, "NO PLOT");
                                            MLeader ml2 = Functions.creaza_mleader(new Point3d(x, y, 0), continut, 0.5, 1.5, 3.5, 0.5, 0.5, 0.5, "NO PLOT");
                                            Trans1.TransactionManager.QueueForGraphicsFlush();
                                            lista_objid.Add(ml1.ObjectId);
                                            lista_chain.Add(Functions.Get_chainage_from_double(stax, "m", Convert.ToInt32(comboBox_round_sta.Text)));
                                            lista_objid.Add(ml2.ObjectId);
                                            lista_chain.Add(Functions.Get_chainage_from_double(stax, "m", Convert.ToInt32(comboBox_round_sta.Text)));
                                        }

                                    }
                                } while (run1 == true);
                            }
                        }

                        if (lista_chain.Count > 0)
                        {
                            Append_object_data_to_ODXXX(lista_objid, segment_current, lista_chain);
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
            this.MdiParent.WindowState = FormWindowState.Normal;
        }


        private void Create_mleader_object_data_table()
        {

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

                            List<string> List1 = new List<string>();
                            List<string> List2 = new List<string>();
                            List<Autodesk.Gis.Map.Constants.DataType> List3 = new List<Autodesk.Gis.Map.Constants.DataType>();

                            List1.Add("MMID");
                            List2.Add("ObjectID of the mleader");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            List1.Add("Segment");
                            List2.Add("Centerline Version");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            List1.Add("Chainage");
                            List2.Add("Chainage");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            List1.Add("UserName");
                            List2.Add("Generated by");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            List1.Add("Date");
                            List2.Add("Date and Time");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            Functions.Get_object_data_table("ODXXX", "Generated by Profiler", List1, List2, List3);

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

        private void Append_object_data_to_ODXXX(List<ObjectId> lista1, string segment1, List<string> mleader_content)
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
                        Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                        for (int i = 0; i < lista1.Count; ++i)
                        {

                            List<object> Lista_val = new List<object>();
                            List<Autodesk.Gis.Map.Constants.DataType> Lista_type = new List<Autodesk.Gis.Map.Constants.DataType>();

                            ObjectId id1 = lista1[i];

                            Lista_val.Add(id1.Handle.Value.ToString());
                            Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                            Lista_val.Add(segment1);
                            Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                            Lista_val.Add(mleader_content[i]);
                            Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                            Lista_val.Add(Environment.UserName.ToUpper());
                            Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            Lista_val.Add(System.DateTime.Today.Year + "-" + System.DateTime.Today.Month + "-" + System.DateTime.Today.Day + " at " + System.DateTime.Now.Hour + ":" + System.DateTime.Now.Minute);
                            Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);


                            Functions.Populate_object_data_table_from_objectid(Tables1, id1, "ODXXX", Lista_val, Lista_type);
                        }

                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        private void button_place_points_on_profile_Click(object sender, EventArgs e)
        {
            if (dt_centerline == null || dt_centerline.Rows.Count <= 1 || poly_centerline == null)
            {
                MessageBox.Show("No Agen project is loaded\r\nOperation aborted.");
                return;
            }
            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.Civil.ApplicationServices.CivilDocument CivilDrawing = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument;

            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {
                set_enable_false();
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    List<ObjectId> lista_objid = new List<ObjectId>();
                    List<string> lista_chain = new List<string>();


                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;




                        this.MdiParent.WindowState = FormWindowState.Minimized;


                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_profileview;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions prompt_profview;
                        prompt_profview = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect Profile View (GRID):");
                        prompt_profview.SetRejectMessage("\nSelect a profile View!");
                        prompt_profview.AllowNone = true;
                        prompt_profview.AddAllowedClass(typeof(Autodesk.Civil.DatabaseServices.ProfileView), false);
                        Rezultat_profileview = ThisDrawing.Editor.GetEntity(prompt_profview);

                        if (Rezultat_profileview.Status != PromptStatus.OK)
                        {
                            Editor1.SetImpliedSelection(Empty_array);
                            Editor1.WriteMessage("\nCommand:");
                            set_enable_true();
                            this.MdiParent.WindowState = FormWindowState.Normal;
                            return;
                        }
                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_alignment;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions prompt_alg;
                        prompt_alg = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect Alignment:");
                        prompt_alg.SetRejectMessage("\nSelect an alignment!");
                        prompt_alg.AllowNone = true;
                        prompt_alg.AddAllowedClass(typeof(Autodesk.Civil.DatabaseServices.Alignment), false);
                        Rezultat_alignment = ThisDrawing.Editor.GetEntity(prompt_alg);

                        if (Rezultat_alignment.Status != PromptStatus.OK)
                        {
                            Editor1.SetImpliedSelection(Empty_array);
                            Editor1.WriteMessage("\nCommand:");
                            set_enable_true();
                            this.MdiParent.WindowState = FormWindowState.Normal;
                            return;
                        }


                        Autodesk.Civil.DatabaseServices.ProfileView profview1 = Trans1.GetObject(Rezultat_profileview.ObjectId, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.ProfileView;
                        Autodesk.Civil.DatabaseServices.Alignment align1 = Trans1.GetObject(Rezultat_alignment.ObjectId, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Alignment;


                        if (align1 != null && profview1 != null)
                        {
                            Functions.Creaza_layer("NO PLOT", 40, false);
                            Create_mleader_object_data_table();
                            bool run1 = true;


                            ObjectIdCollection col_surf = CivilDrawing.GetSurfaceIds();

                            ObjectId surfaceid1 = ObjectId.Null;


                            for (int j = 0; j < col_surf.Count; ++j)
                            {
                                Autodesk.Civil.DatabaseServices.Surface surf1 = Trans1.GetObject(col_surf[j], OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Surface;
                                if (surf1 != null)
                                {
                                    if (surf1.Name == comboBox_surface.Text)
                                    {
                                        surfaceid1 = col_surf[j];
                                    }
                                }
                            }

                            if (surfaceid1 != ObjectId.Null)
                            {
                                Autodesk.Civil.DatabaseServices.Surface surface1 = Trans1.GetObject(surfaceid1, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Surface;

                                do
                                {
                                    Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                    Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                    PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify point:");
                                    PP1.AllowNone = false;
                                    Point_res1 = Editor1.GetPoint(PP1);

                                    if (Point_res1.Status != PromptStatus.OK)
                                    {
                                        run1 = false;
                                    }

                                    if (run1 == true)
                                    {
                                        Point3d pickpt = Point_res1.Value;

                                        Point3d point_on_align = align1.GetClosestPointTo(pickpt, Vector3d.ZAxis, false);

                                        double elev1 = surface1.FindElevationAtXY(point_on_align.X, point_on_align.Y);

                                        double station_alg = -1.234;
                                        double off1 = 0;
                                        align1.StationOffset(point_on_align.X, point_on_align.Y, ref station_alg, ref off1);

                                        if (station_alg != -1.234)
                                        {
                                            double x_profview = -1.234;
                                            double y_profview = -1.234;
                                            profview1.FindXYAtStationAndElevation(station_alg, elev1, ref x_profview, ref y_profview);

                                            if (x_profview != -1.234 && y_profview != -1.234)
                                            {
                                                Point3d pt_on_poly = poly_centerline.GetClosestPointTo(pickpt, Vector3d.ZAxis, false);
                                                int idx1 = Convert.ToInt32(Math.Floor(poly_centerline.GetParameterAtPoint(pt_on_poly)));
                                                if (idx1 < dt_centerline.Rows.Count - 1)
                                                {
                                                    double x1 = Convert.ToDouble(dt_centerline.Rows[idx1][Col_x]);
                                                    double y1 = Convert.ToDouble(dt_centerline.Rows[idx1][Col_y]);

                                                    double x2 = Convert.ToDouble(dt_centerline.Rows[idx1 + 1][Col_x]);
                                                    double y2 = Convert.ToDouble(dt_centerline.Rows[idx1 + 1][Col_y]);

                                                    double sta1 = Convert.ToDouble(dt_centerline.Rows[idx1][Col_3DSta]);
                                                    double sta2 = Convert.ToDouble(dt_centerline.Rows[idx1 + 1][Col_3DSta]);

                                                    if (dt_centerline.Rows[idx1][Col_AheadSta] != DBNull.Value)
                                                    {
                                                        sta1 = Convert.ToDouble(dt_centerline.Rows[idx1][Col_AheadSta]);
                                                    }

                                                    if (dt_centerline.Rows[idx1 + 1][Col_BackSta] != DBNull.Value)
                                                    {
                                                        sta2 = Convert.ToDouble(dt_centerline.Rows[idx1 + 1][Col_BackSta]);
                                                    }

                                                    double d1 = Math.Pow(Math.Pow(point_on_align.X - x1, 2) + Math.Pow(point_on_align.Y - y1, 2), 0.5);
                                                    double d = Math.Pow(Math.Pow(x2 - x1, 2) + Math.Pow(y2 - y1, 2), 0.5);

                                                    double stax = sta1 + ((sta2 - sta1) * d1) / d;


                                                    string continut = Functions.Get_chainage_from_double(stax, "m", Convert.ToInt32(comboBox_round_sta.Text));
                                                    if (checkBox_sta.Checked == true)
                                                    {
                                                        continut = continut + "\r\nSTA " + Functions.Get_chainage_from_double(station_alg, "m", Convert.ToInt32(comboBox_round_sta.Text));
                                                    }
                                                    if (checkBox_elev.Checked == true)
                                                    {
                                                        continut = continut + "\r\nEL. " + Functions.Get_String_Rounded(elev1, Convert.ToInt32(comboBox_round_elev.Text));
                                                    }

                                                    MLeader ml1 = Functions.creaza_mleader(pickpt, continut, 0.5, 1.5, -3.5, 0.5, 0.5, 0.5, "NO PLOT");
                                                    MLeader ml2 = Functions.creaza_mleader(new Point3d(x_profview, y_profview, 0), continut, 0.5, 1.5, 3.5, 0.5, 0.5, 0.5, "NO PLOT");
                                                    ml1.ColorIndex = 1;
                                                    ml2.ColorIndex = 1;

                                                    Trans1.TransactionManager.QueueForGraphicsFlush();
                                                    lista_objid.Add(ml1.ObjectId);
                                                    lista_chain.Add(Functions.Get_chainage_from_double(stax, "m", Convert.ToInt32(comboBox_round_sta.Text)));
                                                    lista_objid.Add(ml2.ObjectId);
                                                    lista_chain.Add(Functions.Get_chainage_from_double(stax, "m", Convert.ToInt32(comboBox_round_sta.Text)));
                                                }
                                            }
                                        }
                                    }
                                } while (run1 == true);
                            }
                        }

                        if (lista_chain.Count > 0)
                        {
                            Append_object_data_to_ODXXX(lista_objid, segment_current, lista_chain);
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
            this.MdiParent.WindowState = FormWindowState.Normal;

        }

        private void button_load_surf_Click(object sender, EventArgs e)
        {
            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.Civil.ApplicationServices.CivilDocument CivilDrawing = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {
                set_enable_false();
                load_surfaces_to_combobox(comboBox_surface);

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            this.MdiParent.WindowState = FormWindowState.Normal;

            set_enable_true();
        }
        private void load_surfaces_to_combobox(ComboBox combo1)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.Civil.ApplicationServices.CivilDocument CivilDrawing = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument;
            using (DocumentLock lock1 = ThisDrawing.LockDocument())
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;
                    ObjectIdCollection col_surf = CivilDrawing.GetSurfaceIds();
                    combo1.Items.Clear();
                    for (int j = 0; j < col_surf.Count; ++j)
                    {
                        Autodesk.Civil.DatabaseServices.Surface surf1 = Trans1.GetObject(col_surf[j], OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Surface;
                        if (surf1 != null)
                        {
                            combo1.Items.Add(surf1.Name);
                        }
                    }
                    Trans1.Commit();
                }
            }
        }


        private void button_load_blocks_Click(object sender, EventArgs e)
        {
            Functions.Load_existing_Blocks_to_combobox(comboBox_blocks);
            if (comboBox_blocks.Items.Contains("TAG_XING") == true)
            {
                comboBox_blocks.SelectedIndex = comboBox_blocks.Items.IndexOf("TAG_XING");
            }

        }

        private void comboBox_blocks_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                List<string> lista1 = Functions.Incarca_existing_Atributes_to_list(comboBox_blocks.Text);
                comboBox_chainage.Items.Clear();
                comboBox_descr.Items.Clear();
                comboBox_elev.Items.Clear();
                comboBox_sta.Items.Clear();
                if (lista1.Count > 0)
                {
                    comboBox_chainage.Items.Add("");
                    comboBox_descr.Items.Add("");
                    comboBox_elev.Items.Add("");
                    comboBox_sta.Items.Add("");
                    for (int i = 0; i < lista1.Count; i++)
                    {
                        comboBox_chainage.Items.Add(lista1[i]);
                        comboBox_descr.Items.Add(lista1[i]);
                        comboBox_elev.Items.Add(lista1[i]);
                        comboBox_sta.Items.Add(lista1[i]);
                    }
                }


                if (comboBox_blocks.Text == "TAG_XING")
                {
                    if (lista1.Contains("STA") == true)
                    {
                        comboBox_sta.SelectedIndex = comboBox_sta.Items.IndexOf("STA");
                    }
                    if (lista1.Contains("CHAINAGE") == true)
                    {
                        comboBox_chainage.SelectedIndex = comboBox_chainage.Items.IndexOf("CHAINAGE");
                    }
                    if (lista1.Contains("ELEV") == true)
                    {
                        comboBox_elev.SelectedIndex = comboBox_elev.Items.IndexOf("ELEV");
                    }
                    if (lista1.Contains("DESCR") == true)
                    {
                        comboBox_descr.SelectedIndex = comboBox_descr.Items.IndexOf("DESCR");
                    }
                }

            }
            catch (System.Exception ex)
            {

                MessageBox.Show(ex.Message);
            }


        }

        private void button_place_blocks_on_profile_Click(object sender, EventArgs e)
        {
            if (dt_centerline == null || dt_centerline.Rows.Count <= 1 || poly_centerline == null)
            {
                MessageBox.Show("No Agen project is loaded\r\nOperation aborted.");
                return;
            }

            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.Civil.ApplicationServices.CivilDocument CivilDrawing = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument;

            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {
                set_enable_false();
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    List<ObjectId> lista_objid = new List<ObjectId>();
                    List<string> lista_chain = new List<string>();


                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;




                        this.MdiParent.WindowState = FormWindowState.Minimized;


                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_profileview;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions prompt_profview;
                        prompt_profview = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect Profile View (GRID):");
                        prompt_profview.SetRejectMessage("\nSelect a profile View!");
                        prompt_profview.AllowNone = true;
                        prompt_profview.AddAllowedClass(typeof(Autodesk.Civil.DatabaseServices.ProfileView), false);
                        Rezultat_profileview = ThisDrawing.Editor.GetEntity(prompt_profview);

                        if (Rezultat_profileview.Status != PromptStatus.OK)
                        {
                            Editor1.SetImpliedSelection(Empty_array);
                            Editor1.WriteMessage("\nCommand:");
                            set_enable_true();
                            this.MdiParent.WindowState = FormWindowState.Normal;
                            return;
                        }
                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_alignment;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions prompt_alg;
                        prompt_alg = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect Alignment:");
                        prompt_alg.SetRejectMessage("\nSelect an alignment!");
                        prompt_alg.AllowNone = true;
                        prompt_alg.AddAllowedClass(typeof(Autodesk.Civil.DatabaseServices.Alignment), false);
                        Rezultat_alignment = ThisDrawing.Editor.GetEntity(prompt_alg);

                        if (Rezultat_alignment.Status != PromptStatus.OK)
                        {
                            Editor1.SetImpliedSelection(Empty_array);
                            Editor1.WriteMessage("\nCommand:");
                            set_enable_true();
                            this.MdiParent.WindowState = FormWindowState.Normal;
                            return;
                        }


                        Autodesk.Civil.DatabaseServices.ProfileView profview1 = Trans1.GetObject(Rezultat_profileview.ObjectId, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.ProfileView;
                        Autodesk.Civil.DatabaseServices.Alignment align1 = Trans1.GetObject(Rezultat_alignment.ObjectId, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Alignment;


                        if (align1 != null && profview1 != null)
                        {
                            Functions.Creaza_layer("NO PLOT", 40, false);
                            string layer1 = "TEXT";
                            Functions.Creaza_layer(layer1, 3, true);
                            Create_mleader_object_data_table();
                            bool run1 = true;


                            ObjectIdCollection col_surf = CivilDrawing.GetSurfaceIds();

                            ObjectId surfaceid1 = ObjectId.Null;


                            for (int j = 0; j < col_surf.Count; ++j)
                            {
                                Autodesk.Civil.DatabaseServices.Surface surf1 = Trans1.GetObject(col_surf[j], OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Surface;
                                if (surf1 != null)
                                {
                                    if (surf1.Name == comboBox_surface.Text)
                                    {
                                        surfaceid1 = col_surf[j];
                                    }
                                }
                            }

                            if (surfaceid1 != ObjectId.Null)
                            {
                                Autodesk.Civil.DatabaseServices.Surface surface1 = Trans1.GetObject(surfaceid1, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Surface;

                                do
                                {
                                    Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                    Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                    PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify point:");
                                    PP1.AllowNone = false;
                                    Point_res1 = Editor1.GetPoint(PP1);

                                    if (Point_res1.Status != PromptStatus.OK)
                                    {
                                        run1 = false;
                                    }

                                    if (run1 == true)
                                    {
                                        Point3d pickpt = Point_res1.Value;

                                        Point3d point_on_align = align1.GetClosestPointTo(pickpt, Vector3d.ZAxis, false);

                                        double elev1 = surface1.FindElevationAtXY(point_on_align.X, point_on_align.Y);

                                        double station_alg = -1.234;
                                        double off1 = 0;
                                        align1.StationOffset(point_on_align.X, point_on_align.Y, ref station_alg, ref off1);

                                        if (station_alg != -1.234)
                                        {
                                            double x_profview = -1.234;
                                            double y_profview = -1.234;
                                            profview1.FindXYAtStationAndElevation(station_alg, elev1, ref x_profview, ref y_profview);

                                            if (x_profview != -1.234 && y_profview != -1.234)
                                            {
                                                Point3d pt_on_poly = poly_centerline.GetClosestPointTo(pickpt, Vector3d.ZAxis, false);
                                                int idx1 = Convert.ToInt32(Math.Floor(poly_centerline.GetParameterAtPoint(pt_on_poly)));
                                                if (idx1 < dt_centerline.Rows.Count - 1)
                                                {
                                                    double x1 = Convert.ToDouble(dt_centerline.Rows[idx1][Col_x]);
                                                    double y1 = Convert.ToDouble(dt_centerline.Rows[idx1][Col_y]);

                                                    double x2 = Convert.ToDouble(dt_centerline.Rows[idx1 + 1][Col_x]);
                                                    double y2 = Convert.ToDouble(dt_centerline.Rows[idx1 + 1][Col_y]);

                                                    double sta1 = Convert.ToDouble(dt_centerline.Rows[idx1][Col_3DSta]);
                                                    double sta2 = Convert.ToDouble(dt_centerline.Rows[idx1 + 1][Col_3DSta]);

                                                    if (dt_centerline.Rows[idx1][Col_AheadSta] != DBNull.Value)
                                                    {
                                                        sta1 = Convert.ToDouble(dt_centerline.Rows[idx1][Col_AheadSta]);
                                                    }

                                                    if (dt_centerline.Rows[idx1 + 1][Col_BackSta] != DBNull.Value)
                                                    {
                                                        sta2 = Convert.ToDouble(dt_centerline.Rows[idx1 + 1][Col_BackSta]);
                                                    }

                                                    double d1 = Math.Pow(Math.Pow(point_on_align.X - x1, 2) + Math.Pow(point_on_align.Y - y1, 2), 0.5);
                                                    double d = Math.Pow(Math.Pow(x2 - x1, 2) + Math.Pow(y2 - y1, 2), 0.5);

                                                    double stax = sta1 + ((sta2 - sta1) * d1) / d;

                                                    string continut = Functions.Get_chainage_from_double(stax, "m", Convert.ToInt32(comboBox_round_sta.Text));
                                                    if (checkBox_sta.Checked == true)
                                                    {
                                                        continut = continut + "\r\nSTA " + Functions.Get_chainage_from_double(station_alg, "m", Convert.ToInt32(comboBox_round_sta.Text));
                                                    }
                                                    if (checkBox_elev.Checked == true)
                                                    {
                                                        continut = continut + "\r\nEL. " + Functions.Get_String_Rounded(elev1, Convert.ToInt32(comboBox_round_elev.Text));
                                                    }

                                                    MLeader ml1 = Functions.creaza_mleader(pickpt, continut, 0.5, 1.5, -3.5, 0.5, 0.5, 0.5, "NO PLOT");
                                                    MLeader ml2 = Functions.creaza_mleader(new Point3d(x_profview, y_profview, 0), continut, 0.5, 1.5, 3.5, 0.5, 0.5, 0.5, "NO PLOT");
                                                    ml1.ColorIndex = 1;
                                                    ml2.ColorIndex = 1;

                                                    BlockReference br1 = null;
                                                    if (comboBox_blocks.Text != "")
                                                    {
                                                        System.Collections.Specialized.StringCollection col_atr = new System.Collections.Specialized.StringCollection();
                                                        System.Collections.Specialized.StringCollection col_val = new System.Collections.Specialized.StringCollection();
                                                        col_atr.Add(comboBox_sta.Text);
                                                        col_val.Add(Functions.Get_chainage_from_double(station_alg, "m", Convert.ToInt32(comboBox_round_sta.Text)));
                                                        col_atr.Add(comboBox_chainage.Text);
                                                        col_val.Add("(" + Functions.Get_chainage_from_double(stax, "m", Convert.ToInt32(comboBox_round_sta.Text)) + ")");
                                                        col_atr.Add(comboBox_elev.Text);
                                                        col_val.Add(Functions.Get_String_Rounded(elev1, Convert.ToInt32(comboBox_round_elev.Text)));


                                                        br1 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "", comboBox_blocks.Text, new Point3d(x_profview, y_profview, 0), 1, 0, layer1, col_atr, col_val);
                                                        lista_objid.Add(br1.ObjectId);
                                                        lista_chain.Add(Functions.Get_chainage_from_double(stax, "m", Convert.ToInt32(comboBox_round_sta.Text)));

                                                    }


                                                    Trans1.TransactionManager.QueueForGraphicsFlush();
                                                    lista_objid.Add(ml1.ObjectId);
                                                    lista_chain.Add(Functions.Get_chainage_from_double(stax, "m", Convert.ToInt32(comboBox_round_sta.Text)));
                                                    lista_objid.Add(ml2.ObjectId);
                                                    lista_chain.Add(Functions.Get_chainage_from_double(stax, "m", Convert.ToInt32(comboBox_round_sta.Text)));
                                                }
                                            }
                                        }
                                    }
                                } while (run1 == true);
                            }
                        }

                        if (lista_chain.Count > 0)
                        {
                            Append_object_data_to_ODXXX(lista_objid, segment_current, lista_chain);
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
            this.MdiParent.WindowState = FormWindowState.Normal;

        }

        private void button_calc_staZ_of_block_Click(object sender, EventArgs e)
        {
            if (dt_centerline == null || dt_centerline.Rows.Count <= 1 || poly_centerline == null)
            {
                MessageBox.Show("No Agen project is loaded\r\nOperation aborted.");
                return;
            }

            if (comboBox_chainage.Text == "" && comboBox_elev.Text == "" && comboBox_sta.Text == "")
            {
                MessageBox.Show("No Block Attributes specified\r\nOperation aborted.");
                return;
            }

            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.Civil.ApplicationServices.CivilDocument CivilDrawing = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument;

            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {
                set_enable_false();
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    List<ObjectId> lista_objid = new List<ObjectId>();
                    List<string> lista_chain = new List<string>();


                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        this.MdiParent.WindowState = FormWindowState.Minimized;





                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_profile;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions prompt_prof;
                        prompt_prof = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect profile (GROUND):");
                        prompt_prof.SetRejectMessage("\nSelect a profile!");
                        prompt_prof.AllowNone = true;
                        prompt_prof.AddAllowedClass(typeof(Autodesk.Civil.DatabaseServices.Profile), false);
                        Rezultat_profile = ThisDrawing.Editor.GetEntity(prompt_prof);

                        if (Rezultat_profile.Status != PromptStatus.OK)
                        {
                            Editor1.SetImpliedSelection(Empty_array);
                            Editor1.WriteMessage("\nCommand:");
                            set_enable_true();
                            this.MdiParent.WindowState = FormWindowState.Normal;
                            return;
                        }


                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_profileview;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions prompt_profview;
                        prompt_profview = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect Profile View (GRID):");
                        prompt_profview.SetRejectMessage("\nSelect a profile View!");
                        prompt_profview.AllowNone = true;
                        prompt_profview.AddAllowedClass(typeof(Autodesk.Civil.DatabaseServices.ProfileView), false);
                        Rezultat_profileview = ThisDrawing.Editor.GetEntity(prompt_profview);

                        if (Rezultat_profileview.Status != PromptStatus.OK)
                        {
                            Editor1.SetImpliedSelection(Empty_array);
                            Editor1.WriteMessage("\nCommand:");
                            set_enable_true();
                            this.MdiParent.WindowState = FormWindowState.Normal;
                            return;
                        }

                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_blocks;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez.MessageForAdding = "\nSelect blocks:";
                        Prompt_rez.SingleOnly = false;
                        Rezultat_blocks = ThisDrawing.Editor.GetSelection(Prompt_rez);

                        if (Rezultat_blocks.Status != PromptStatus.OK)
                        {
                            Editor1.SetImpliedSelection(Empty_array);
                            Editor1.WriteMessage("\nCommand:");
                            set_enable_true();
                            this.MdiParent.WindowState = FormWindowState.Normal; ;
                            return;
                        }

                        Autodesk.Civil.DatabaseServices.ProfileView profview1 = Trans1.GetObject(Rezultat_profileview.ObjectId, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.ProfileView;
                        Autodesk.Civil.DatabaseServices.Profile prof1 = Trans1.GetObject(Rezultat_profile.ObjectId, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Profile;


                        if (prof1 != null && profview1 != null)
                        {

                            Create_mleader_object_data_table();
                            ObjectId align_id = prof1.AlignmentId;
                            Autodesk.Civil.DatabaseServices.Alignment align1 = Trans1.GetObject(align_id, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Alignment;
                            if (align1 != null)
                            {
                                for (int i = 0; i < Rezultat_blocks.Value.Count; i++)
                                {
                                    BlockReference block1 = Trans1.GetObject(Rezultat_blocks.Value[i].ObjectId, OpenMode.ForWrite) as BlockReference;
                                    if (block1 != null)
                                    {
                                        if (block1.AttributeCollection.Count > 0)
                                        {
                                            double sta = -1.234;
                                            double elev = -1.234;
                                            profview1.FindStationAndElevationAtXY(block1.Position.X, block1.Position.Y, ref sta, ref elev);
                                            double x = -1.234;
                                            double y = -1.234;
                                            align1.PointLocation(sta, 0, ref x, ref y);
                                            Point3d pt_on_poly = poly_centerline.GetClosestPointTo(new Point3d(x, y, 0), Vector3d.ZAxis, false);
                                            int idx1 = Convert.ToInt32(Math.Floor(poly_centerline.GetParameterAtPoint(pt_on_poly)));
                                            if (idx1 < dt_centerline.Rows.Count - 1)
                                            {
                                                double x1 = Convert.ToDouble(dt_centerline.Rows[idx1][Col_x]);
                                                double y1 = Convert.ToDouble(dt_centerline.Rows[idx1][Col_y]);

                                                double x2 = Convert.ToDouble(dt_centerline.Rows[idx1 + 1][Col_x]);
                                                double y2 = Convert.ToDouble(dt_centerline.Rows[idx1 + 1][Col_y]);

                                                double sta1 = Convert.ToDouble(dt_centerline.Rows[idx1][Col_3DSta]);
                                                double sta2 = Convert.ToDouble(dt_centerline.Rows[idx1 + 1][Col_3DSta]);

                                                if (dt_centerline.Rows[idx1][Col_AheadSta] != DBNull.Value)
                                                {
                                                    sta1 = Convert.ToDouble(dt_centerline.Rows[idx1][Col_AheadSta]);
                                                }

                                                if (dt_centerline.Rows[idx1 + 1][Col_BackSta] != DBNull.Value)
                                                {
                                                    sta2 = Convert.ToDouble(dt_centerline.Rows[idx1 + 1][Col_BackSta]);
                                                }

                                                double d1 = Math.Pow(Math.Pow(x - x1, 2) + Math.Pow(y - y1, 2), 0.5);
                                                double d = Math.Pow(Math.Pow(x2 - x1, 2) + Math.Pow(y2 - y1, 2), 0.5);

                                                double stax = sta1 + ((sta2 - sta1) * d1) / d;

                                                string continut_chainage = Functions.Get_chainage_from_double(stax, "m", Convert.ToInt32(comboBox_round_sta.Text));
                                                string continut_sta = Functions.Get_chainage_from_double(sta, "m", Convert.ToInt32(comboBox_round_sta.Text));
                                                string continut_elev = Functions.Get_String_Rounded(elev, Convert.ToInt32(comboBox_round_elev.Text));

                                                foreach (ObjectId id1 in block1.AttributeCollection)
                                                {
                                                    AttributeReference atr1 = Trans1.GetObject(id1, OpenMode.ForWrite) as AttributeReference;
                                                    if (atr1 != null)
                                                    {
                                                        if (atr1.Tag == comboBox_elev.Text)
                                                        {
                                                            atr1.TextString = continut_elev;
                                                        }
                                                        if (atr1.Tag == comboBox_sta.Text)
                                                        {
                                                            atr1.TextString = continut_sta;
                                                        }
                                                        if (atr1.Tag == comboBox_chainage.Text)
                                                        {
                                                            atr1.TextString = continut_chainage;
                                                        }
                                                    }
                                                }
                                                lista_objid.Add(block1.ObjectId);
                                                lista_chain.Add(Functions.Get_chainage_from_double(stax, "m", Convert.ToInt32(comboBox_round_sta.Text)));
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        if (lista_chain.Count > 0)
                        {
                            Append_object_data_to_ODXXX(lista_objid, segment_current, lista_chain);
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
            this.MdiParent.WindowState = FormWindowState.Normal;

        }


        private void button_calcZ_of_text_Click(object sender, EventArgs e)
        {
            if (comboBox_surface.Text == "")
            {
                MessageBox.Show("no surface selected");
                return;
            }

            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.Civil.ApplicationServices.CivilDocument CivilDrawing = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument;
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
                        ObjectIdCollection col_surf = CivilDrawing.GetSurfaceIds();
                        for (int j = 0; j < col_surf.Count; ++j)
                        {
                            Autodesk.Civil.DatabaseServices.Surface surf1 = Trans1.GetObject(col_surf[j], OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Surface;
                            if (surf1 != null)
                            {
                                if (surf1.Name == comboBox_surface.Text)
                                {
                                    System.Data.DataTable dt1 = new System.Data.DataTable();

                                    dt1.Columns.Add("Point Name", typeof(string));
                                    dt1.Columns.Add("X", typeof(double));
                                    dt1.Columns.Add("Y", typeof(double));
                                    dt1.Columns.Add("Z picked", typeof(double));
                                    dt1.Columns.Add("Z on surface", typeof(double));
                                    dt1.Columns.Add("cover", typeof(double));
                                    dt1.Columns.Add("Surface", typeof(string));
                                    dt1.Columns.Add("USER", typeof(string));
                                    bool run1 = true;
                                    do
                                    {

                                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                                        Prompt_rez.MessageForAdding = "\nSelect point:";
                                        Prompt_rez.SingleOnly = false;
                                        Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                                        if (Rezultat1.Status != PromptStatus.OK)
                                        {
                                            run1 = false;
                                        }
                                        else
                                        {
                                            for (int i = 0; i < Rezultat1.Value.Count; i++)
                                            {
                                                DBText text1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as DBText;
                                                if (text1 != null)
                                                {

                                                    Point3d ptins = text1.Position;
                                                    double x = Math.Round(ptins.X, 3);
                                                    double y = Math.Round(ptins.Y, 3);
                                                    double z = Math.Round(ptins.Z, 3);

                                                    double elev1 = -1.23;
                                                    try
                                                    {
                                                        elev1 = surf1.FindElevationAtXY(ptins.X, ptins.Y);
                                                        ThisDrawing.Editor.WriteMessage("\nX=" + Convert.ToString(x) + ", Y=" + Convert.ToString(y) +
                                                                                         ", Z=" + Convert.ToString(z) + "(" + Convert.ToString(Math.Round(elev1, 3)) + " on surface)");
                                                    }
                                                    catch (System.Exception)
                                                    {
                                                    }

                                                    dt1.Rows.Add();
                                                    dt1.Rows[dt1.Rows.Count - 1][0] = text1.TextString;
                                                    dt1.Rows[dt1.Rows.Count - 1][1] = x;
                                                    dt1.Rows[dt1.Rows.Count - 1][2] = y;
                                                    dt1.Rows[dt1.Rows.Count - 1][3] = z;
                                                    dt1.Rows[dt1.Rows.Count - 1][4] = Math.Round(elev1, 3);
                                                    dt1.Rows[dt1.Rows.Count - 1][5] = Math.Round(elev1, 3) - z;
                                                    dt1.Rows[dt1.Rows.Count - 1][6] = comboBox_surface.Text;
                                                    dt1.Rows[dt1.Rows.Count - 1][7] = Environment.UserName.ToUpper();
                                                }
                                            }

                                        }
                                    } while (run1 == true);

                                    string name1 = System.DateTime.Now.Year + "-" + System.DateTime.Now.Month + "-" + System.DateTime.Now.Day + "_" + System.DateTime.Now.Hour + "h" + System.DateTime.Now.Minute + "m";

                                    Functions.Transfer_datatable_to_new_excel_spreadsheet(dt1, name1);

                                    dt1 = null;
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
            this.MdiParent.WindowState = FormWindowState.Normal;

            set_enable_true();
        }

        private void button_pick_pipe_elevations_Click(object sender, EventArgs e)
        {


            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.Civil.ApplicationServices.CivilDocument CivilDrawing = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

            System.Data.DataTable dt1 = new System.Data.DataTable();

            dt1.Columns.Add("Point Name", typeof(string));
            dt1.Columns.Add("X", typeof(double));
            dt1.Columns.Add("Y", typeof(double));
            dt1.Columns.Add("Z picked", typeof(double));
            dt1.Columns.Add("USER", typeof(string));


            try
            {
                set_enable_false();
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    bool run1 = true;
                    this.MdiParent.WindowState = FormWindowState.Minimized;
                    do
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                            ObjectIdCollection col_surf = CivilDrawing.GetSurfaceIds();

                            Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rez.MessageForAdding = "\nSelect point label (MTEXT):";
                            Prompt_rez.SingleOnly = true;
                            Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                            if (Rezultat1.Status != PromptStatus.OK)
                            {
                                run1 = false;
                            }
                            else
                            {
                                DBText text1 = Trans1.GetObject(Rezultat1.Value[0].ObjectId, OpenMode.ForWrite) as DBText;
                                MText mtext1 = Trans1.GetObject(Rezultat1.Value[0].ObjectId, OpenMode.ForWrite) as MText;
                                if (text1 != null || mtext1 != null)
                                {
                                    string pn = "xxx";
                                    if (text1 != null)
                                    {
                                        pn = text1.TextString;
                                    }

                                    if (mtext1 != null)
                                    {
                                        pn = mtext1.Contents;
                                    }

                                    Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                    Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                    PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify the XY position point");
                                    PP1.AllowNone = false;
                                    Point_res1 = Editor1.GetPoint(PP1);

                                    if (Point_res1.Status != PromptStatus.OK)
                                    {
                                        run1 = false;
                                    }

                                    if (run1 == true)
                                    {
                                        Point3d ptins = Point_res1.Value;

                                        double x = Math.Round(ptins.X, 3);
                                        double y = Math.Round(ptins.Y, 3);
                                        double z = Math.Round(ptins.Z, 3);

                                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP2;
                                        PP2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify the elevation point");
                                        PP2.AllowNone = false;
                                        Point_res2 = Editor1.GetPoint(PP2);
                                        if (Point_res2.Status == PromptStatus.OK)
                                        {
                                            z = Math.Round(Point_res2.Value.Z, 3);
                                        }



                                        dt1.Rows.Add();
                                        dt1.Rows[dt1.Rows.Count - 1][0] = pn;
                                        dt1.Rows[dt1.Rows.Count - 1][1] = x;
                                        dt1.Rows[dt1.Rows.Count - 1][2] = y;
                                        dt1.Rows[dt1.Rows.Count - 1][3] = z;
                                        dt1.Rows[dt1.Rows.Count - 1]["USER"] = Environment.UserName.ToUpper();


                                        for (int j = 0; j < col_surf.Count; ++j)
                                        {
                                            Autodesk.Civil.DatabaseServices.Surface surf1 = Trans1.GetObject(col_surf[j], OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Surface;
                                            if (surf1 != null)
                                            {
                                                string surf_name = surf1.Name;
                                                double elev1 = -123456789.123;
                                                try
                                                {
                                                    elev1 = surf1.FindElevationAtXY(ptins.X, ptins.Y);
                                                    ThisDrawing.Editor.WriteMessage("\n" + pn + ": X=" + Convert.ToString(x) + ", Y=" + Convert.ToString(y) +
                                                                                     ", Z=" + Convert.ToString(z) + "(" + Convert.ToString(Math.Round(elev1, 3)) + " on " + surf_name + ")");
                                                }
                                                catch (System.Exception)
                                                {
                                                }

                                                if (dt1.Columns.Contains("Z on " + surf_name) == false)
                                                {
                                                    dt1.Columns.Add("Z on " + surf_name, typeof(double));
                                                }

                                                dt1.Rows[dt1.Rows.Count - 1]["Z on " + surf_name] = Math.Round(elev1, 3);
                                            }
                                        }

                                        if (text1 != null)
                                        {

                                            text1.ColorIndex = 1;
                                        }

                                        if (mtext1 != null)
                                        {

                                            mtext1.ColorIndex = 1;

                                        }
                                    }
                                }
                            }
                            Trans1.Commit();
                        }

                    } while (run1 == true);

                    string name1 = System.DateTime.Now.Year + "-" + System.DateTime.Now.Month + "-" + System.DateTime.Now.Day + "_" + System.DateTime.Now.Hour + "h" + System.DateTime.Now.Minute + "m";

                    Functions.Transfer_datatable_to_new_excel_spreadsheet(dt1, name1);

                    dt1 = null;
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            this.MdiParent.WindowState = FormWindowState.Normal;

            set_enable_true();
        }
    }
}
