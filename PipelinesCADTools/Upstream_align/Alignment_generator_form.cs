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





namespace Alignment_generator
{
    public partial class form_AGEN : Form
    {
        private bool clickdragdown;
        private Point lastLocation;

        private bool mouseDown;
        private ContextMenuStrip collectionRoundMenuStrip;

        double Vw_scale = 1;
        double Vw_height = 0;
        double Vw_width = 0;
        double Vw_ps_x = 0;
        double Vw_ps_y = 0;



        double Match_distance = 5280;
        string Layer_name_ML_rectangle = "Template_ML";
        string Layer_name_VP_rectangle = "Template_VP";
        string Layer_North_Arrow = "NORTH";
        string Layer_name_Main_Viewport = "VP";

        Polyline Poly2D;
        string Layer_name_Poly2D = "0";
        Int16 Color_index_Layer_name_Poly2D = 7;
        string Poly2D_handle = "";

        string NA_name = "";
        double NA_x = 0;
        double NA_y = 0;
        double NA_scale = 0;

        bool Freeze_operations = false;


        System.Data.DataTable Data_table_matchline;
        System.Data.DataTable Data_table_Main_VP;
        System.Data.DataTable Data_table_centerline;
        System.Data.DataTable Data_table_Config_files_path;
        System.Data.DataTable Data_table_viewport_target_areas;


        string CL_file = "";
        string Sheet_index_file = "";


        public form_AGEN()
        {
            InitializeComponent();

            //tabControl_Nav.Region = new System.Drawing.Region(tabControl_Nav.DisplayRectangle);
            //tabControl_work.Region = new System.Drawing.Region(tabControl_Nav.DisplayRectangle);
            panel_sheet_index_basefile.Visible = false;
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
            this.Close();
        }



        private void mouserovercolor_coolblue_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button)
            {
                Button entercolorevent = (Button)sender;
                ForeColor = entercolorevent.ForeColor = Color.FromArgb(0, 122, 204);
                entercolorevent.ForeColor = Color.FromArgb(0, 122, 204);
            }
            if (sender is LinkLabel)
            {
                LinkLabel entercolorevent = (LinkLabel)sender;
                ForeColor = entercolorevent.LinkColor = Color.FromArgb(0, 122, 204);
                entercolorevent.LinkColor = Color.FromArgb(0, 122, 204);
            }
        }


        private void mouserovercolor_orange_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button)
            {
                Button entercolorevent = (Button)sender;
                ForeColor = entercolorevent.ForeColor = Color.Orange;
                entercolorevent.ForeColor = Color.Orange;
            }
            if (sender is LinkLabel)
            {
                LinkLabel entercolorevent = (LinkLabel)sender;
                ForeColor = entercolorevent.LinkColor = Color.Orange;
                entercolorevent.LinkColor = Color.Orange;
            }
        }

        private void mouseovercolor_white_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button)
            {
                Button leavecolorevent = (Button)sender;
                leavecolorevent.ForeColor = Color.White;
            }

            if (sender is LinkLabel)
            {
                LinkLabel leavecolorevent = (LinkLabel)sender;
                leavecolorevent.LinkColor = Color.White;
            }
        }

        private void contactsupport_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process proc = new System.Diagnostics.Process();
            proc.StartInfo.FileName = "mailto:Support.CADTechnolgies@mottmacna.com?subject=DMax Help";
            proc.Start();
        }

        private void help_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
        {
            Process.Start(@"G:\_HMM\Design\Programming\Documentation");
        }

        private void inputNumber_KeyPress_integer(object sender, KeyPressEventArgs e)
        {

            if (char.IsControl(e.KeyChar) == false && char.IsDigit(e.KeyChar) == false)
            {
                e.Handled = true;
            }


        }

        private void inputNumber_KeyPress_double(object sender, KeyPressEventArgs e)
        {

            if (char.IsControl(e.KeyChar) == false && char.IsDigit(e.KeyChar) == false && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }

            if ((e.KeyChar == '.') && ((sender as TextBox).Text.Contains(".") == true))
            {
                e.Handled = true;
            }
        }

        private void button_navbar_alignments_config_Click(object sender, EventArgs e)
        {
            tabControl_work.SelectedTab = tabPageA;

        }

        private void button_navbar_align_sheetindex_Click(object sender, EventArgs e)
        {
            tabControl_work.SelectedTab = tabPageB;
        }

        private void Read_config_file_Click(object sender, EventArgs e)
        {
            tabControl_work.SelectedTab = tabPageA;
            using (OpenFileDialog fbd = new OpenFileDialog())
            {
                fbd.Multiselect = false;
                fbd.Filter = "Excel files (*.xlsx)|*.xlsx";

                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    String File1 = fbd.FileName;

                    Load_existing_config_file(File1);
                    tabControl_work.SelectedTab = tabPageB;
                    dataGridView_config_files_path.DataSource = Data_table_Config_files_path;
                    dataGridView_config_files_path.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                    tabControl_work.SelectedTab = tabPageA;
                }

            }
        }

        private void Load_existing_config_file(String File1)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application Excel1 = new Microsoft.Office.Interop.Excel.Application();
                if (Excel1 == null)
                {
                    MessageBox.Show("PROBLEM WITH EXCEL!");
                    return;
                }

                Excel1.Visible = true;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(File1);
                Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];

                try
                {

                    Data_table_Config_files_path = new System.Data.DataTable();
                    Data_table_Config_files_path.Columns.Add("TYPE", typeof(String));
                    Data_table_Config_files_path.Columns.Add("PATH", typeof(String));

                    textBox_client_name.Text = W1.Range["B1"].Value2.ToString();
                    textBox_project_name.Text = W1.Range["B2"].Value2.ToString();
                    textBox_segment_name.Text = W1.Range["B3"].Value2.ToString();

                    string Template = W1.Range["B4"].Value2.ToString();
                    if (System.IO.File.Exists(Template) == true)
                    {
                        textBox_template_name.Text = Template;
                    }
                    else
                    {
                        MessageBox.Show("the dwt file specified does not exists\r\n" + Template);
                        return;
                    }

                    string Output = W1.Range["B5"].Value2.ToString();
                    if (System.IO.Directory.Exists(Output) == true)
                    {
                        textBox_output_folder.Text = Output;
                    }
                    else
                    {
                        MessageBox.Show("the output folder specified does not exists\r\n" + Output);
                        return;
                    }

                    string FileName = W1.Range["B6"].Value2.ToString();
                    if (FileName != "")
                    {
                        textBox_prefix_name.Text = FileName;
                    }
                    else
                    {
                        MessageBox.Show("the file prefix is not specified");
                        return;
                    }

                    string Startno = W1.Range["B7"].Value2.ToString();
                    if (Functions.IsNumeric(Startno) == true)
                    {
                        textBox_name_start_number.Text = Startno;
                    }
                    else
                    {
                        MessageBox.Show("the start number has to be a number to be incremented\r\nplease verify");
                        return;
                    }


                    string Increment = W1.Range["B8"].Value2.ToString();


                    if (Functions.IsNumeric(Increment) == true)
                    {
                        textBox_name_increment.Text = Increment;
                    }
                    else
                    {
                        MessageBox.Show("the increment has to be numeric!\r\nplease verify");
                        return;
                    }


                    string s4 = W1.Range["B9"].Value2.ToString();
                    if (Functions.IsNumeric(s4) == true)
                    {
                        Vw_ps_x = Convert.ToDouble(s4);

                    }
                    else
                    {
                        MessageBox.Show("Not numeric value for the main viewport X paperspace position");
                        return;
                    }

                    string s5 = W1.Range["B10"].Value2.ToString();
                    if (Functions.IsNumeric(s5) == true)
                    {
                        Vw_ps_y = Convert.ToDouble(s5);
                    }
                    else
                    {
                        MessageBox.Show("Not numeric value for the main viewport Y paperspace position");
                        return;
                    }

                    string s2 = W1.Range["B11"].Value2.ToString();
                    if (Functions.IsNumeric(s2) == true)
                    {
                        Vw_width = Convert.ToDouble(s2);
                    }
                    else
                    {
                        MessageBox.Show("Not numeric value for the main viewport width");
                        return;
                    }


                    textBox_config_file_location.Text = File1;
                    string s1 = W1.Range["B12"].Value2.ToString();
                    if (Functions.IsNumeric(s1) == true)
                    {
                        Vw_height = Convert.ToDouble(s1);

                    }
                    else
                    {
                        MessageBox.Show("Not numeric value for the main viewport height");
                        return;
                    }

                    string s3 = W1.Range["B13"].Value2.ToString();
                    if (Functions.IsNumeric(s3) == true)
                    {
                        Vw_scale = Convert.ToDouble(s3);

                    }
                    else
                    {
                        MessageBox.Show("Not numeric value for the main viewport scale");
                        return;
                    }


                    string basef = W1.Range["B14"].Value2.ToString();
                    if (System.IO.Directory.Exists(basef) == true)
                    {
                        textBox_basefiles_folder.Text = basef;
                    }
                    else
                    {
                        MessageBox.Show("the basefile folder is not specified");
                        return;
                    }

                    string Matchlines = W1.Range["B15"].Value2.ToString();
                    if (System.IO.File.Exists(Matchlines) == true)
                    {
                        Data_table_Config_files_path.Rows.Add();
                        Data_table_Config_files_path.Rows[Data_table_Config_files_path.Rows.Count - 1][0] = "Sheet Index";
                        Data_table_Config_files_path.Rows[Data_table_Config_files_path.Rows.Count - 1][1] = Matchlines;
                        Sheet_index_file = Matchlines;

                        Microsoft.Office.Interop.Excel.Workbook Workbook2 = null;
                        Microsoft.Office.Interop.Excel.Worksheet W2 = null;
                        try
                        {

                            Workbook2 = Excel1.Workbooks.Open(Matchlines);
                            W2 = Workbook2.Worksheets[1];
                            Data_table_matchline = Functions.Build_Data_table_matchline_from_excel(W2);
                            dataGridView_sheet_index.DataSource = Data_table_matchline;
                            dataGridView_sheet_index.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                            Workbook2.Close();


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
                        MessageBox.Show("the sheet index file does not exists\r\n" + Matchlines);
                        return;
                    }

                    string Centerline = W1.Range["B16"].Value2.ToString();
                    if (System.IO.File.Exists(Centerline) == true)
                    {
                        Data_table_Config_files_path.Rows.Add();
                        Data_table_Config_files_path.Rows[Data_table_Config_files_path.Rows.Count - 1][0] = "Centerline";
                        Data_table_Config_files_path.Rows[Data_table_Config_files_path.Rows.Count - 1][1] = Centerline;
                        CL_file = Centerline;
                        Microsoft.Office.Interop.Excel.Workbook Workbook3 = null;
                        Microsoft.Office.Interop.Excel.Worksheet W3 = null;
                        try
                        {

                            Workbook3 = Excel1.Workbooks.Open(Centerline);
                            W3 = Workbook3.Worksheets[1];
                            Data_table_centerline = Functions.Build_Data_table_centerline_from_excel(W3);
                            Workbook3.Close();


                        }
                        catch (System.Exception ex)
                        {
                            System.Windows.Forms.MessageBox.Show(ex.Message);

                        }
                        finally
                        {
                            if (W3 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W3);
                            if (Workbook3 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook3);

                        }

                    }
                    else
                    {
                        MessageBox.Show("the centerline file does not exists\r\n" + Centerline);
                        return;
                    }

                    NA_name = W1.Range["B17"].Value2.ToString();

                    string Xna = W1.Range["B18"].Value2.ToString();
                    if (Functions.IsNumeric(Xna) == true)
                    {
                        NA_x = Convert.ToDouble(Xna);
                    }

                    string yna = W1.Range["B19"].Value2.ToString();
                    if (Functions.IsNumeric(yna) == true)
                    {
                        NA_y = Convert.ToDouble(yna);
                    }

                    string sc = W1.Range["B20"].Value2.ToString();
                    if (Functions.IsNumeric(sc) == true)
                    {
                        NA_scale = Convert.ToDouble(sc);
                    }
                    string ln = W1.Range["B21"].Value2.ToString();
                    Layer_name_Poly2D = ln;

                    string ciln = W1.Range["B22"].Value2.ToString();
                    if (Functions.IsNumeric(ciln) == true)
                    {
                        Color_index_Layer_name_Poly2D = Convert.ToInt16(Math.Abs(Convert.ToDouble(ciln)));
                        if (Color_index_Layer_name_Poly2D == 0 | Color_index_Layer_name_Poly2D > 255) Color_index_Layer_name_Poly2D = 7;

                    }

                    object hi =W1.Range["B23"].Value2;
                    if (hi != null)
                    {
                        Poly2D_handle = hi.ToString();
                    }
                    Workbook1.Close();
                    Excel1.Quit();
                    Populate_dataGridView_viewport_target_areas();
                }
                catch (System.Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);

                }
                finally
                {
                    if (W1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W1);
                    if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                    if (Excel1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }

        }

        private void Populate_dataGridView_viewport_target_areas()
        {
            if (Vw_height != 0 && Vw_width != 0)
            {

                Data_table_viewport_target_areas = new System.Data.DataTable();
                Data_table_viewport_target_areas.Columns.Add("TYPE", typeof(String));
                Data_table_viewport_target_areas.Columns.Add("CUSTOMSCALE", typeof(double));
                Data_table_viewport_target_areas.Columns.Add("WIDTH", typeof(double));
                Data_table_viewport_target_areas.Columns.Add("HEIGHT", typeof(double));
                Data_table_viewport_target_areas.Columns.Add("PS_X", typeof(double));
                Data_table_viewport_target_areas.Columns.Add("PS_Y", typeof(double));

                Data_table_viewport_target_areas.Rows.Add();
                Data_table_viewport_target_areas.Rows[Data_table_viewport_target_areas.Rows.Count - 1][0] = comboBox_viewport_target_areas.Items[1];
                Data_table_viewport_target_areas.Rows[Data_table_viewport_target_areas.Rows.Count - 1][1] = Vw_scale;
                Data_table_viewport_target_areas.Rows[Data_table_viewport_target_areas.Rows.Count - 1][2] = Vw_width;
                Data_table_viewport_target_areas.Rows[Data_table_viewport_target_areas.Rows.Count - 1][3] = Vw_height;
                Data_table_viewport_target_areas.Rows[Data_table_viewport_target_areas.Rows.Count - 1][4] = Vw_ps_x;
                Data_table_viewport_target_areas.Rows[Data_table_viewport_target_areas.Rows.Count - 1][5] = Vw_ps_y;
                dataGridView_viewport_target_areas.DataSource = Data_table_viewport_target_areas;
                dataGridView_viewport_target_areas.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            }
        }

        private void button_place_Match_rectangles_Click(object sender, EventArgs e)
        {
            if (System.IO.File.Exists(textBox_config_file_location.Text) == false)
            {
                MessageBox.Show("no config file loaded\r\nOperation aborted");
                return;
            }

            if (Functions.IsNumeric(TextBox_matchline_length.Text) == false)
            {
                MessageBox.Show("no matchlines distance specified\r\nOperation aborted");
                return;
            }

            if (Freeze_operations == false)
            {
                Freeze_operations = true;

                Erase_viewports_templates();

                Create_ML_object_data();

                if (Functions.IsNumeric(TextBox_matchline_length.Text) == true)
                {
                    Match_distance = Convert.ToDouble(TextBox_matchline_length.Text);
                }



                ObjectId[] Empty_array = null;

                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                try
                {

                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {
                        Functions.Creaza_layer(Layer_name_ML_rectangle, 4, false);

                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_optionsCL = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect Centerline:");
                            Prompt_optionsCL.SetRejectMessage("\nYou did not selected a polyline");
                            Prompt_optionsCL.AddAllowedClass(typeof(Polyline), true);

                            Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_CL = Editor1.GetEntity(Prompt_optionsCL);
                            if (Rezultat_CL.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                            {
                                Freeze_operations = false;
                                Editor1.SetImpliedSelection(Empty_array);
                                Editor1.WriteMessage("\nCommand:");
                                return;
                            }
                            Poly2D = (Polyline)Trans1.GetObject(Rezultat_CL.ObjectId, OpenMode.ForRead);
                            Layer_name_Poly2D = Poly2D.Layer;

                            BlockTable BlockTable1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                            //Dim BTrecord_MS As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(BlockTable1(BlockTableRecord.ModelSpace), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                            BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                            Autodesk.AutoCAD.DatabaseServices.LayerTable LayerTable1 = (Autodesk.AutoCAD.DatabaseServices.LayerTable)Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                            LayerTableRecord Layer_poly2d = (LayerTableRecord)Trans1.GetObject(LayerTable1[Layer_name_Poly2D], OpenMode.ForRead);
                            Color_index_Layer_name_Poly2D = Layer_poly2d.Color.ColorIndex;
                            Poly2D_handle = Poly2D.ObjectId.Handle.Value.ToString();

                            double dist1 = 0;

                            Data_table_matchline = new System.Data.DataTable();
                            Data_table_matchline.Columns.Add("OBJECT_ID", typeof(string));
                            Data_table_matchline.Columns.Add("M1", typeof(Double));
                            Data_table_matchline.Columns.Add("M2", typeof(Double));
                            Data_table_matchline.Columns.Add("X", typeof(Double));
                            Data_table_matchline.Columns.Add("Y", typeof(Double));
                            Data_table_matchline.Columns.Add("ROTATION", typeof(Double));
                            Data_table_matchline.Columns.Add("WIDTH", typeof(Double));
                            Data_table_matchline.Columns.Add("HEIGHT", typeof(Double));
                            Data_table_matchline.Columns.Add("FILE_NAME", typeof(string));

                            Data_table_centerline = new System.Data.DataTable();
                            Data_table_centerline.Columns.Add("X", typeof(Double));
                            Data_table_centerline.Columns.Add("Y", typeof(Double));
                            Data_table_centerline.Columns.Add("Z", typeof(Double));

                            for (int i = 0; i < Poly2D.NumberOfVertices; ++i)
                            {
                                Data_table_centerline.Rows.Add();
                                Data_table_centerline.Rows[Data_table_centerline.Rows.Count - 1]["X"] = Poly2D.GetPointAtParameter(i).X;
                                Data_table_centerline.Rows[Data_table_centerline.Rows.Count - 1]["Y"] = Poly2D.GetPointAtParameter(i).Y;
                                Data_table_centerline.Rows[Data_table_centerline.Rows.Count - 1]["Z"] = Poly2D.GetPointAtParameter(i).Z;
                            }

                            double dist2 = dist1 + Match_distance;
                            bool Ultimul = false;
                            int Colorindex = 1;

                            bool Este_primul = true;
                            Point3d Last_pt = new Point3d();

                        l123:
                            Point3d Point1 = new Point3d();
                            Point1 = Poly2D.GetPointAtDist(dist1);
                            Point3d Point2 = new Point3d();
                            Point2 = Poly2D.GetPointAtDist(dist2);


                            Polyline Poly1r = new Autodesk.AutoCAD.DatabaseServices.Polyline();

                            Poly1r = creaza_rectangle_Matchline(Point1, Point2, Colorindex);
                            Poly1r.Layer = Layer_name_ML_rectangle;

                            Point3dCollection Col_int = new Point3dCollection();
                            Col_int = Functions.Intersect_on_both_operands(Poly2D, Poly1r);
                            if (Col_int.Count == 2)
                            {
                                BTrecord.AppendEntity(Poly1r);
                                Trans1.AddNewlyCreatedDBObject(Poly1r, true);
                                Trans1.TransactionManager.QueueForGraphicsFlush();
                                Data_table_matchline.Rows.Add();
                                Data_table_matchline.Rows[Data_table_matchline.Rows.Count - 1]["OBJECT_ID"] = Poly1r.ObjectId.Handle.Value.ToString();
                                Data_table_matchline.Rows[Data_table_matchline.Rows.Count - 1]["M1"] = dist1;
                                Data_table_matchline.Rows[Data_table_matchline.Rows.Count - 1]["M2"] = dist2;
                                Data_table_matchline.Rows[Data_table_matchline.Rows.Count - 1]["X"] = (Poly1r.GetPoint3dAt(0).X + Poly1r.GetPoint3dAt(2).X) / 2;
                                Data_table_matchline.Rows[Data_table_matchline.Rows.Count - 1]["Y"] = (Poly1r.GetPoint3dAt(0).Y + Poly1r.GetPoint3dAt(2).Y) / 2;
                                Data_table_matchline.Rows[Data_table_matchline.Rows.Count - 1]["ROTATION"] = Functions.GET_Bearing_rad(Poly1r.GetPoint3dAt(1).X, Poly1r.GetPoint3dAt(1).Y, Poly1r.GetPoint3dAt(2).X, Poly1r.GetPoint3dAt(2).Y) * 180 / Math.PI;
                                Data_table_matchline.Rows[Data_table_matchline.Rows.Count - 1]["WIDTH"] = Poly1r.GetPoint3dAt(1).DistanceTo(Poly1r.GetPoint3dAt(2));
                                Data_table_matchline.Rows[Data_table_matchline.Rows.Count - 1]["HEIGHT"] = Poly1r.GetPoint3dAt(0).DistanceTo(Poly1r.GetPoint3dAt(1));

                                dist1 = dist2;
                                dist2 = dist2 + Match_distance;

                                Colorindex = Colorindex + 1;
                                if (Colorindex > 7) Colorindex = 1;

                                if (Ultimul == false)
                                {
                                    if (Math.Round(Poly2D.Length, 0) <= Math.Round(dist2, 0))
                                    {
                                        dist2 = Poly2D.Length;
                                        Ultimul = true;
                                    }
                                    Este_primul = true;
                                    goto l123;
                                }
                            }
                            else
                            {
                                Point3d Pointm1 = new Point3d();

                                if (dist1 > 0)
                                {
                                    Autodesk.AutoCAD.EditorInput.PromptPointResult Result_point_m1;
                                    Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1m = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease pick start location:");


                                    if (Este_primul == true)
                                    {
                                        PP1m.AllowNone = false;
                                        Result_point_m1 = Editor1.GetPoint(PP1m);

                                        if (Result_point_m1.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel)
                                        {
                                            Trans1.Commit();
                                            goto end1;
                                        }
                                        Pointm1 = Result_point_m1.Value;
                                        Last_pt = Poly2D.GetClosestPointTo(Pointm1, Vector3d.ZAxis, false);
                                    }
                                }

                                if (dist1 == 0)
                                {
                                    Last_pt = Poly2D.GetPointAtParameter(0);
                                    Pointm1 = Last_pt;
                                }

                            labl1:
                                Jig_rectangle_viewport_SHEET_CUTTER_manual_pt2 Jig2 = new Jig_rectangle_viewport_SHEET_CUTTER_manual_pt2();
                                Autodesk.AutoCAD.EditorInput.PromptPointResult Result_point_m2 = Jig2.StartJig(Vw_scale, Convert.ToDouble(TextBox_matchline_length.Text), Vw_height, Poly2D, Last_pt, 10, Match_distance);

                                if (Result_point_m2.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel)
                                {
                                    Trans1.Commit();
                                    goto end1;
                                }


                                double dist1m;
                                if (Este_primul == true)
                                {
                                    dist1m = Poly2D.GetDistAtPoint(Poly2D.GetClosestPointTo(Pointm1, Vector3d.ZAxis, false));
                                }
                                else
                                {
                                    dist1m = Poly2D.GetDistAtPoint(Poly2D.GetClosestPointTo(Last_pt, Vector3d.ZAxis, false));
                                }

                                Last_pt = Poly2D.GetClosestPointTo(Result_point_m2.Value, Vector3d.ZAxis, false);

                                double dist2m = Poly2D.GetDistAtPoint(Last_pt);


                                if (dist1m > dist2m)
                                {
                                    goto labl1;
                                }

                                Point3d Point1m = new Point3d();
                                Point1m = Poly2D.GetPointAtDist(dist1m);
                                Point3d Point2m = new Point3d();
                                Point2m = Poly2D.GetPointAtDist(dist2m);

                                Polyline Poly1rm = new Autodesk.AutoCAD.DatabaseServices.Polyline();

                                Poly1rm = creaza_rectangle_Matchline(Point1m, Point2m, Colorindex);
                                Poly1rm.Layer = Layer_name_ML_rectangle;

                                BTrecord.AppendEntity(Poly1rm);
                                Trans1.AddNewlyCreatedDBObject(Poly1rm, true);


                                Line Line1 = new Line(Poly1rm.GetPointAtParameter(2), Poly1rm.GetPointAtParameter(3));
                                Line1.TransformBy(Matrix3d.Scaling(10000, Line1.StartPoint));
                                Line1.TransformBy(Matrix3d.Scaling(10000, Line1.EndPoint));

                                Jig_rectangle_viewport_SHEET_CUTTER_manual_up_down Jig1 = new Jig_rectangle_viewport_SHEET_CUTTER_manual_up_down(Point2m, Line1);
                                Jig1.AddEntity(Poly1rm);
                                Autodesk.AutoCAD.EditorInput.PromptResult jigRes = ThisDrawing.Editor.Drag(Jig1);
                                if (jigRes.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                {
                                    Jig1.TransformEntities();
                                }

                                Trans1.TransactionManager.QueueForGraphicsFlush();

                                if (Este_primul == true)
                                {
                                    if (Data_table_matchline.Rows.Count > 0)
                                    {
                                        Double M1_p = (Double)Data_table_matchline.Rows[Data_table_matchline.Rows.Count - 1]["M1"];
                                        Double M2_p = (Double)Data_table_matchline.Rows[Data_table_matchline.Rows.Count - 1]["M2"];
                                        string ob_id = (string)Data_table_matchline.Rows[Data_table_matchline.Rows.Count - 1]["OBJECT_ID"];
                                        ObjectId Oid = Functions.GetObjectId(ThisDrawing.Database, ob_id);
                                        Polyline PolyR = (Polyline)Trans1.GetObject(Oid, OpenMode.ForWrite);

                                        Point3d Point01 = new Point3d();
                                        Point01 = Poly2D.GetPointAtDist(M1_p);
                                        Point3d Point02 = new Point3d();
                                        Point02 = Poly2D.GetPointAtDist(M2_p);
                                        Polyline Poly0r = new Autodesk.AutoCAD.DatabaseServices.Polyline();

                                        Poly0r = creaza_rectangle_Matchline(Point01, Point1m, PolyR.ColorIndex);
                                        Poly0r.Layer = Layer_name_ML_rectangle;

                                        BTrecord.AppendEntity(Poly0r);
                                        Trans1.AddNewlyCreatedDBObject(Poly0r, true);
                                        Trans1.TransactionManager.QueueForGraphicsFlush();

                                        PolyR.Erase();
                                        Trans1.TransactionManager.QueueForGraphicsFlush();

                                        Data_table_matchline.Rows[Data_table_matchline.Rows.Count - 1]["M2"] = dist1m;
                                        Data_table_matchline.Rows[Data_table_matchline.Rows.Count - 1]["OBJECT_ID"] = Poly0r.ObjectId.Handle.Value.ToString();
                                        Data_table_matchline.Rows[Data_table_matchline.Rows.Count - 1]["X"] = (Poly0r.GetPoint3dAt(0).X + Poly0r.GetPoint3dAt(2).X) / 2;
                                        Data_table_matchline.Rows[Data_table_matchline.Rows.Count - 1]["Y"] = (Poly0r.GetPoint3dAt(0).Y + Poly0r.GetPoint3dAt(2).Y) / 2;
                                        Data_table_matchline.Rows[Data_table_matchline.Rows.Count - 1]["ROTATION"] = Functions.GET_Bearing_rad(Poly0r.GetPoint3dAt(1).X, Poly0r.GetPoint3dAt(1).Y, Poly0r.GetPoint3dAt(2).X, Poly0r.GetPoint3dAt(2).Y) * 180 / Math.PI;
                                        Data_table_matchline.Rows[Data_table_matchline.Rows.Count - 1]["WIDTH"] = Poly0r.GetPoint3dAt(1).DistanceTo(Poly0r.GetPoint3dAt(2));
                                        Data_table_matchline.Rows[Data_table_matchline.Rows.Count - 1]["HEIGHT"] = Poly0r.GetPoint3dAt(0).DistanceTo(Poly0r.GetPoint3dAt(1));
                                    }
                                }


                                Data_table_matchline.Rows.Add();
                                Data_table_matchline.Rows[Data_table_matchline.Rows.Count - 1]["M1"] = dist1m;
                                Data_table_matchline.Rows[Data_table_matchline.Rows.Count - 1]["M2"] = dist2m;
                                Data_table_matchline.Rows[Data_table_matchline.Rows.Count - 1]["OBJECT_ID"] = Poly1rm.ObjectId.Handle.Value.ToString();
                                Data_table_matchline.Rows[Data_table_matchline.Rows.Count - 1]["X"] = (Poly1rm.GetPoint3dAt(0).X + Poly1rm.GetPoint3dAt(2).X) / 2;
                                Data_table_matchline.Rows[Data_table_matchline.Rows.Count - 1]["Y"] = (Poly1rm.GetPoint3dAt(0).Y + Poly1rm.GetPoint3dAt(2).Y) / 2;
                                Data_table_matchline.Rows[Data_table_matchline.Rows.Count - 1]["ROTATION"] = Functions.GET_Bearing_rad(Poly1rm.GetPoint3dAt(1).X, Poly1rm.GetPoint3dAt(1).Y, Poly1rm.GetPoint3dAt(2).X, Poly1rm.GetPoint3dAt(2).Y) * 180 / Math.PI;
                                Data_table_matchline.Rows[Data_table_matchline.Rows.Count - 1]["WIDTH"] = Poly1rm.GetPoint3dAt(1).DistanceTo(Poly1rm.GetPoint3dAt(2));
                                Data_table_matchline.Rows[Data_table_matchline.Rows.Count - 1]["HEIGHT"] = Poly1rm.GetPoint3dAt(0).DistanceTo(Poly1rm.GetPoint3dAt(1));

                                Colorindex = Colorindex + 1;
                                if (Colorindex > 7) Colorindex = 1;
                                Este_primul = false;
                                dist1 = dist2m;
                                dist2 = dist2m + Match_distance;
                                if (Math.Round(dist1, 0) == Math.Round(Poly2D.Length, 0))
                                {
                                    goto l124;
                                }
                                if (Math.Round(dist2, 0) > Math.Round(Poly2D.Length, 0))
                                {
                                    dist2 = Poly2D.Length;
                                    Ultimul = true;
                                }
                                goto l123;
                            }
                        l124: Editor1.WriteMessage("\nCommand:");
                            Trans1.Commit();
                        }
                    }
                end1:
                    if (Data_table_matchline != null)
                    {
                        if (Data_table_matchline.Rows.Count > 0)
                        {
                            Populate_data_table_matchline_file_names(0, textBox_prefix_name.Text);
                            dataGridView_sheet_index.DataSource = Data_table_matchline;
                            dataGridView_sheet_index.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                            Append_ML_object_data();
                            if (Data_table_Config_files_path != null)
                            {
                                if (Data_table_Config_files_path.Rows.Count > 0)
                                {
                                    for (int i = 0; i < Data_table_Config_files_path.Rows.Count; ++i)
                                    {
                                        string Type1 = Data_table_Config_files_path.Rows[i][0].ToString();

                                        if (Type1 == "Centerline")
                                        {
                                            Populate_centerline_file(Data_table_Config_files_path.Rows[i][1].ToString());
                                            Update_poly2d_handle_layer_colorindex();
                                        }
                                        if (Type1 == "Sheet Index")
                                        {
                                            Populate_sheet_index_file(Data_table_Config_files_path.Rows[i][1].ToString());
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    Freeze_operations = false;
                    MessageBox.Show(ex.Message);
                }

                Freeze_operations = false;
            }
        }

        private void button_draw_Viewport_templates_Click(object sender, EventArgs e)
        {
            if (System.IO.File.Exists(textBox_config_file_location.Text) == false)
            {
                MessageBox.Show("no config file loaded\r\nOperation aborted");
                return;
            }
            if (Freeze_operations == false)
            {
                Freeze_operations = true;

                Create_VP_object_data();




                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                try
                {

                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {
                        Functions.Creaza_layer(Layer_name_VP_rectangle, 7, false);

                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {

                            BlockTable BlockTable1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                            BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                            Data_table_matchline = Functions.Build_Data_table_matchline_from_object_data();
                            if (Data_table_matchline != null)
                            {
                                if (Data_table_matchline.Rows.Count > 0)
                                {

                                    Data_table_Main_VP = new System.Data.DataTable();
                                    Data_table_Main_VP.Columns.Add("OBJECT_ID", typeof(string));
                                    Data_table_Main_VP.Columns.Add("X", typeof(Double));
                                    Data_table_Main_VP.Columns.Add("Y", typeof(Double));
                                    Data_table_Main_VP.Columns.Add("ROTATION", typeof(Double));
                                    Data_table_Main_VP.Columns.Add("WIDTH", typeof(Double));
                                    Data_table_Main_VP.Columns.Add("HEIGHT", typeof(Double));
                                    Data_table_Main_VP.Columns.Add("FILE_NAME", typeof(string));
                                    for (int i = 0; i < Data_table_matchline.Rows.Count; ++i)
                                    {
                                        Data_table_Main_VP.Rows.Add();

                                    }
                                    for (int i = 0; i < Data_table_matchline.Rows.Count; ++i)
                                    {
                                        double X = (double)Data_table_matchline.Rows[i]["X"];
                                        double Y = (double)Data_table_matchline.Rows[i]["Y"];
                                        double Rotation = (double)Data_table_matchline.Rows[i]["ROTATION"];




                                        string Objectid_string = Data_table_matchline.Rows[i]["OBJECT_ID"].ToString();

                                        ObjectId ObjectID1 = Functions.GetObjectId(ThisDrawing.Database, Objectid_string);

                                        int CI = 256;
                                        if (ObjectID1 != null)
                                        {
                                            Entity Ent1 = (Entity)Trans1.GetObject(ObjectID1, OpenMode.ForRead);
                                            if (Ent1 != null)
                                            {
                                                CI = Ent1.ColorIndex;
                                            }
                                        }
                                        Polyline Poly1 = creaza_rectangle_from_one_point(new Point3d(X, Y, 0), Rotation * Math.PI / 180, Vw_width, Vw_height, CI);
                                        Poly1.Layer = Layer_name_VP_rectangle;
                                        BTrecord.AppendEntity(Poly1);
                                        Trans1.AddNewlyCreatedDBObject(Poly1, true);

                                        Data_table_Main_VP.Rows[i]["OBJECT_ID"] = Poly1.ObjectId.Handle.Value.ToString();
                                        Data_table_Main_VP.Rows[i]["X"] = X;
                                        Data_table_Main_VP.Rows[i]["Y"] = Y;
                                        Data_table_Main_VP.Rows[i]["ROTATION"] = Rotation;
                                        Data_table_Main_VP.Rows[i]["WIDTH"] = Vw_width;
                                        Data_table_Main_VP.Rows[i]["HEIGHT"] = Vw_height;
                                        Data_table_Main_VP.Rows[i]["FILE_NAME"] = Data_table_matchline.Rows[i]["FILE_NAME"].ToString();

                                    }
                                }
                            }











                            Editor1.WriteMessage("\nCommand:");

                            Trans1.Commit();
                        }
                    }




                    if (Data_table_Main_VP != null)
                    {
                        if (Data_table_Main_VP.Rows.Count > 0)
                        {
                            Append_VP_object_data();


                        }
                    }
                }
                catch (System.Exception ex)
                {
                    Freeze_operations = false;
                    MessageBox.Show(ex.Message);
                }

                Freeze_operations = false;
            }
        }

        private void Populate_sheet_index_file(String File1)
        {
            try
            {
                if (System.IO.File.Exists(File1) == false)
                {
                    MessageBox.Show("the sheet index file does not exists\r\n" + File1);
                    return;
                }

                Microsoft.Office.Interop.Excel.Application Excel1 = new Microsoft.Office.Interop.Excel.Application();
                if (Excel1 == null)
                {
                    MessageBox.Show("PROBLEM WITH EXCEL!");
                    return;
                }
                Excel1.Visible = true;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(File1);
                Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];

                try
                {
                    Functions.Transfer_to_w1_Data_table(W1, Data_table_matchline,3);



                    Workbook1.Save();
                    Workbook1.Close();
                    Excel1.Quit();
                }
                catch (System.Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);

                }
                finally
                {
                    if (W1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W1);
                    if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                    if (Excel1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }

        }

        private void Populate_centerline_file(String File1)
        {
            try
            {
                if (System.IO.File.Exists(File1) == false)
                {
                    MessageBox.Show("the centerline file does not exists\r\n" + File1);
                    return;
                }

                Microsoft.Office.Interop.Excel.Application Excel1 = new Microsoft.Office.Interop.Excel.Application();
                if (Excel1 == null)
                {
                    MessageBox.Show("PROBLEM WITH EXCEL!");
                    return;
                }
                Excel1.Visible = true;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(File1);
                Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];

                try
                {
                    Functions.Transfer_to_w1_Data_table(W1, Data_table_centerline,3);



                    Workbook1.Save();
                    Workbook1.Close();
                    Excel1.Quit();
                }
                catch (System.Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);

                }
                finally
                {
                    if (W1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W1);
                    if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                    if (Excel1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }

        }

        private void Update_poly2d_handle_layer_colorindex()
        {
            try
            {
                if (System.IO.File.Exists(textBox_config_file_location.Text) == false)
                {
                    MessageBox.Show("the sheet index file does not exists\r\n" + textBox_config_file_location.Text);
                    return;
                }

                Microsoft.Office.Interop.Excel.Application Excel1 = new Microsoft.Office.Interop.Excel.Application();
                if (Excel1 == null)
                {
                    MessageBox.Show("PROBLEM WITH EXCEL!");
                    return;
                }
                Excel1.Visible = true;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(textBox_config_file_location.Text);
                Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];

                try
                {
                    W1.Range["B21"].Value2 = Layer_name_Poly2D;
                    W1.Range["B22"].Value2 = Color_index_Layer_name_Poly2D;
                    W1.Range["B23"].Value2 = Poly2D_handle;


                    Workbook1.Save();
                    Workbook1.Close();
                    Excel1.Quit();
                }
                catch (System.Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);

                }
                finally
                {
                    if (W1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W1);
                    if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                    if (Excel1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }

        }

        private void Populate_data_table_matchline_file_names(int Start_line, string Old_name)
        {
            if (Data_table_matchline != null)
            {
                if (Data_table_matchline.Rows.Count > Start_line)
                {
                    string Suffix = "";
                    if (Start_line == 0)
                    {
                        Suffix = textBox_name_start_number.Text;
                    }

                    string New_name = Old_name + Suffix;
                    for (int i = Start_line; i < Data_table_matchline.Rows.Count; ++i)
                    {

                        int Increment = 1;
                        if (Functions.IsNumeric(textBox_name_increment.Text) == true)
                        {
                            Increment = Convert.ToInt32(textBox_name_increment.Text);
                        }

                        if (i != 0)
                        {
                            New_name = Functions.get_new_file_name_incremented(Increment, New_name);
                        }
                        Data_table_matchline.Rows[i]["FILE_NAME"] = New_name;
                    }
                }
            }


        }

        private Polyline creaza_rectangle_Matchline(Point3d Point1, Point3d Point2, int cid)
        {

            Autodesk.AutoCAD.DatabaseServices.Line Line1R = new Autodesk.AutoCAD.DatabaseServices.Line(Point1, Point2);
            Point3d Point_distR = new Point3d();
            if (Line1R.Length > Vw_height / Vw_scale)
            {
                Point_distR = Line1R.GetPointAtDist(Vw_height / Vw_scale);
                Line1R.EndPoint = Point_distR;
            }
            else
            {
                Line1R.TransformBy(Matrix3d.Scaling((Vw_height / Vw_scale) / Line1R.Length, Line1R.StartPoint));
            }

            Line1R.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, Point1));
            Point3d Point_middler = new Point3d((Point1.X + Line1R.EndPoint.X) / 2, (Point1.Y + Line1R.EndPoint.Y) / 2, 0);

            Line1R.TransformBy(Matrix3d.Displacement(Point_middler.GetVectorTo(Point1)));
            Point3d Pt1r = new Point3d();
            Pt1r = Line1R.StartPoint;
            Point3d Pt2r = new Point3d();
            Pt2r = Line1R.EndPoint;
            Line1R.TransformBy(Matrix3d.Displacement(Point1.GetVectorTo(Point2)));

            Point3d Pt4r = new Point3d();
            Pt4r = Line1R.StartPoint;
            Point3d Pt3r = new Point3d();
            Pt3r = Line1R.EndPoint;


            Polyline Poly1r = new Autodesk.AutoCAD.DatabaseServices.Polyline();
            Poly1r.AddVertexAt(0, new Point2d(Pt1r.X, Pt1r.Y), 0, 0, 0);
            Poly1r.AddVertexAt(1, new Point2d(Pt2r.X, Pt2r.Y), 0, 0, 0);
            Poly1r.AddVertexAt(2, new Point2d(Pt3r.X, Pt3r.Y), 0, 0, 0);
            Poly1r.AddVertexAt(3, new Point2d(Pt4r.X, Pt4r.Y), 0, 0, 0);
            Poly1r.Closed = true;
            Poly1r.ColorIndex = cid;
            return Poly1r;
        }

        private Polyline creaza_rectangle_from_one_point(Point3d Point1, double Rotation_rad, double Width1, double Height1, int cid)
        {
            Polyline Poly1r = new Autodesk.AutoCAD.DatabaseServices.Polyline();
            Poly1r.AddVertexAt(0, new Point2d(Point1.X - Width1 / 2, Point1.Y - Height1 / 2), 0, 0, 0);
            Poly1r.AddVertexAt(1, new Point2d(Point1.X - Width1 / 2, Point1.Y + Height1 / 2), 0, 0, 0);
            Poly1r.AddVertexAt(2, new Point2d(Point1.X + Width1 / 2, Point1.Y + Height1 / 2), 0, 0, 0);
            Poly1r.AddVertexAt(3, new Point2d(Point1.X + Width1 / 2, Point1.Y - Height1 / 2), 0, 0, 0);


            Poly1r.Closed = true;
            Poly1r.ColorIndex = cid;

            Poly1r.TransformBy(Matrix3d.Rotation(Rotation_rad, Vector3d.ZAxis, Point1));

            return Poly1r;
        }

        private void button_adjust_rectangle_Click(object sender, EventArgs e)
        {
            if (System.IO.File.Exists(textBox_config_file_location.Text) == false)
            {
                MessageBox.Show("no config file loaded\r\nOperation aborted");
                return;
            }
            if (Freeze_operations == false)
            {
                Freeze_operations = true;
                Erase_viewports_templates();
                ObjectId[] Empty_array = null;
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                try
                {

                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {

                            string Path_toCL = "";
                            if (Data_table_Config_files_path != null)
                            {
                                if (Data_table_Config_files_path.Rows.Count > 0)
                                {
                                    for (int i = 0; i < Data_table_Config_files_path.Rows.Count; ++i)
                                    {
                                        string Type1 = Data_table_Config_files_path.Rows[i][0].ToString();

                                        if (Type1 == "Centerline")
                                        {
                                            Path_toCL = Data_table_Config_files_path.Rows[i][1].ToString();
                                        }

                                    }
                                }
                            }

                            if (Path_toCL == "")
                            {
                                Freeze_operations = false;
                                MessageBox.Show("No centerline file loaded");
                                return;
                            }

                            if (System.IO.File.Exists(Path_toCL) == false)
                            {
                                Freeze_operations = false;
                                MessageBox.Show("No centerline file not found");
                                return;
                            }


                            ObjectId Ob1 = Functions.GetObjectId(ThisDrawing.Database, Poly2D_handle);
                            Poly2D = (Polyline)Trans1.GetObject(Ob1, OpenMode.ForRead);
                            if (Poly2D == null)
                            {

                                Freeze_operations = false;
                                Editor1.SetImpliedSelection(Empty_array);
                                Editor1.WriteMessage("\nCommand:");
                                MessageBox.Show("there is no centerline into the current drawing");
                                return;

                            }

                            Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_optionsrec = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect rectangle:");
                            Prompt_optionsrec.SetRejectMessage("\nYou did not selected a polyline");
                            Prompt_optionsrec.AddAllowedClass(typeof(Polyline), true);

                            Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_rec = Editor1.GetEntity(Prompt_optionsrec);
                            if (Rezultat_rec.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                            {
                                Freeze_operations = false;
                                Editor1.SetImpliedSelection(Empty_array);
                                Editor1.WriteMessage("\nCommand:");
                                return;
                            }

                            Polyline Rect_0 = (Polyline)Trans1.GetObject(Rezultat_rec.ObjectId, OpenMode.ForWrite);

                            Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable1 = (Autodesk.AutoCAD.DatabaseServices.BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                            //Dim BTrecord_MS As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(BlockTable1(BlockTableRecord.ModelSpace), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (Autodesk.AutoCAD.DatabaseServices.BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                            ObjectId Obj_id_old = Rect_0.ObjectId;


                            Data_table_matchline = Functions.Build_Data_table_matchline_from_object_data();



                            if (Data_table_matchline != null)
                            {
                                if (Data_table_matchline.Rows.Count > 0)
                                {
                                    int Index0 = -1;
                                    for (int i = 0; i < Data_table_matchline.Rows.Count; ++i)
                                    {
                                        if (Obj_id_old == Functions.GetObjectId(ThisDrawing.Database, (string)Data_table_matchline.Rows[i]["OBJECT_ID"]))
                                        {
                                            Index0 = i;
                                            i = Data_table_matchline.Rows.Count;
                                        }
                                    }

                                    if (Index0 != -1)
                                    {

                                        Line Line1 = new Line(Rect_0.GetPointAtParameter(2), Rect_0.GetPointAtParameter(3));
                                        Line1.TransformBy(Matrix3d.Scaling(10000, Line1.StartPoint));
                                        Line1.TransformBy(Matrix3d.Scaling(10000, Line1.EndPoint));
                                        double M2 = (double)Data_table_matchline.Rows[Index0]["M2"];
                                        Point3dCollection Col1 = Functions.Intersect_with_extend(Rect_0, Poly2D);
                                        if (Col1.Count == 0)
                                        {
                                            MessageBox.Show("The rectangle does not intersect the centerline....");
                                            Freeze_operations = false;
                                            return;
                                        }

                                        for (int i = 0; i < Col1.Count; ++i)
                                        {
                                            try
                                            {
                                                if (Math.Round(Poly2D.GetDistAtPoint(Col1[i]), 0) == M2)
                                                {
                                                    M2 = Poly2D.GetDistAtPoint(Col1[i]);
                                                    i = Col1.Count;
                                                }
                                            }
                                            catch (System.Exception ex)
                                            {
                                                MessageBox.Show("The rectangle does not intersect the centerline....");
                                                Freeze_operations = false;
                                                return;
                                            }

                                        }

                                        Jig_rectangle_viewport_SHEET_CUTTER_manual_up_down Jig1 = new Jig_rectangle_viewport_SHEET_CUTTER_manual_up_down(Poly2D.GetPointAtDist(M2), Line1);
                                        Jig1.AddEntity(Rect_0);
                                        Autodesk.AutoCAD.EditorInput.PromptResult jigRes = ThisDrawing.Editor.Drag(Jig1);
                                        if (jigRes.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                        {
                                            Jig1.TransformEntities();


                                            Data_table_matchline.Rows[Index0]["X"] = (Rect_0.GetPoint3dAt(0).X + Rect_0.GetPoint3dAt(2).X) / 2;
                                            Data_table_matchline.Rows[Index0]["Y"] = (Rect_0.GetPoint3dAt(0).Y + Rect_0.GetPoint3dAt(2).Y) / 2;
                                        }

                                        Trans1.TransactionManager.QueueForGraphicsFlush();


                                    }
                                }
                            }

                            Trans1.Commit();
                            Editor1.WriteMessage("\nCommand:");
                        }

                    }

                    if (Data_table_matchline != null)
                    {
                        if (Data_table_matchline.Rows.Count > 0)
                        {
                            Append_ML_object_data();
                            dataGridView_sheet_index.DataSource = Data_table_matchline;
                            dataGridView_sheet_index.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

                            if (Data_table_Config_files_path != null)
                            {
                                if (Data_table_Config_files_path.Rows.Count > 0)
                                {
                                    for (int i = 0; i < Data_table_Config_files_path.Rows.Count; ++i)
                                    {
                                        string Type1 = Data_table_Config_files_path.Rows[i][0].ToString();

                                        if (Type1 == "Centerline")
                                        {
                                            Populate_centerline_file(Data_table_Config_files_path.Rows[i][1].ToString());
                                        }
                                        if (Type1 == "Sheet Index")
                                        {
                                            Populate_sheet_index_file(Data_table_Config_files_path.Rows[i][1].ToString());
                                        }
                                    }
                                }
                            }


                        }
                    }
                }
                catch (System.Exception ex)
                {
                    Freeze_operations = false;
                    MessageBox.Show(ex.Message);
                }
                Freeze_operations = false;
            }

        }

        private void button_Fill_gaps_Click(object sender, EventArgs e)
        {
            if (System.IO.File.Exists(textBox_config_file_location.Text) == false)
            {
                MessageBox.Show("no config file loaded\r\nOperation aborted");
                return;
            }


            if (Freeze_operations == false)
            {
                Freeze_operations = true;

                Erase_viewports_templates();

                if (Functions.IsNumeric(TextBox_matchline_length.Text) == true)
                {
                    Match_distance = Convert.ToDouble(TextBox_matchline_length.Text);
                }




                ObjectId[] Empty_array = null;

                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                try
                {


                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {
                        Functions.Creaza_layer(Layer_name_ML_rectangle, 4, false);

                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            string Path_toCL = "";
                            if (Data_table_Config_files_path != null)
                            {
                                if (Data_table_Config_files_path.Rows.Count > 0)
                                {
                                    for (int i = 0; i < Data_table_Config_files_path.Rows.Count; ++i)
                                    {
                                        string Type1 = Data_table_Config_files_path.Rows[i][0].ToString();

                                        if (Type1 == "Centerline")
                                        {
                                            Path_toCL = Data_table_Config_files_path.Rows[i][1].ToString();
                                        }

                                    }
                                }
                            }

                            if (Path_toCL == "")
                            {
                                Freeze_operations = false;
                                MessageBox.Show("No centerline file loaded");
                                return;
                            }

                            if (System.IO.File.Exists(Path_toCL) == false)
                            {
                                Freeze_operations = false;
                                MessageBox.Show("No centerline file not found");
                                return;
                            }



                            ObjectId Ob1 = Functions.GetObjectId(ThisDrawing.Database, Poly2D_handle);
                            Poly2D = (Polyline)Trans1.GetObject(Ob1, OpenMode.ForRead);
                            if (Poly2D == null)
                            {

                                Freeze_operations = false;
                                Editor1.SetImpliedSelection(Empty_array);
                                Editor1.WriteMessage("\nCommand:");
                                MessageBox.Show("there is no centerline into the current drawing");
                                return;

                            }


                            BlockTable BlockTable1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                            //Dim BTrecord_MS As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(BlockTable1(BlockTableRecord.ModelSpace), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                            BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);



                            Data_table_matchline = Functions.Build_Data_table_matchline_from_object_data();

                            List<int> List1 = new List<int>();



                            if (Data_table_matchline.Rows.Count > 0)
                            {
                                for (int i = 0; i < Data_table_matchline.Rows.Count; ++i)
                                {
                                    string OdId = Data_table_matchline.Rows[i]["OBJECT_ID"].ToString();

                                    try
                                    {
                                        ObjectId ObjID1 = Functions.GetObjectId(ThisDrawing.Database, OdId);
                                        DBObject Obj1 = Trans1.GetObject(ObjID1, OpenMode.ForRead);
                                        if (Obj1 is Polyline)
                                        {
                                            Polyline Rec1 = (Polyline)Obj1;
                                        }
                                    }
                                    catch (System.Exception Ex)
                                    {
                                        List1.Add(i);
                                    }
                                }



                                if (List1.Count > 0)
                                {
                                    for (int i = List1.Count - 1; i >= 0; --i)
                                    {

                                        Data_table_matchline.Rows[List1[i]].Delete();


                                    }
                                }

                            }


                            List<int> List2 = new List<int>();

                            double prev_m = 0;
                            for (int i = 0; i < Data_table_matchline.Rows.Count; ++i)
                            {
                                string OdId = Data_table_matchline.Rows[i]["OBJECT_ID"].ToString();
                                double m1 = (double)Data_table_matchline.Rows[i]["M1"];
                                double m2 = (double)Data_table_matchline.Rows[i]["M2"];
                                if (m1 != prev_m)
                                {
                                    List2.Add(i);
                                }
                                prev_m = m2;

                            }





                            int Colorindex = 1;


                            Point3d Last_pt = new Point3d();

                            double dist1 = 0;
                            double dist2 = 0;
                            Double Mnext = 0;




                            for (int i = 0; i < List2.Count; ++i)
                            {

                                Mnext = (double)Data_table_matchline.Rows[List2[i]]["M1"];

                                if (List2[i] != 0)
                                {
                                    dist1 = (double)Data_table_matchline.Rows[List2[i] - 1]["M2"];
                                }


                            l1234:
                                Last_pt = Poly2D.GetPointAtDist(dist1);

                                Jig_rectangle_viewport_SHEET_CUTTER_manual_pt2 Jig2m = new Jig_rectangle_viewport_SHEET_CUTTER_manual_pt2();
                                Autodesk.AutoCAD.EditorInput.PromptPointResult Result_point_m2m = Jig2m.StartJig(Vw_scale, Vw_width, Vw_height, Poly2D, Last_pt, 10, Match_distance);

                                if (Result_point_m2m.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel)
                                {
                                    Trans1.Commit();
                                    goto end1;
                                }

                                Last_pt = Poly2D.GetClosestPointTo(Result_point_m2m.Value, Vector3d.ZAxis, false);

                                dist2 = Poly2D.GetDistAtPoint(Last_pt);

                                if (Math.Round(dist1, 0) > Math.Round(dist2, 0))
                                {
                                    goto l1234;
                                }

                                Point3d Point1m = new Point3d();
                                Point1m = Poly2D.GetPointAtDist(dist1);
                                Point3d Point2m = new Point3d();
                                Point2m = Poly2D.GetPointAtDist(dist2);

                                Polyline Poly1rm = new Autodesk.AutoCAD.DatabaseServices.Polyline();

                                Poly1rm = creaza_rectangle_Matchline(Point1m, Point2m, Colorindex);
                                Poly1rm.Layer = Layer_name_ML_rectangle;

                                BTrecord.AppendEntity(Poly1rm);
                                Trans1.AddNewlyCreatedDBObject(Poly1rm, true);


                                Line Line1 = new Line(Poly1rm.GetPointAtParameter(2), Poly1rm.GetPointAtParameter(3));
                                Line1.TransformBy(Matrix3d.Scaling(10000, Line1.StartPoint));
                                Line1.TransformBy(Matrix3d.Scaling(10000, Line1.EndPoint));

                                Jig_rectangle_viewport_SHEET_CUTTER_manual_up_down Jig1 = new Jig_rectangle_viewport_SHEET_CUTTER_manual_up_down(Point2m, Line1);
                                Jig1.AddEntity(Poly1rm);
                                Autodesk.AutoCAD.EditorInput.PromptResult jigRes = ThisDrawing.Editor.Drag(Jig1);
                                if (jigRes.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                {
                                    Jig1.TransformEntities();
                                }

                                Trans1.TransactionManager.QueueForGraphicsFlush();


                                System.Data.DataRow Row1 = Data_table_matchline.NewRow();

                                Row1["M1"] = dist1;
                                Row1["M2"] = dist2;
                                Row1["OBJECT_ID"] = Poly1rm.ObjectId.Handle.Value.ToString();
                                Row1["X"] = (Poly1rm.GetPoint3dAt(0).X + Poly1rm.GetPoint3dAt(2).X) / 2;
                                Row1["Y"] = (Poly1rm.GetPoint3dAt(0).Y + Poly1rm.GetPoint3dAt(2).Y) / 2;
                                Row1["ROTATION"] = Functions.GET_Bearing_rad(Poly1rm.GetPoint3dAt(1).X, Poly1rm.GetPoint3dAt(1).Y, Poly1rm.GetPoint3dAt(2).X, Poly1rm.GetPoint3dAt(2).Y) * 180 / Math.PI;
                                Row1["WIDTH"] = Poly1rm.GetPoint3dAt(1).DistanceTo(Poly1rm.GetPoint3dAt(2));
                                Row1["HEIGHT"] = Poly1rm.GetPoint3dAt(0).DistanceTo(Poly1rm.GetPoint3dAt(1));



                                Data_table_matchline.Rows.InsertAt(Row1, List2[i]);




                                for (int j = i; j < List2.Count; ++j)
                                {
                                    List2[j] = List2[j] + 1;
                                }


                                Colorindex = Colorindex + 1;
                                if (Colorindex > 7) Colorindex = 1;


                                dist1 = dist2;

                                if (Math.Round(dist2, 0) < Math.Round(Mnext, 0))
                                {
                                    goto l1234;
                                }
                                else
                                {
                                    Point3d Point1mn = new Point3d();
                                    Point1mn = Poly2D.GetPointAtDist(dist1);
                                    double dist2mn = (double)Data_table_matchline.Rows[List2[i]]["M2"];
                                    Point3d Point2mn = new Point3d();
                                    Point2mn = Poly2D.GetPointAtDist(dist2mn);
                                    Polyline Poly1rmn = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                                    Poly1rmn = creaza_rectangle_Matchline(Point1mn, Point2mn, Colorindex);
                                    Poly1rmn.Layer = Layer_name_ML_rectangle;
                                    BTrecord.AppendEntity(Poly1rmn);
                                    Trans1.AddNewlyCreatedDBObject(Poly1rmn, true);

                                    string OBid1 = Data_table_matchline.Rows[List2[i]]["OBJECT_ID"].ToString();
                                    ObjectId obid = Functions.GetObjectId(ThisDrawing.Database, OBid1);
                                    DBObject Dbobj1 = Trans1.GetObject(obid, OpenMode.ForWrite);
                                    Dbobj1.Erase();

                                    Data_table_matchline.Rows[List2[i]]["M1"] = dist1;
                                    Data_table_matchline.Rows[List2[i]]["OBJECT_ID"] = Poly1rmn.ObjectId.Handle.Value.ToString();
                                    Data_table_matchline.Rows[List2[i]]["X"] = (Poly1rmn.GetPoint3dAt(0).X + Poly1rmn.GetPoint3dAt(2).X) / 2;
                                    Data_table_matchline.Rows[List2[i]]["Y"] = (Poly1rmn.GetPoint3dAt(0).Y + Poly1rmn.GetPoint3dAt(2).Y) / 2;
                                    Data_table_matchline.Rows[List2[i]]["ROTATION"] = Functions.GET_Bearing_rad(Poly1rmn.GetPoint3dAt(1).X, Poly1rmn.GetPoint3dAt(1).Y, Poly1rmn.GetPoint3dAt(2).X, Poly1rmn.GetPoint3dAt(2).Y) * 180 / Math.PI;
                                    Data_table_matchline.Rows[List2[i]]["WIDTH"] = Poly1rmn.GetPoint3dAt(1).DistanceTo(Poly1rmn.GetPoint3dAt(2));
                                    Data_table_matchline.Rows[List2[i]]["HEIGHT"] = Poly1rmn.GetPoint3dAt(0).DistanceTo(Poly1rmn.GetPoint3dAt(1));
                                    Trans1.TransactionManager.QueueForGraphicsFlush();

                                }


                            }




                            Double Lastdist1 = 0;

                            if (Data_table_matchline.Rows.Count > 0)
                            {
                                Lastdist1 = (double)Data_table_matchline.Rows[Data_table_matchline.Rows.Count - 1]["M2"];
                            }

                            if (Math.Round(Poly2D.Length, 0) > Math.Round(Lastdist1, 0))
                            {
                                Last_pt = Poly2D.GetPointAtDist(Lastdist1);

                            l1235:
                                Jig_rectangle_viewport_SHEET_CUTTER_manual_pt2 Jig2m = new Jig_rectangle_viewport_SHEET_CUTTER_manual_pt2();
                                Autodesk.AutoCAD.EditorInput.PromptPointResult Result_point_m2m = Jig2m.StartJig(Vw_scale, Vw_width, Vw_height, Poly2D, Last_pt, 10, Match_distance);

                                if (Result_point_m2m.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel)
                                {
                                    Trans1.Commit();
                                    goto end1;
                                }

                                Last_pt = Poly2D.GetClosestPointTo(Result_point_m2m.Value, Vector3d.ZAxis, false);

                                Double Lastdist2 = Poly2D.GetDistAtPoint(Last_pt);

                                if (Math.Round(Lastdist1, 0) > Math.Round(Lastdist2, 0))
                                {
                                    goto l1235;
                                }

                                Point3d Point1m = new Point3d();
                                Point1m = Poly2D.GetPointAtDist(Lastdist1);
                                Point3d Point2m = new Point3d();
                                Point2m = Poly2D.GetPointAtDist(Lastdist2);

                                Polyline Poly1rm = new Autodesk.AutoCAD.DatabaseServices.Polyline();

                                Poly1rm = creaza_rectangle_Matchline(Point1m, Point2m, Colorindex);
                                Poly1rm.Layer = Layer_name_ML_rectangle;

                                BTrecord.AppendEntity(Poly1rm);
                                Trans1.AddNewlyCreatedDBObject(Poly1rm, true);


                                Line Line1 = new Line(Poly1rm.GetPointAtParameter(2), Poly1rm.GetPointAtParameter(3));
                                Line1.TransformBy(Matrix3d.Scaling(10000, Line1.StartPoint));
                                Line1.TransformBy(Matrix3d.Scaling(10000, Line1.EndPoint));

                                Jig_rectangle_viewport_SHEET_CUTTER_manual_up_down Jig1 = new Jig_rectangle_viewport_SHEET_CUTTER_manual_up_down(Point2m, Line1);
                                Jig1.AddEntity(Poly1rm);
                                Autodesk.AutoCAD.EditorInput.PromptResult jigRes = ThisDrawing.Editor.Drag(Jig1);
                                if (jigRes.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                {
                                    Jig1.TransformEntities();
                                }

                                Trans1.TransactionManager.QueueForGraphicsFlush();




                                Data_table_matchline.Rows.Add();
                                Data_table_matchline.Rows[Data_table_matchline.Rows.Count - 1]["M1"] = Lastdist1;
                                Data_table_matchline.Rows[Data_table_matchline.Rows.Count - 1]["M2"] = Lastdist2;
                                Data_table_matchline.Rows[Data_table_matchline.Rows.Count - 1]["OBJECT_ID"] = Poly1rm.ObjectId.Handle.Value.ToString();
                                Data_table_matchline.Rows[Data_table_matchline.Rows.Count - 1]["X"] = (Poly1rm.GetPoint3dAt(0).X + Poly1rm.GetPoint3dAt(2).X) / 2;
                                Data_table_matchline.Rows[Data_table_matchline.Rows.Count - 1]["Y"] = (Poly1rm.GetPoint3dAt(0).Y + Poly1rm.GetPoint3dAt(2).Y) / 2;
                                Data_table_matchline.Rows[Data_table_matchline.Rows.Count - 1]["ROTATION"] = Functions.GET_Bearing_rad(Poly1rm.GetPoint3dAt(1).X, Poly1rm.GetPoint3dAt(1).Y, Poly1rm.GetPoint3dAt(2).X, Poly1rm.GetPoint3dAt(2).Y) * 180 / Math.PI;
                                Data_table_matchline.Rows[Data_table_matchline.Rows.Count - 1]["WIDTH"] = Poly1rm.GetPoint3dAt(1).DistanceTo(Poly1rm.GetPoint3dAt(2));
                                Data_table_matchline.Rows[Data_table_matchline.Rows.Count - 1]["HEIGHT"] = Poly1rm.GetPoint3dAt(0).DistanceTo(Poly1rm.GetPoint3dAt(1));





                                Colorindex = Colorindex + 1;
                                if (Colorindex > 7) Colorindex = 1;
                                if (Math.Round(Lastdist2, 0) == Math.Round(Poly2D.Length, 0))
                                {
                                    Lastdist2 = Poly2D.Length;
                                }


                                Lastdist1 = Lastdist2;

                                if (Math.Round(Lastdist2, 0) < Math.Round(Poly2D.Length, 0))
                                {
                                    goto l1235;
                                }



                            }


                            Editor1.WriteMessage("\nCommand:");

                            Trans1.Commit();



                        }
                    }


                end1:

                    if (Data_table_matchline != null)
                    {
                        if (Data_table_matchline.Rows.Count > 0)
                        {
                            Populate_data_table_matchline_file_names(0, textBox_prefix_name.Text);
                            dataGridView_sheet_index.DataSource = Data_table_matchline;
                            dataGridView_sheet_index.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

                            Append_ML_object_data();

                            if (Data_table_Config_files_path != null)
                            {
                                if (Data_table_Config_files_path.Rows.Count > 0)
                                {
                                    for (int i = 0; i < Data_table_Config_files_path.Rows.Count; ++i)
                                    {
                                        string Type1 = Data_table_Config_files_path.Rows[i][0].ToString();

                                        if (Type1 == "Centerline")
                                        {
                                            Populate_centerline_file(Data_table_Config_files_path.Rows[i][1].ToString());
                                        }
                                        if (Type1 == "Sheet Index")
                                        {
                                            Populate_sheet_index_file(Data_table_Config_files_path.Rows[i][1].ToString());
                                        }
                                    }
                                }
                            }

                        }
                    }
                }
                catch (System.Exception ex)
                {
                    Freeze_operations = false;
                    MessageBox.Show(ex.Message);
                }

                Freeze_operations = false;
            }
        }

        private void Erase_viewports_templates()
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
                            Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                            foreach (ObjectId Odid in BTrecord)
                            {
                                Entity Ent1 = (Entity)Trans1.GetObject(Odid, OpenMode.ForRead);
                                if (Ent1 != null)
                                {
                                    if (Ent1 is Polyline && Ent1.Layer == Layer_name_VP_rectangle)
                                    {
                                        Ent1.UpgradeOpen();
                                        Ent1.Erase();
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
            }

        }

        private void button_generate_sheets_Click(object sender, EventArgs e)
        {
            try
            {
                if (System.IO.File.Exists(textBox_config_file_location.Text) == false)
                {
                    MessageBox.Show("no config file loaded\r\nOperation aborted");
                    return;
                }

                if (Freeze_operations == false)
                {
                    Freeze_operations = true;
                    Data_table_matchline = Functions.Build_Data_table_matchline_from_object_data();




                    string Template_file_name = textBox_template_name.Text;
                    string Output_folder = textBox_output_folder.Text;


                    Point3d MSpoint;
                    Point3d PSpoint = new Point3d(Vw_ps_x, Vw_ps_y, 0);

                    Double Twist = 0;



                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            DocumentCollection DocumentManager1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
                            Document New_doc = DocumentCollectionExtension.Add(DocumentManager1, Template_file_name);
                            DocumentManager1.MdiActiveDocument = New_doc;
                            using (DocumentLock lock2 = New_doc.LockDocument())
                            {

                                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans2 = New_doc.Database.TransactionManager.StartTransaction())
                                {
                                    BlockTableRecord BtrecordPS = Functions.get_first_layout_as_paperspace(Trans2, New_doc.Database);
                                    BtrecordPS.UpgradeOpen();
                                    Layout Layout1 = Functions.get_first_layout(Trans2, New_doc.Database);
                                    Layout1.UpgradeOpen();
                                    Layout1.LayoutName = Data_table_matchline.Rows[0]["FILE_NAME"].ToString();
                                    Trans2.Commit();
                                    New_doc.Database.SaveAs(Output_folder + Data_table_matchline.Rows[0]["FILE_NAME"].ToString() + ".dwg", true, DwgVersion.Current, ThisDrawing.Database.SecurityParameters);

                                }

                            }
                            New_doc.CloseAndDiscard();

                            if (Data_table_matchline.Rows.Count > 1)
                            {
                                for (int i = 1; i < Data_table_matchline.Rows.Count; ++i)
                                {

                                    string Fisier1 = Output_folder + Data_table_matchline.Rows[i - 1]["FILE_NAME"].ToString() + ".dwg";
                                    string Fisier2 = Output_folder + Data_table_matchline.Rows[i]["FILE_NAME"].ToString() + ".dwg";
                                    System.IO.File.Copy(Fisier1, Fisier2, false);

                                }
                            }

                            for (int i = 0; i < Data_table_matchline.Rows.Count; ++i)
                            {

                                string Fisier = Output_folder + Data_table_matchline.Rows[i]["FILE_NAME"].ToString() + ".dwg";
                                using (Database Database2 = new Database(false, true))
                                {
                                    HostApplicationServices.WorkingDatabase = Database2;
                                    Database2.ReadDwgFile(Fisier, System.IO.FileShare.ReadWrite, false, null);

                                    Functions.Creaza_layer_on_database(Database2, Layer_name_Main_Viewport, 4, false);
                                    Functions.Creaza_layer_on_database(Database2, Layer_North_Arrow, 7, true);



                                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans2 = Database2.TransactionManager.StartTransaction())
                                    {
                                        BlockTableRecord BtrecordPS = Functions.get_first_layout_as_paperspace(Trans2, Database2);
                                        BtrecordPS.UpgradeOpen();
                                        Layout Layout1 = Functions.get_first_layout(Trans2, Database2);
                                        Layout1.UpgradeOpen();
                                        Layout1.LayoutName = Data_table_matchline.Rows[i]["FILE_NAME"].ToString();
                                        MSpoint = new Point3d((double)Data_table_matchline.Rows[i]["X"], (double)Data_table_matchline.Rows[i]["Y"], 0);
                                        Twist = 2 * Math.PI - (double)Data_table_matchline.Rows[i]["ROTATION"] * Math.PI / 180;

                                        Viewport Viewport1 = Functions.Create_viewport(MSpoint, PSpoint, Vw_width, Vw_height, Vw_scale, Twist);
                                        Viewport1.Layer = Layer_name_Main_Viewport;
                                        BtrecordPS.AppendEntity(Viewport1);
                                        Trans2.AddNewlyCreatedDBObject(Viewport1, true);

                                        BlockReference North_arrow = Functions.InsertBlock_with_multiple_atributes_with_database(Database2, BtrecordPS,
                                            "", NA_name, new Point3d(NA_x, NA_y, 0),
                                            NA_scale, Twist, Layer_North_Arrow, new System.Collections.Specialized.StringCollection(), new System.Collections.Specialized.StringCollection());


                                        Trans2.Commit();
                                        Database2.SaveAs(Fisier, true, DwgVersion.Current, ThisDrawing.Database.SecurityParameters);
                                    }
                                }
                            }
                            HostApplicationServices.WorkingDatabase = ThisDrawing.Database;
                            Trans1.Commit();
                        }
                    }
                }
            }

            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
            Freeze_operations = false;

        }

        private void Create_ML_object_data()
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
                            List2.Add("ObjectID of the rectangle");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            List1.Add("DrawingNum");
                            List2.Add("Alignment_number");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            List1.Add("BeginSta");
                            List2.Add("Matchline start");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                            List1.Add("EndSta");
                            List2.Add("Matchline end");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                            List1.Add("Center_X");
                            List2.Add("X in modelspace");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                            List1.Add("Center_Y");
                            List2.Add("Y in modelspace");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                            List1.Add("Rotation");
                            List2.Add("E-W viewport line rotation");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                            List1.Add("Width");
                            List2.Add("Matchline rectangle width");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                            List1.Add("Height");
                            List2.Add("Matchline rectangle height");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                            List1.Add("Type");
                            List2.Add("Type of drawing related to the rectangle");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            List1.Add("Note1");
                            List2.Add("Notes");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            List1.Add("Version");
                            List2.Add("Version number");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            List1.Add("DateMod");
                            List2.Add("DateMod");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            Functions.Get_object_data_table("AGEN_DrawingIndex_ML", "Generated by AGEN", List1, List2, List3);


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

        private void Append_ML_object_data()
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

                            Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

                            for (int i = 0; i < Data_table_matchline.Rows.Count; ++i)
                            {

                                List<object> List1 = new List<object>();
                                List<Autodesk.Gis.Map.Constants.DataType> List2 = new List<Autodesk.Gis.Map.Constants.DataType>();

                                String ObjID = Data_table_matchline.Rows[i]["OBJECT_ID"].ToString();

                                List1.Add(ObjID);
                                List2.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                Entity Ent1 = (Entity)Trans1.GetObject(Functions.GetObjectId(ThisDrawing.Database, ObjID), OpenMode.ForWrite);

                                List1.Add(Data_table_matchline.Rows[i]["FILE_NAME"].ToString());
                                List2.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                List1.Add((double)Data_table_matchline.Rows[i]["M1"]);
                                List2.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                List1.Add((double)Data_table_matchline.Rows[i]["M2"]);
                                List2.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                List1.Add((double)Data_table_matchline.Rows[i]["X"]);
                                List2.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                List1.Add((double)Data_table_matchline.Rows[i]["Y"]);
                                List2.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                List1.Add((double)Data_table_matchline.Rows[i]["ROTATION"]);
                                List2.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                List1.Add((double)Data_table_matchline.Rows[i]["WIDTH"]);
                                List2.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                List1.Add((double)Data_table_matchline.Rows[i]["HEIGHT"]);
                                List2.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                List1.Add("Alignment Sheet");
                                List2.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                Functions.Populate_object_data_table(Tables1, ObjID, "AGEN_DrawingIndex_ML", List1, List2);
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

        }

        private void Create_VP_object_data()
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
                            List2.Add("ObjectID of the rectangle");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            List1.Add("DrawingNum");
                            List2.Add("Alignment_number");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            List1.Add("Center_X");
                            List2.Add("X in modelspace");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                            List1.Add("Center_Y");
                            List2.Add("Y in modelspace");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                            List1.Add("Rotation");
                            List2.Add("E-W viewport line rotation");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                            List1.Add("Width");
                            List2.Add("Matchline rectangle width");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                            List1.Add("Height");
                            List2.Add("Matchline rectangle height");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                            List1.Add("Type");
                            List2.Add("Type of drawing related to the rectangle");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            List1.Add("Note1");
                            List2.Add("Notes");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            List1.Add("Version");
                            List2.Add("Version number");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            List1.Add("DateMod");
                            List2.Add("DateMod");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            Functions.Get_object_data_table("AGEN_DrawingIndex_VP", "Generated by AGEN", List1, List2, List3);


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

        private void Append_VP_object_data()
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
                            Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                            for (int i = 0; i < Data_table_Main_VP.Rows.Count; ++i)
                            {

                                List<object> List1 = new List<object>();
                                List<Autodesk.Gis.Map.Constants.DataType> List2 = new List<Autodesk.Gis.Map.Constants.DataType>();

                                String ObjID = Data_table_Main_VP.Rows[i]["OBJECT_ID"].ToString();

                                List1.Add(ObjID);
                                List2.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                Entity Ent1 = (Entity)Trans1.GetObject(Functions.GetObjectId(ThisDrawing.Database, ObjID), OpenMode.ForWrite);

                                List1.Add(Data_table_Main_VP.Rows[i]["FILE_NAME"].ToString());
                                List2.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                List1.Add((double)Data_table_Main_VP.Rows[i]["X"]);
                                List2.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                List1.Add((double)Data_table_Main_VP.Rows[i]["Y"]);
                                List2.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                List1.Add((double)Data_table_Main_VP.Rows[i]["ROTATION"]);
                                List2.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                List1.Add((double)Data_table_Main_VP.Rows[i]["WIDTH"]);
                                List2.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                List1.Add((double)Data_table_Main_VP.Rows[i]["HEIGHT"]);
                                List2.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                List1.Add("Alignment Sheet");
                                List2.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                Functions.Populate_object_data_table(Tables1, ObjID, "AGEN_DrawingIndex_VP", List1, List2);
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

        }

        private void button_add_config_file_Click(object sender, EventArgs e)
        {
            if (Freeze_operations == false)
            {
                Freeze_operations = true;
                try
                {
                    if (dataGridView_config_files_path.Rows.Count == 0)
                    {
                        Data_table_Config_files_path = new System.Data.DataTable();
                        Data_table_Config_files_path.Columns.Add("TYPE", typeof(String));
                        Data_table_Config_files_path.Columns.Add("PATH", typeof(String));
                    }

                    if (comboBox_config_files.Text != "")
                    {
                        string config_type = comboBox_config_files.Text;
                        if (config_type != "")
                        {
                            bool Exista = false;
                            if (Data_table_Config_files_path.Rows.Count > 0)
                            {

                                for (int i = 0; i < Data_table_Config_files_path.Rows.Count; ++i)
                                {
                                    string CT = Data_table_Config_files_path.Rows[i][0].ToString();
                                    if (config_type == CT)
                                    {
                                        Exista = true;
                                        using (OpenFileDialog fbd = new OpenFileDialog())
                                        {
                                            fbd.Multiselect = false;
                                            fbd.Filter = "Excel files (*.xlsx)|*.xlsx";

                                            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                                            {
                                                Data_table_Config_files_path.Rows[i][1] = fbd.FileName;
                                                dataGridView_config_files_path.DataSource = Data_table_Config_files_path;
                                                dataGridView_config_files_path.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

                                                if (config_type == "Centerline")
                                                {
                                                    CL_file = fbd.FileName;
                                                }
                                                if (config_type == "Sheet Index")
                                                {
                                                    Sheet_index_file = fbd.FileName;
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            if (Exista == false)
                            {
                                using (OpenFileDialog fbd = new OpenFileDialog())
                                {
                                    fbd.Multiselect = false;
                                    fbd.Filter = "Excel files (*.xlsx)|*.xlsx";
                                    if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                                    {
                                        Data_table_Config_files_path.Rows.Add();
                                        Data_table_Config_files_path.Rows[Data_table_Config_files_path.Rows.Count - 1][0] = config_type;
                                        Data_table_Config_files_path.Rows[Data_table_Config_files_path.Rows.Count - 1][1] = fbd.FileName;
                                        dataGridView_config_files_path.DataSource = Data_table_Config_files_path;
                                        dataGridView_config_files_path.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

                                        if (config_type == "Centerline")
                                        {
                                            CL_file = fbd.FileName;
                                        }
                                        if (config_type == "Sheet Index")
                                        {
                                            Sheet_index_file = fbd.FileName;
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
                Freeze_operations = false;
            }

        }

        private void button_viewport_target_areas_Click(object sender, EventArgs e)
        {

            if (Freeze_operations == false)
            {
                Freeze_operations = true;
                try
                {

                    if (dataGridView_viewport_target_areas.Rows.Count == 0)
                    {
                        Data_table_viewport_target_areas = new System.Data.DataTable();
                        Data_table_viewport_target_areas.Columns.Add("TYPE", typeof(String));
                        Data_table_viewport_target_areas.Columns.Add("CUSTOMSCALE", typeof(double));
                        Data_table_viewport_target_areas.Columns.Add("WIDTH", typeof(double));
                        Data_table_viewport_target_areas.Columns.Add("HEIGHT", typeof(double));
                        Data_table_viewport_target_areas.Columns.Add("PS_X", typeof(double));
                        Data_table_viewport_target_areas.Columns.Add("PS_Y", typeof(double));
                    }

                    double x1 = 0;
                    double y1 = 0;
                    double x2 = 0;
                    double y2 = 0;

                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                    using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {

                            if (comboBox_viewport_target_areas.Text != "")
                            {

                                Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                                Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);

                                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the lower left point");
                                PP1.AllowNone = false;
                                Point_res1 = Editor1.GetPoint(PP1);


                                if (Point_res1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                {
                                    Freeze_operations = false;
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    return;
                                }

                                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                                Autodesk.AutoCAD.EditorInput.PromptPointOptions PP2;
                                PP2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the upper right point");
                                PP2.AllowNone = false;
                                PP2.UseBasePoint = true;
                                PP2.BasePoint = Point_res1.Value;

                                Point_res2 = Editor1.GetPoint(PP2);

                                if (Point_res2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                {
                                    Freeze_operations = false;
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    return;
                                }

                                x1 = Point_res1.Value.X;
                                y1 = Point_res1.Value.Y;
                                x2 = Point_res2.Value.X;
                                y2 = Point_res2.Value.Y;


                                string viewport_type = comboBox_viewport_target_areas.Text;
                                bool Exista = false;

                                if (viewport_type != "")
                                {

                                    if (Data_table_viewport_target_areas.Rows.Count > 0)
                                    {

                                        for (int i = 0; i < Data_table_viewport_target_areas.Rows.Count; ++i)
                                        {
                                            string CT = Data_table_viewport_target_areas.Rows[i][0].ToString();
                                            if (viewport_type == CT)
                                            {
                                                Vw_width = Math.Abs(x1 - x2);
                                                Vw_height = Math.Abs(y1 - y2);
                                                Vw_ps_x = (x1 + x2) / 2;
                                                Vw_ps_y = (y1 + y2) / 2;

                                                Data_table_viewport_target_areas.Rows[i][1] = Vw_scale;
                                                Data_table_viewport_target_areas.Rows[i][2] = Vw_width;
                                                Data_table_viewport_target_areas.Rows[i][3] = Vw_height;
                                                Data_table_viewport_target_areas.Rows[i][4] = Vw_ps_x;
                                                Data_table_viewport_target_areas.Rows[i][5] = Vw_ps_y;
                                                dataGridView_viewport_target_areas.DataSource = Data_table_viewport_target_areas;
                                                dataGridView_viewport_target_areas.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);


                                                Exista = true;

                                            }
                                        }

                                    }

                                    if (Exista == false)
                                    {
                                        Vw_width = Math.Abs(x1 - x2);
                                        Vw_height = Math.Abs(y1 - y2);
                                        Vw_ps_x = (x1 + x2) / 2;
                                        Vw_ps_y = (y1 + y2) / 2;

                                        Data_table_viewport_target_areas.Rows.Add();
                                        Data_table_viewport_target_areas.Rows[Data_table_viewport_target_areas.Rows.Count - 1][0] = comboBox_viewport_target_areas.Text;
                                        Data_table_viewport_target_areas.Rows[Data_table_viewport_target_areas.Rows.Count - 1][1] = Vw_scale;
                                        Data_table_viewport_target_areas.Rows[Data_table_viewport_target_areas.Rows.Count - 1][2] = Vw_width;
                                        Data_table_viewport_target_areas.Rows[Data_table_viewport_target_areas.Rows.Count - 1][3] = Vw_height;
                                        Data_table_viewport_target_areas.Rows[Data_table_viewport_target_areas.Rows.Count - 1][4] = Vw_ps_x;
                                        Data_table_viewport_target_areas.Rows[Data_table_viewport_target_areas.Rows.Count - 1][5] = Vw_ps_y;
                                        dataGridView_viewport_target_areas.DataSource = Data_table_viewport_target_areas;
                                        dataGridView_viewport_target_areas.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

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
                Freeze_operations = false;
            }
        }

        private void button_browse_select_output_folder_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    textBox_output_folder.Text = fbd.SelectedPath.ToString();
                }

            }
        }

        private void button_browser_dwt_Click(object sender, EventArgs e)
        {

            using (OpenFileDialog fbd = new OpenFileDialog())
            {
                fbd.Multiselect = false;
                fbd.Filter = "autocad template files (*.dwt)|*.dwt";
                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    textBox_template_name.Text = fbd.FileName;
                }

            }
        }

        private void button_align_config_saveall_Click(object sender, EventArgs e)
        {

            System.Data.DataTable Data_table_config = new System.Data.DataTable();
            Data_table_config.Columns.Add("A", typeof(String));
            Data_table_config.Columns.Add("B", typeof(String));

            for (int i = 1; i <= 20; ++i)
            {
                Data_table_config.Rows.Add();


            }

            Data_table_config.Rows[0][0] = "Client Name";
            Data_table_config.Rows[0][1] = textBox_client_name.Text;

            Data_table_config.Rows[1][0] = "Project Name";
            Data_table_config.Rows[1][1] = textBox_project_name.Text;

            Data_table_config.Rows[2][0] = "Segment Name";
            Data_table_config.Rows[2][1] = textBox_segment_name.Text;

            Data_table_config.Rows[3][0] = "Template";
            Data_table_config.Rows[3][1] = textBox_template_name.Text;

            Data_table_config.Rows[4][0] = "Output folder";
            Data_table_config.Rows[4][1] = textBox_output_folder.Text;

            Data_table_config.Rows[5][0] = "Prefix File Name";
            Data_table_config.Rows[5][1] = textBox_prefix_name.Text;

            Data_table_config.Rows[6][0] = "Start numbering";
            Data_table_config.Rows[6][1] = textBox_name_start_number.Text;

            Data_table_config.Rows[7][0] = "Increment";
            Data_table_config.Rows[7][1] = textBox_name_increment.Text;

            Data_table_config.Rows[8][0] = "Main Viewport PS X center";
            Data_table_config.Rows[8][1] = Vw_ps_x.ToString();

            Data_table_config.Rows[9][0] = "Main Viewport PS Y center";
            Data_table_config.Rows[9][1] = Vw_ps_y.ToString();

            Data_table_config.Rows[10][0] = "Main Viewport width";
            Data_table_config.Rows[10][1] = Vw_width.ToString();

            Data_table_config.Rows[11][0] = "Main Viewport height";
            Data_table_config.Rows[11][1] = Vw_height.ToString();

            Data_table_config.Rows[12][0] = "Main Viewport scale";
            Data_table_config.Rows[12][1] = Vw_scale.ToString();

            Data_table_config.Rows[13][0] = "Basefile location folder";
            Data_table_config.Rows[13][1] = textBox_basefiles_folder.Text;

            Data_table_config.Rows[14][0] = "Sheet Index excel file";
            Data_table_config.Rows[14][1] = Sheet_index_file;

            Data_table_config.Rows[15][0] = "Centerline excel file";
            Data_table_config.Rows[15][1] = CL_file;

            Data_table_config.Rows[16][0] = "North Arrow Block name";
            Data_table_config.Rows[16][1] = "";

            Data_table_config.Rows[17][0] = "North Arrow PS X";
            Data_table_config.Rows[17][1] = "0";

            Data_table_config.Rows[18][0] = "North Arrow PS Y";
            Data_table_config.Rows[18][1] = "0";

            Data_table_config.Rows[19][0] = "North Arrow scale";
            Data_table_config.Rows[19][1] = "1";

            Data_table_config.Rows[20][0] = "Centerline layer name";
            Data_table_config.Rows[20][1] = Layer_name_Poly2D;

            Data_table_config.Rows[21][0] = "Centerline layer color";
            Data_table_config.Rows[21][1] = Color_index_Layer_name_Poly2D.ToString();

            Data_table_config.Rows[22][0] = "Centerline Handle (Object ID)";
            Data_table_config.Rows[22][1] = "0";



            save_new_config_file(Data_table_config);


        }

        private void save_new_config_file(System.Data.DataTable Data_table_config)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application Excel1 = new Microsoft.Office.Interop.Excel.Application();
                if (Excel1 == null)
                {
                    MessageBox.Show("PROBLEM WITH EXCEL!");
                    return;
                }

                Excel1.Visible = true;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Add();
                Microsoft.Office.Interop.Excel._Worksheet W1 = Workbook1.Worksheets[1];

                try
                {
                    W1.Cells.NumberFormat = "@";

                    int maxRows = Data_table_config.Rows.Count, maxCols = Data_table_config.Columns.Count;
                    Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[1, 1], W1.Cells[maxRows, maxCols]];

                    object[,] values = new object[maxRows, maxCols];
                    for (int row = 0; row < maxRows; row++)
                    {
                        for (int col = 0; col < maxCols; col++)
                        {
                            if (Data_table_config.Rows[row][col] != DBNull.Value)
                            {
                                values[row, col] = Data_table_config.Rows[row][col];
                            }
                        }
                    }
                    range1.Value2 = values;

                    range1.Columns.AutoFit();


                    SaveFileDialog Save_dlg = new SaveFileDialog();
                    Save_dlg.Filter = "Excel file|*.xlsx";


                    if (Save_dlg.ShowDialog() == DialogResult.OK)
                    {
                        string path1 = Save_dlg.FileName;
                        Workbook1.SaveAs(path1);
                        Workbook1.Close();
                        Excel1.Quit();
                        textBox_config_file_location.Text = path1;

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
                    if (Excel1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }

        }

        private void Button_home_navbar_alignments_Click(object sender, EventArgs e)
        {
            tabControl_Nav.SelectedTab = tabPage2;
            tabControl_work.SelectedTab = tabPageblank;
        }

        private void label_back_Click(object sender, EventArgs e)
        {
            tabControl_Nav.SelectedTab = tabPage1;
            tabControl_work.SelectedTab = tabPagehome;
        }

        private void button_browse_basefiles_folder_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    textBox_basefiles_folder.Text = fbd.SelectedPath.ToString();
                }

            }
        }

        private void button_build_cl_and_ml_Click(object sender, EventArgs e)
        {
            if (System.IO.File.Exists(textBox_config_file_location.Text) == false)
            {
                MessageBox.Show("no config file loaded\r\nOperation aborted");
                return;
            }

            if (Freeze_operations == false)
            {
                Freeze_operations = true;

                if (Data_table_matchline != null && Data_table_centerline != null)
                {
                    if (Data_table_matchline.Rows.Count > 0 && Data_table_centerline.Rows.Count > 0)
                    {
                        try
                        {
                            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                            Functions.Creaza_layer(Layer_name_Poly2D, Color_index_Layer_name_Poly2D, true);

                            using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                            {
                                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                                {
                                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                                    Poly2D = new Polyline();
                                    for (int i = 0; i < Data_table_centerline.Rows.Count; ++i)
                                    {
                                        double x = 0;
                                        double y = 0;
                                        if (Data_table_centerline.Rows[i]["X"] != DBNull.Value)
                                        {
                                            x = (double)Data_table_centerline.Rows[i]["X"];
                                        }
                                        else
                                        {
                                            Freeze_operations = false;
                                            MessageBox.Show("no X value for centerline in row " + (i).ToString());
                                            return;
                                        }
                                        if (Data_table_centerline.Rows[i]["Y"] != DBNull.Value)
                                        {
                                            y = (double)Data_table_centerline.Rows[i]["Y"];
                                        }
                                        else
                                        {
                                            Freeze_operations = false;
                                            MessageBox.Show("no Y value for centerline in row " + (i).ToString());
                                            return;
                                        }
                                        Poly2D.AddVertexAt(i, new Point2d(x, y), 0, 0, 0);

                                    }
                                    Poly2D.Layer = Layer_name_Poly2D;

                                    BTrecord.AppendEntity(Poly2D);
                                    Trans1.AddNewlyCreatedDBObject(Poly2D, true);
                                    Poly2D_handle = Poly2D.ObjectId.Handle.Value.ToString();


                                    int CI = 1;
                                    Functions.Creaza_layer(Layer_name_ML_rectangle, 4, false);

                                    for (int i = 0; i < Data_table_matchline.Rows.Count; ++i)
                                    {
                                        double Cx = 0;
                                        double Cy = 0;
                                        double rotation = 0;
                                        double width1 = 0;
                                        double height1 = 0;
                                        if (Data_table_matchline.Rows[i]["X"] != DBNull.Value)
                                        {
                                            Cx = (double)Data_table_matchline.Rows[i]["X"];
                                        }
                                        else
                                        {
                                            Freeze_operations = false;
                                            MessageBox.Show("no matchline rectangle X value for sheet index in row " + (i).ToString());
                                            return;
                                        }
                                        if (Data_table_matchline.Rows[i]["Y"] != DBNull.Value)
                                        {
                                            Cy = (double)Data_table_matchline.Rows[i]["Y"];
                                        }
                                        else
                                        {
                                            Freeze_operations = false;
                                            MessageBox.Show("no matchline rectangle Y value for sheet index in row " + (i).ToString());
                                            return;
                                        }

                                        if (Data_table_matchline.Rows[i]["ROTATION"] != DBNull.Value)
                                        {
                                            rotation = (double)Data_table_matchline.Rows[i]["ROTATION"] * Math.PI / 180;
                                        }
                                        else
                                        {
                                            Freeze_operations = false;
                                            MessageBox.Show("no matchline rectangle ROTATION value for sheet index in row " + (i).ToString());
                                            return;
                                        }

                                        if (Data_table_matchline.Rows[i]["HEIGHT"] != DBNull.Value)
                                        {
                                            height1 = (double)Data_table_matchline.Rows[i]["HEIGHT"];
                                        }
                                        else
                                        {
                                            Freeze_operations = false;
                                            MessageBox.Show("no matchline rectangle Height value for sheet index in row " + (i).ToString());
                                            return;
                                        }

                                        if (Data_table_matchline.Rows[i]["WIDTH"] != DBNull.Value)
                                        {
                                            width1 = (double)Data_table_matchline.Rows[i]["WIDTH"];
                                        }
                                        else
                                        {
                                            Freeze_operations = false;
                                            MessageBox.Show("no matchline rectangle Width value for sheet index in row " + (i).ToString());
                                            return;
                                        }

                                        Polyline Poly1 = creaza_rectangle_from_one_point(new Point3d(Cx, Cy, 0), rotation, width1, height1, CI);
                                        Poly1.Layer = Layer_name_ML_rectangle;
                                        BTrecord.AppendEntity(Poly1);
                                        Trans1.AddNewlyCreatedDBObject(Poly1, true);
                                        Data_table_matchline.Rows[i]["OBJECT_ID"] = Poly1.ObjectId.Handle.Value.ToString();
                                        CI = CI + 1;
                                        if (CI > 7) CI = 1;
                                    }

                                    Create_ML_object_data();
                                    Append_ML_object_data();

                                    Trans1.Commit();
                                    dataGridView_sheet_index.DataSource = Data_table_matchline;
                                    dataGridView_sheet_index.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

                                    if (Data_table_Config_files_path != null)
                                    {
                                        if (Data_table_Config_files_path.Rows.Count > 0)
                                        {
                                            for (int i = 0; i < Data_table_Config_files_path.Rows.Count; ++i)
                                            {
                                                string Type1 = Data_table_Config_files_path.Rows[i][0].ToString();

                                                if (Type1 == "Centerline")
                                                {
                                                    Populate_centerline_file(Data_table_Config_files_path.Rows[i][1].ToString());
                                                    Update_poly2d_handle_layer_colorindex();
                                                }
                                                if (Type1 == "Sheet Index")
                                                {
                                                    Populate_sheet_index_file(Data_table_Config_files_path.Rows[i][1].ToString());
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
                    }
                    Freeze_operations = false;
                }
            }
        }








    }
}
