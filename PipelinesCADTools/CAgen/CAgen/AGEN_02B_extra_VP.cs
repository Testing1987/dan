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
    public partial class AGEN_extra_vp_form : Form
    {
        bool clickdragdown;
        Point lastLocation;
        bool Template_is_open = false;

        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button1);
            lista_butoane.Add(button2);
            lista_butoane.Add(button3);
            lista_butoane.Add(button4);
            lista_butoane.Add(button5);
            lista_butoane.Add(button_minus1);
            lista_butoane.Add(button_minimize);
            lista_butoane.Add(button_close);
            lista_butoane.Add(button_create_extra);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = false;
            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button1);
            lista_butoane.Add(button2);
            lista_butoane.Add(button3);
            lista_butoane.Add(button4);
            lista_butoane.Add(button5);
            lista_butoane.Add(button_minus1);
            lista_butoane.Add(button_minimize);
            lista_butoane.Add(button_close);
            lista_butoane.Add(button_create_extra);
            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }






        public AGEN_extra_vp_form()
        {
            InitializeComponent();


            if (_AGEN_mainform.Data_Table_extra_mainVP != null && _AGEN_mainform.Data_Table_extra_mainVP.Rows.Count > 0)
            {

                for (int i = 0; i < _AGEN_mainform.Data_Table_extra_mainVP.Rows.Count; ++i)
                {


                    if (_AGEN_mainform.Data_Table_extra_mainVP.Rows[i]["Custom_scale"] != DBNull.Value &&
                        _AGEN_mainform.Data_Table_extra_mainVP.Rows[i]["viewport_width"] != DBNull.Value && _AGEN_mainform.Data_Table_extra_mainVP.Rows[i]["viewport_height"] != DBNull.Value &&
                        _AGEN_mainform.Data_Table_extra_mainVP.Rows[i]["viewport_ps_x"] != DBNull.Value && _AGEN_mainform.Data_Table_extra_mainVP.Rows[i]["viewport_ps_y"] != DBNull.Value &&
                        Functions.IsNumeric(Convert.ToString(_AGEN_mainform.Data_Table_extra_mainVP.Rows[i]["Custom_scale"])) == true &&
                        Functions.IsNumeric(Convert.ToString(_AGEN_mainform.Data_Table_extra_mainVP.Rows[i]["viewport_width"])) == true &&
                        Functions.IsNumeric(Convert.ToString(_AGEN_mainform.Data_Table_extra_mainVP.Rows[i]["viewport_height"])) == true &&
                        Functions.IsNumeric(Convert.ToString(_AGEN_mainform.Data_Table_extra_mainVP.Rows[i]["viewport_ps_x"])) == true &&
                        Functions.IsNumeric(Convert.ToString(_AGEN_mainform.Data_Table_extra_mainVP.Rows[i]["viewport_ps_y"])) == true)
                    {
                        if (i == 0)
                        {
                            button1.BackgroundImage = Alignment_mdi.Properties.Resources.check;
                            label1.ForeColor = Color.White;
                        }

                        if (i == 1)
                        {
                            button2.BackgroundImage = Alignment_mdi.Properties.Resources.check;
                            label2.ForeColor = Color.White;
                        }

                        if (i == 2)
                        {
                            button3.BackgroundImage = Alignment_mdi.Properties.Resources.check;
                            label3.ForeColor = Color.White;
                        }

                        if (i == 3)
                        {
                            button4.BackgroundImage = Alignment_mdi.Properties.Resources.check;
                            label4.ForeColor = Color.White;
                        }

                        if (i == 4)
                        {
                            button5.BackgroundImage = Alignment_mdi.Properties.Resources.check;
                            label5.ForeColor = Color.White;
                        }
                    }

                }
            }
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
            maximize_agen();

            this.Close();
        }
        private void button_minimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }





        private void button_remove1_Click(object sender, EventArgs e)
        {

            if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count > 0)
            {
                _AGEN_mainform.Data_Table_extra_mainVP.Rows[_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count-1].Delete();
            }
            if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count == 1)
            {
                button2.BackgroundImage = Alignment_mdi.Properties.Resources.selectbluexs;
                label2.ForeColor = Color.Black;
            }
            if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count == 2)
            {
                button3.BackgroundImage = Alignment_mdi.Properties.Resources.selectbluexs;
                label3.ForeColor = Color.Black;
            }
            if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count == 3)
            {
                button4.BackgroundImage = Alignment_mdi.Properties.Resources.selectbluexs;
                label4.ForeColor = Color.Black;
            }
            if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count == 4)
            {
                button5.BackgroundImage = Alignment_mdi.Properties.Resources.selectbluexs;
                label5.ForeColor = Color.Black;
            }

        }


        private void button_remove2_Click(object sender, EventArgs e)
        {

            if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count > 1)
            {
                _AGEN_mainform.Data_Table_extra_mainVP.Rows[1].Delete();
            }
            if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count == 2)
            {
                button3.BackgroundImage = Alignment_mdi.Properties.Resources.selectbluexs;
                label3.ForeColor = Color.Black;
            }
            if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count == 3)
            {
                button4.BackgroundImage = Alignment_mdi.Properties.Resources.selectbluexs;
                label4.ForeColor = Color.Black;
            }
            if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count == 4)
            {
                button5.BackgroundImage = Alignment_mdi.Properties.Resources.selectbluexs;
                label5.ForeColor = Color.Black;
            }
        }

        private void button_remove3_Click(object sender, EventArgs e)
        {
            if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count > 2)
            {
                _AGEN_mainform.Data_Table_extra_mainVP.Rows[2].Delete();
            }

            if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count == 3)
            {
                button4.BackgroundImage = Alignment_mdi.Properties.Resources.selectbluexs;
                label4.ForeColor = Color.Black;
            }
            if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count == 4)
            {
                button5.BackgroundImage = Alignment_mdi.Properties.Resources.selectbluexs;
                label5.ForeColor = Color.Black;
            }
        }

        private void button_remove4_Click(object sender, EventArgs e)
        {
            if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count > 3)
            {
                _AGEN_mainform.Data_Table_extra_mainVP.Rows[3].Delete();
            }

            if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count == 4)
            {
                button5.BackgroundImage = Alignment_mdi.Properties.Resources.selectbluexs;
                label5.ForeColor = Color.Black;
            }
        }

        private void button_remove5_Click(object sender, EventArgs e)
        {
            if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count == 5)
            {
                _AGEN_mainform.Data_Table_extra_mainVP.Rows[4].Delete();
                button5.BackgroundImage = Alignment_mdi.Properties.Resources.selectbluexs;
                label5.ForeColor = Color.Black;
            }
        }



        private void button_create_bands_Click(object sender, EventArgs e)
        {







            transfera_extra_vp_settings_to_excel();

            _AGEN_mainform.tpage_setup.display_checkboxes_into_generation_page();
            _AGEN_mainform.tpage_viewport_settings.creeaza_display_data_table(Functions.Creaza_lista_regular_vp_picked(), Functions.Creaza_lista_custom_vp_picked(), Functions.Creaza_lista_custom_vp_extra_picked());


            maximize_agen();
            button_Exit_Click(sender, e);


        }

        private void maximize_agen()
        {
            foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
            {
                if (Forma1 is Alignment_mdi._AGEN_mainform)
                {
                    Forma1.Focus();
                    Forma1.WindowState = System.Windows.Forms.FormWindowState.Normal;
                }
            }
        }




        public void transfera_extra_vp_settings_to_excel()
        {
            if (_AGEN_mainform.Data_Table_extra_mainVP != null)
            {
                if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count > 0)
                {
                    Functions.Kill_excel();

                    string cfg1 = System.IO.Path.GetFileName(_AGEN_mainform.config_path);
                    if (Functions.Get_if_workbook_is_open_in_Excel(cfg1) == true)
                    {
                        MessageBox.Show("Please close the " + cfg1 + " file");
                        return;
                    }

                    set_enable_false();
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

                        if (Excel1.Workbooks.Count==0) Excel1.Visible = _AGEN_mainform.ExcelVisible;

                        if (System.IO.File.Exists(_AGEN_mainform.config_path) == true)
                        {

                            Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(_AGEN_mainform.config_path);

                            Microsoft.Office.Interop.Excel.Worksheet W1 = null;

                            foreach (Microsoft.Office.Interop.Excel.Worksheet wsh1 in Workbook1.Worksheets)
                            {
                                if (wsh1.Name == "ExtraVP_data")
                                {
                                    W1 = wsh1;
                                }
                            }

                            if (W1 == null)
                            {
                                W1 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[Workbook1.Worksheets.Count], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                                W1.Name = "ExtraVP_data";

                            }


                            W1.Columns["A:XX"].Delete();


                            try
                            {

                                int maxRows = _AGEN_mainform.Data_Table_extra_mainVP.Rows.Count;
                                int maxCols = _AGEN_mainform.Data_Table_extra_mainVP.Columns.Count;

                                Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[2, 1], W1.Cells[maxRows + 1, maxCols]];
                                object[,] values1 = new object[maxRows, maxCols];

                                for (int i = 0; i < maxRows; ++i)
                                {
                                    for (int j = 0; j < maxCols; ++j)
                                    {
                                        if (_AGEN_mainform.Data_Table_extra_mainVP.Rows[i][j] != DBNull.Value)
                                        {
                                            values1[i, j] = _AGEN_mainform.Data_Table_extra_mainVP.Rows[i][j];
                                        }
                                    }
                                }

                                for (int i = 0; i < _AGEN_mainform.Data_Table_extra_mainVP.Columns.Count; ++i)
                                {
                                    W1.Cells[1, i + 1].value2 = _AGEN_mainform.Data_Table_extra_mainVP.Columns[i].ColumnName;
                                }

                                range1.Cells.NumberFormat = "@";
                                range1.Value2 = values1;

                                Functions.Color_border_range_inside(range1, 0);

                                Workbook1.Save();
                                Workbook1.Close();
                                if (Excel1.Workbooks.Count==0)
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
                    catch (System.Exception ex)
                    {
                        System.Windows.Forms.MessageBox.Show(ex.Message);

                    }
                    set_enable_true();




                }
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                set_enable_false();
                string strTemplatePath = _AGEN_mainform.tpage_viewport_settings.get_template_name_from_text_box();

                DocumentCollection DocumentManager1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = null;


                if (System.IO.File.Exists(strTemplatePath) == false)
                {
                    MessageBox.Show("template file not found");
                    set_enable_true();
                    return;
                }


                foreach (Document Doc in DocumentManager1)
                {
                    if (Doc.Name == strTemplatePath)
                    {
                        Template_is_open = true;
                        ThisDrawing = Doc;
                        DocumentManager1.MdiActiveDocument = ThisDrawing;
                    }
                }

                if (Template_is_open == false)
                {
                    ThisDrawing = DocumentCollectionExtension.Open(DocumentManager1, strTemplatePath, false);
                    Template_is_open = true;
                }

                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
                {
                    if (Forma1 is Alignment_mdi._AGEN_mainform)
                    {
                        if (Forma1.WindowState == System.Windows.Forms.FormWindowState.Normal)
                        {
                            Forma1.WindowState = System.Windows.Forms.FormWindowState.Minimized;
                        }
                    }
                }


                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {

                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        Functions.make_first_layout_active(Trans1, ThisDrawing.Database);

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                        PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify bottom left of extra plan view ");
                        PP1.AllowNone = false;
                        Point_res1 = Editor1.GetPoint(PP1);


                        if (Point_res1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            set_enable_true();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                        Alignment_mdi.Jig_rectangle_viewport_pick_points Jig2 = new Alignment_mdi.Jig_rectangle_viewport_pick_points();
                        Point_res2 = Jig2.StartJig(Point_res1.Value, 1, "\nSpecify top right of extra plan view ");

                        if (Point_res2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            set_enable_true();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }
                        double y1 = 0;
                        double y2 = 0;

                        double x1 = 0;
                        double x2 = 0;

                        y1 = Point_res1.Value.Y;
                        y2 = Point_res2.Value.Y;

                        if (y2 < y1)
                        {
                            double t1 = y1;
                            y1 = y2;
                            y2 = t1;
                        }

                        x1 = Point_res1.Value.X;
                        x2 = Point_res2.Value.X;

                        if (x2 < x1)
                        {
                            double t1 = x1;
                            x1 = x2;
                            x2 = t1;
                        }

                        if (_AGEN_mainform.Data_Table_extra_mainVP == null)
                        {
                            _AGEN_mainform.Data_Table_extra_mainVP = Functions.creeaza_extra_mainVP_data_table_structure();
                        }



                        if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count == 0)
                        {
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows.Add();
                        }

                        _AGEN_mainform.Data_Table_extra_mainVP.Rows[0]["Extra plan view"] = "Extra1";
                        _AGEN_mainform.Data_Table_extra_mainVP.Rows[0]["viewport_width"] = Math.Abs(x2 - x1);
                        _AGEN_mainform.Data_Table_extra_mainVP.Rows[0]["viewport_height"] = Math.Abs(y2 - y1);
                        _AGEN_mainform.Data_Table_extra_mainVP.Rows[0]["viewport_ps_x"] = (x2 + x1) / 2;
                        _AGEN_mainform.Data_Table_extra_mainVP.Rows[0]["viewport_ps_y"] = (y2 + y1) / 2;
                        _AGEN_mainform.Data_Table_extra_mainVP.Rows[0]["Custom_scale"] = _AGEN_mainform.Vw_scale;
                        Trans1.Commit();
                        button1.BackgroundImage = Alignment_mdi.Properties.Resources.check;
                        label1.ForeColor = Color.White;
                    }
                }



            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            set_enable_true();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                set_enable_false();
                string strTemplatePath = _AGEN_mainform.tpage_viewport_settings.get_template_name_from_text_box();

                DocumentCollection DocumentManager1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = null;


                if (System.IO.File.Exists(strTemplatePath) == false)
                {
                    MessageBox.Show("template file not found");
                    set_enable_true();
                    return;
                }

                foreach (Document Doc in DocumentManager1)
                {
                    if (Doc.Name == strTemplatePath)
                    {
                        Template_is_open = true;
                        ThisDrawing = Doc;
                        DocumentManager1.MdiActiveDocument = ThisDrawing;
                    }
                }

                if (Template_is_open == false)
                {
                    ThisDrawing = DocumentCollectionExtension.Open(DocumentManager1, strTemplatePath, false);
                    Template_is_open = true;
                }

                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
                {
                    if (Forma1 is Alignment_mdi._AGEN_mainform)
                    {
                        if (Forma1.WindowState == System.Windows.Forms.FormWindowState.Normal)
                        {
                            Forma1.WindowState = System.Windows.Forms.FormWindowState.Minimized;
                        }
                    }
                }


                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {

                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        Functions.make_first_layout_active(Trans1, ThisDrawing.Database);

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                        PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify bottom left of extra plan view ");
                        PP1.AllowNone = false;
                        Point_res1 = Editor1.GetPoint(PP1);


                        if (Point_res1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            set_enable_true();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                        Alignment_mdi.Jig_rectangle_viewport_pick_points Jig2 = new Alignment_mdi.Jig_rectangle_viewport_pick_points();
                        Point_res2 = Jig2.StartJig(Point_res1.Value, 1, "\nSpecify top right of extra plan view ");

                        if (Point_res2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            set_enable_true();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }
                        double y1 = 0;
                        double y2 = 0;

                        double x1 = 0;
                        double x2 = 0;

                        y1 = Point_res1.Value.Y;
                        y2 = Point_res2.Value.Y;

                        if (y2 < y1)
                        {
                            double t1 = y1;
                            y1 = y2;
                            y2 = t1;
                        }

                        x1 = Point_res1.Value.X;
                        x2 = Point_res2.Value.X;

                        if (x2 < x1)
                        {
                            double t1 = x1;
                            x1 = x2;
                            x2 = t1;
                        }

                        if (_AGEN_mainform.Data_Table_extra_mainVP == null)
                        {
                            _AGEN_mainform.Data_Table_extra_mainVP = Functions.creeaza_extra_mainVP_data_table_structure();
                        }



                        if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count == 0)
                        {
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows.Add();
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[0]["Extra plan view"] = "Extra1";
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[0]["viewport_width"] = Math.Abs(x2 - x1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[0]["viewport_height"] = Math.Abs(y2 - y1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[0]["viewport_ps_x"] = (x2 + x1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[0]["viewport_ps_y"] = (y2 + y1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[0]["Custom_scale"] = _AGEN_mainform.Vw_scale;
                            button1.BackgroundImage = Alignment_mdi.Properties.Resources.check;
                            label1.ForeColor = Color.White;
                        }

                        else if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count == 1)
                        {
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows.Add();
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[1]["Extra plan view"] = "Extra2";
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[1]["viewport_width"] = Math.Abs(x2 - x1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[1]["viewport_height"] = Math.Abs(y2 - y1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[1]["viewport_ps_x"] = (x2 + x1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[1]["viewport_ps_y"] = (y2 + y1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[1]["Custom_scale"] = _AGEN_mainform.Vw_scale;
                            button2.BackgroundImage = Alignment_mdi.Properties.Resources.check;
                            label2.ForeColor = Color.White;
                        }
                        else if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count>1)
                        {
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[1]["Extra plan view"] = "Extra2";
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[1]["viewport_width"] = Math.Abs(x2 - x1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[1]["viewport_height"] = Math.Abs(y2 - y1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[1]["viewport_ps_x"] = (x2 + x1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[1]["viewport_ps_y"] = (y2 + y1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[1]["Custom_scale"] = _AGEN_mainform.Vw_scale;
                            button2.BackgroundImage = Alignment_mdi.Properties.Resources.check;
                            label2.ForeColor = Color.White;
                        }



                        Trans1.Commit();

                    }
                }



            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            set_enable_true();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                set_enable_false();
                string strTemplatePath = _AGEN_mainform.tpage_viewport_settings.get_template_name_from_text_box();

                DocumentCollection DocumentManager1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = null;


                if (System.IO.File.Exists(strTemplatePath) == false)
                {
                    MessageBox.Show("template file not found");
                    set_enable_true();
                    return;
                }

                foreach (Document Doc in DocumentManager1)
                {
                    if (Doc.Name == strTemplatePath)
                    {
                       Template_is_open = true;
                        ThisDrawing = Doc;
                        DocumentManager1.MdiActiveDocument = ThisDrawing;
                    }
                }

                if (Template_is_open == false)
                {
                    ThisDrawing = DocumentCollectionExtension.Open(DocumentManager1, strTemplatePath, false);
                    Template_is_open = true;
                }

                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
                {
                    if (Forma1 is Alignment_mdi._AGEN_mainform)
                    {
                        if (Forma1.WindowState == System.Windows.Forms.FormWindowState.Normal)
                        {
                            Forma1.WindowState = System.Windows.Forms.FormWindowState.Minimized;
                        }
                    }
                }


                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {

                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        Functions.make_first_layout_active(Trans1, ThisDrawing.Database);

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                        PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify bottom left of extra plan view ");
                        PP1.AllowNone = false;
                        Point_res1 = Editor1.GetPoint(PP1);


                        if (Point_res1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            set_enable_true();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                        Alignment_mdi.Jig_rectangle_viewport_pick_points Jig2 = new Alignment_mdi.Jig_rectangle_viewport_pick_points();
                        Point_res2 = Jig2.StartJig(Point_res1.Value, 1, "\nSpecify top right of extra plan view ");

                        if (Point_res2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            set_enable_true();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }
                        double y1 = 0;
                        double y2 = 0;

                        double x1 = 0;
                        double x2 = 0;

                        y1 = Point_res1.Value.Y;
                        y2 = Point_res2.Value.Y;

                        if (y2 < y1)
                        {
                            double t1 = y1;
                            y1 = y2;
                            y2 = t1;
                        }

                        x1 = Point_res1.Value.X;
                        x2 = Point_res2.Value.X;

                        if (x2 < x1)
                        {
                            double t1 = x1;
                            x1 = x2;
                            x2 = t1;
                        }

                        if (_AGEN_mainform.Data_Table_extra_mainVP == null)
                        {
                            _AGEN_mainform.Data_Table_extra_mainVP = Functions.creeaza_extra_mainVP_data_table_structure();
                        }



                        if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count == 0)
                        {
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows.Add();
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[0]["Extra plan view"] = "Extra1";
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[0]["viewport_width"] = Math.Abs(x2 - x1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[0]["viewport_height"] = Math.Abs(y2 - y1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[0]["viewport_ps_x"] = (x2 + x1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[0]["viewport_ps_y"] = (y2 + y1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[0]["Custom_scale"] = _AGEN_mainform.Vw_scale;
                            button1.BackgroundImage = Alignment_mdi.Properties.Resources.check;
                            label1.ForeColor = Color.White;
                        }

                        else if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count == 1)
                        {
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows.Add();
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[1]["Extra plan view"] = "Extra2";
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[1]["viewport_width"] = Math.Abs(x2 - x1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[1]["viewport_height"] = Math.Abs(y2 - y1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[1]["viewport_ps_x"] = (x2 + x1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[1]["viewport_ps_y"] = (y2 + y1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[1]["Custom_scale"] = _AGEN_mainform.Vw_scale;
                            button2.BackgroundImage = Alignment_mdi.Properties.Resources.check;
                            label2.ForeColor = Color.White;
                        }
                        else if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count == 2)
                        {
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows.Add();
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[2]["Extra plan view"] = "Extra3";
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[2]["viewport_width"] = Math.Abs(x2 - x1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[2]["viewport_height"] = Math.Abs(y2 - y1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[2]["viewport_ps_x"] = (x2 + x1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[2]["viewport_ps_y"] = (y2 + y1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[2]["Custom_scale"] = _AGEN_mainform.Vw_scale;
                            button3.BackgroundImage = Alignment_mdi.Properties.Resources.check;
                            label3.ForeColor = Color.White;
                        }
                        else if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count >2)
                        {
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[2]["Extra plan view"] = "Extra3";
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[2]["viewport_width"] = Math.Abs(x2 - x1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[2]["viewport_height"] = Math.Abs(y2 - y1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[2]["viewport_ps_x"] = (x2 + x1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[2]["viewport_ps_y"] = (y2 + y1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[2]["Custom_scale"] = _AGEN_mainform.Vw_scale;
                            button3.BackgroundImage = Alignment_mdi.Properties.Resources.check;
                            label3.ForeColor = Color.White;
                        }


                        Trans1.Commit();

                    }
                }



            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            set_enable_true();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                set_enable_false();
                string strTemplatePath = _AGEN_mainform.tpage_viewport_settings.get_template_name_from_text_box();

                DocumentCollection DocumentManager1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = null;


                if (System.IO.File.Exists(strTemplatePath) == false)
                {
                    MessageBox.Show("template file not found");
                    set_enable_true();
                    return;
                }

                foreach (Document Doc in DocumentManager1)
                {
                    if (Doc.Name == strTemplatePath)
                    {
                       Template_is_open = true;
                        ThisDrawing = Doc;
                        DocumentManager1.MdiActiveDocument = ThisDrawing;
                    }
                }

                if (Template_is_open == false)
                {
                    ThisDrawing = DocumentCollectionExtension.Open(DocumentManager1, strTemplatePath, false);
                   Template_is_open = true;
                }

                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
                {
                    if (Forma1 is Alignment_mdi._AGEN_mainform)
                    {
                        if (Forma1.WindowState == System.Windows.Forms.FormWindowState.Normal)
                        {
                            Forma1.WindowState = System.Windows.Forms.FormWindowState.Minimized;
                        }
                    }
                }


                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {

                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        Functions.make_first_layout_active(Trans1, ThisDrawing.Database);

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                        PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify bottom left of extra plan view ");
                        PP1.AllowNone = false;
                        Point_res1 = Editor1.GetPoint(PP1);


                        if (Point_res1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            set_enable_true();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                        Alignment_mdi.Jig_rectangle_viewport_pick_points Jig2 = new Alignment_mdi.Jig_rectangle_viewport_pick_points();
                        Point_res2 = Jig2.StartJig(Point_res1.Value, 1, "\nSpecify top right of extra plan view ");

                        if (Point_res2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            set_enable_true();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }
                        double y1 = 0;
                        double y2 = 0;

                        double x1 = 0;
                        double x2 = 0;

                        y1 = Point_res1.Value.Y;
                        y2 = Point_res2.Value.Y;

                        if (y2 < y1)
                        {
                            double t1 = y1;
                            y1 = y2;
                            y2 = t1;
                        }

                        x1 = Point_res1.Value.X;
                        x2 = Point_res2.Value.X;

                        if (x2 < x1)
                        {
                            double t1 = x1;
                            x1 = x2;
                            x2 = t1;
                        }

                        if (_AGEN_mainform.Data_Table_extra_mainVP == null)
                        {
                            _AGEN_mainform.Data_Table_extra_mainVP = Functions.creeaza_extra_mainVP_data_table_structure();
                        }



                        if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count == 0)
                        {
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows.Add();
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[0]["Extra plan view"] = "Extra1";
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[0]["viewport_width"] = Math.Abs(x2 - x1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[0]["viewport_height"] = Math.Abs(y2 - y1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[0]["viewport_ps_x"] = (x2 + x1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[0]["viewport_ps_y"] = (y2 + y1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[0]["Custom_scale"] = _AGEN_mainform.Vw_scale;
                            button1.BackgroundImage = Alignment_mdi.Properties.Resources.check;
                            label1.ForeColor = Color.White;
                        }

                        else if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count == 1)
                        {
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows.Add();
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[1]["Extra plan view"] = "Extra2";
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[1]["viewport_width"] = Math.Abs(x2 - x1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[1]["viewport_height"] = Math.Abs(y2 - y1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[1]["viewport_ps_x"] = (x2 + x1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[1]["viewport_ps_y"] = (y2 + y1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[1]["Custom_scale"] = _AGEN_mainform.Vw_scale;
                            button2.BackgroundImage = Alignment_mdi.Properties.Resources.check;
                            label2.ForeColor = Color.White;
                        }
                        else if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count == 2)
                        {
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows.Add();
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[2]["Extra plan view"] = "Extra3";
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[2]["viewport_width"] = Math.Abs(x2 - x1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[2]["viewport_height"] = Math.Abs(y2 - y1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[2]["viewport_ps_x"] = (x2 + x1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[2]["viewport_ps_y"] = (y2 + y1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[2]["Custom_scale"] = _AGEN_mainform.Vw_scale;
                            button3.BackgroundImage = Alignment_mdi.Properties.Resources.check;
                            label3.ForeColor = Color.White;
                        }
                        else if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count == 3)
                        {
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows.Add();
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[3]["Extra plan view"] = "Extra4";
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[3]["viewport_width"] = Math.Abs(x2 - x1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[3]["viewport_height"] = Math.Abs(y2 - y1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[3]["viewport_ps_x"] = (x2 + x1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[3]["viewport_ps_y"] = (y2 + y1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[3]["Custom_scale"] = _AGEN_mainform.Vw_scale;
                            button4.BackgroundImage = Alignment_mdi.Properties.Resources.check;
                            label4.ForeColor = Color.White;
                        }
                        else if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count >3 )
                        {
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[3]["Extra plan view"] = "Extra4";
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[3]["viewport_width"] = Math.Abs(x2 - x1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[3]["viewport_height"] = Math.Abs(y2 - y1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[3]["viewport_ps_x"] = (x2 + x1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[3]["viewport_ps_y"] = (y2 + y1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[3]["Custom_scale"] = _AGEN_mainform.Vw_scale;
                            button4.BackgroundImage = Alignment_mdi.Properties.Resources.check;
                            label4.ForeColor = Color.White;
                        }

                        Trans1.Commit();

                    }
                }



            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            set_enable_true();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                set_enable_false();
                string strTemplatePath = _AGEN_mainform.tpage_viewport_settings.get_template_name_from_text_box();

                DocumentCollection DocumentManager1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = null;


                if (System.IO.File.Exists(strTemplatePath) == false)
                {
                    MessageBox.Show("template file not found");
                    set_enable_true();
                    return;
                }

                foreach (Document Doc in DocumentManager1)
                {
                    if (Doc.Name == strTemplatePath)
                    {
                       Template_is_open = true;
                        ThisDrawing = Doc;
                        DocumentManager1.MdiActiveDocument = ThisDrawing;
                    }
                }

                if (Template_is_open == false)
                {
                    ThisDrawing = DocumentCollectionExtension.Open(DocumentManager1, strTemplatePath, false);
                    Template_is_open = true;
                }

                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
                {
                    if (Forma1 is Alignment_mdi._AGEN_mainform)
                    {
                        if (Forma1.WindowState == System.Windows.Forms.FormWindowState.Normal)
                        {
                            Forma1.WindowState = System.Windows.Forms.FormWindowState.Minimized;
                        }
                    }
                }


                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {

                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        Functions.make_first_layout_active(Trans1, ThisDrawing.Database);

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                        PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify bottom left of extra plan view ");
                        PP1.AllowNone = false;
                        Point_res1 = Editor1.GetPoint(PP1);


                        if (Point_res1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            set_enable_true();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                        Alignment_mdi.Jig_rectangle_viewport_pick_points Jig2 = new Alignment_mdi.Jig_rectangle_viewport_pick_points();
                        Point_res2 = Jig2.StartJig(Point_res1.Value, 1, "\nSpecify top right of extra plan view ");

                        if (Point_res2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            set_enable_true();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }
                        double y1 = 0;
                        double y2 = 0;

                        double x1 = 0;
                        double x2 = 0;

                        y1 = Point_res1.Value.Y;
                        y2 = Point_res2.Value.Y;

                        if (y2 < y1)
                        {
                            double t1 = y1;
                            y1 = y2;
                            y2 = t1;
                        }

                        x1 = Point_res1.Value.X;
                        x2 = Point_res2.Value.X;

                        if (x2 < x1)
                        {
                            double t1 = x1;
                            x1 = x2;
                            x2 = t1;
                        }

                        if (_AGEN_mainform.Data_Table_extra_mainVP == null)
                        {
                            _AGEN_mainform.Data_Table_extra_mainVP = Functions.creeaza_extra_mainVP_data_table_structure();
                        }



                        if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count == 0)
                        {
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows.Add();
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[0]["Extra plan view"] = "Extra1";
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[0]["viewport_width"] = Math.Abs(x2 - x1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[0]["viewport_height"] = Math.Abs(y2 - y1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[0]["viewport_ps_x"] = (x2 + x1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[0]["viewport_ps_y"] = (y2 + y1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[0]["Custom_scale"] = _AGEN_mainform.Vw_scale;
                            button1.BackgroundImage = Alignment_mdi.Properties.Resources.check;
                            label1.ForeColor = Color.White;
                        }

                        else if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count == 1)
                        {
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows.Add();
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[1]["Extra plan view"] = "Extra2";
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[1]["viewport_width"] = Math.Abs(x2 - x1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[1]["viewport_height"] = Math.Abs(y2 - y1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[1]["viewport_ps_x"] = (x2 + x1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[1]["viewport_ps_y"] = (y2 + y1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[1]["Custom_scale"] = _AGEN_mainform.Vw_scale;
                            button2.BackgroundImage = Alignment_mdi.Properties.Resources.check;
                            label2.ForeColor = Color.White;
                        }
                        else if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count == 2)
                        {
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows.Add();
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[2]["Extra plan view"] = "Extra3";
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[2]["viewport_width"] = Math.Abs(x2 - x1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[2]["viewport_height"] = Math.Abs(y2 - y1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[2]["viewport_ps_x"] = (x2 + x1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[2]["viewport_ps_y"] = (y2 + y1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[2]["Custom_scale"] = _AGEN_mainform.Vw_scale;
                            button3.BackgroundImage = Alignment_mdi.Properties.Resources.check;
                            label3.ForeColor = Color.White;
                        }
                        else if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count == 3)
                        {
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows.Add();
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[3]["Extra plan view"] = "Extra4";
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[3]["viewport_width"] = Math.Abs(x2 - x1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[3]["viewport_height"] = Math.Abs(y2 - y1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[3]["viewport_ps_x"] = (x2 + x1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[3]["viewport_ps_y"] = (y2 + y1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[3]["Custom_scale"] = _AGEN_mainform.Vw_scale;
                            button4.BackgroundImage = Alignment_mdi.Properties.Resources.check;
                            label4.ForeColor = Color.White;
                        }
                        else if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count ==4)
                        {
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows.Add();

                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[4]["Extra plan view"] = "Extra5";
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[4]["viewport_width"] = Math.Abs(x2 - x1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[4]["viewport_height"] = Math.Abs(y2 - y1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[4]["viewport_ps_x"] = (x2 + x1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[4]["viewport_ps_y"] = (y2 + y1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[4]["Custom_scale"] = _AGEN_mainform.Vw_scale;
                            button5.BackgroundImage = Alignment_mdi.Properties.Resources.check;
                            label5.ForeColor = Color.White;


                        }
                        else if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count ==5)
                        {
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[4]["Extra plan view"] = "Extra5";
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[4]["viewport_width"] = Math.Abs(x2 - x1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[4]["viewport_height"] = Math.Abs(y2 - y1);
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[4]["viewport_ps_x"] = (x2 + x1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[4]["viewport_ps_y"] = (y2 + y1) / 2;
                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[4]["Custom_scale"] = _AGEN_mainform.Vw_scale;
                            button5.BackgroundImage = Alignment_mdi.Properties.Resources.check;
                            label5.ForeColor = Color.White;
                        }
                        Trans1.Commit();

                    }
                }



            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            set_enable_true();
        }
    }
}
