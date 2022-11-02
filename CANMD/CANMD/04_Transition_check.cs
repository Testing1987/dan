using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Alignment_mdi
{
    public partial class form_transitions : Form
    {

        System.Data.DataTable dt_t = null;

        public form_transitions()
        {
            InitializeComponent();
        }


        #region set enable true or false    
        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();


            lista_butoane.Add(button_refresh_blocks);
            lista_butoane.Add(button_transition_checks);



            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = false;
            }
        }
        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();



            lista_butoane.Add(button_refresh_blocks);
            lista_butoane.Add(button_transition_checks);





            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }
        #endregion

        private void button_refresh_blocks_Click(object sender, EventArgs e)
        {
            Functions.Incarca_existing_Blocks_to_combobox(comboBox_elbow);
            Functions.Incarca_existing_Blocks_to_combobox(comboBox_material);
            Functions.Incarca_existing_Blocks_to_combobox(comboBox_transition);
            Functions.Incarca_existing_Blocks_to_combobox(comboBox_block_to_be_placed);

            if (comboBox_elbow.Items.Contains("ELBOW") == true)
            {
                comboBox_elbow.SelectedIndex = comboBox_elbow.Items.IndexOf("ELBOW");
            }
            if (comboBox_material.Items.Contains("MAT") == true)
            {
                comboBox_material.SelectedIndex = comboBox_material.Items.IndexOf("MAT");
            }
            if (comboBox_transition.Items.Contains("T") == true)
            {
                comboBox_transition.SelectedIndex = comboBox_transition.Items.IndexOf("T");
            }
        }

        private void comboBox_material_DropDown(object sender, EventArgs e)
        {

        }

        private void button_transition_checks_Click(object sender, EventArgs e)
        {
            if (dt_t == null || dt_t.Rows.Count == 0) return;

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
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;


                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez.MessageForAdding = "\nSelect the bands:";
                        Prompt_rez.SingleOnly = false;
                        this.MdiParent.WindowState = FormWindowState.Minimized;

                        Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            this.MdiParent.WindowState = FormWindowState.Normal;
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        string objectid = "objectid";
                        string blockname = "Blockname";
                        string X = "x";
                        string Y = "y";
                        string distance1 = "distance1";
                        string mat = "mat";
                        string STA1 = "sta1";
                        string STA2 = "sta2";
                        string STA = "sta";
                        string description = "description";
                        string hasT = "has transition";

                        System.Data.DataTable dt1 = new System.Data.DataTable();
                        dt1.Columns.Add(objectid, typeof(string));
                        dt1.Columns.Add(blockname, typeof(string));
                        dt1.Columns.Add(X, typeof(double));
                        dt1.Columns.Add(Y, typeof(double));
                        dt1.Columns.Add(distance1, typeof(double));
                        dt1.Columns.Add(mat, typeof(string));
                        dt1.Columns.Add(STA1, typeof(double));
                        dt1.Columns.Add(STA2, typeof(double));
                        dt1.Columns.Add(STA, typeof(double));
                        dt1.Columns.Add(description, typeof(string));
                        dt1.Columns.Add(hasT, typeof(string));

                        System.Data.DataTable dt2 = new System.Data.DataTable();
                        dt2 = dt1.Clone();

                        for (int i = 0; i < Rezultat1.Value.Count; i++)
                        {
                            BlockReference bl1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as BlockReference;
                            if (bl1 != null)
                            {
                                string bn = Functions.get_block_name(bl1);
                                if (comboBox_material.Text == bn || comboBox_elbow.Text == bn)
                                {
                                    dt1.Rows.Add();
                                    dt1.Rows[dt1.Rows.Count - 1][objectid] = bl1.ObjectId.Handle.Value.ToString();
                                    dt1.Rows[dt1.Rows.Count - 1][blockname] = bn;
                                    Point3d ptins = bl1.Position;

                                    dt1.Rows[dt1.Rows.Count - 1][X] = Math.Round(ptins.X, 3);
                                    dt1.Rows[dt1.Rows.Count - 1][Y] = Math.Round(ptins.Y, 3);
                                    if (bl1.IsDynamicBlock == true)
                                    {
                                        double d1 = Functions.Get_distance1_block(bl1);
                                        dt1.Rows[dt1.Rows.Count - 1][distance1] = d1;
                                    }

                                    if (bl1.AttributeCollection.Count > 0)
                                    {
                                        foreach (ObjectId id1 in bl1.AttributeCollection)
                                        {
                                            AttributeReference atr1 = Trans1.GetObject(id1, OpenMode.ForRead) as AttributeReference;
                                            if (atr1 != null)
                                            {
                                                string tag1 = atr1.Tag.ToUpper();
                                                string val1 = atr1.TextString.Replace("+", "").Replace(" ", "");
                                                string val2 = atr1.TextString;
                                                if (Functions.IsNumeric(val1) == true)
                                                {
                                                    if (tag1 == "STA")
                                                    {
                                                        dt1.Rows[dt1.Rows.Count - 1][STA] = Convert.ToDouble(val1);

                                                    }
                                                    if (tag1 == "STA1")
                                                    {
                                                        dt1.Rows[dt1.Rows.Count - 1][STA1] = Convert.ToDouble(val1);

                                                    }
                                                    if (tag1 == "STA2")
                                                    {
                                                        dt1.Rows[dt1.Rows.Count - 1][STA2] = Convert.ToDouble(val1);

                                                    }
                                                }

                                                if (tag1 == "MAT")
                                                {
                                                    dt1.Rows[dt1.Rows.Count - 1][mat] = val2;
                                                }

                                                if (tag1 == "DESCR")
                                                {
                                                    dt1.Rows[dt1.Rows.Count - 1][description] = val2;
                                                }
                                            }
                                        }
                                    }
                                }

                                if (comboBox_transition.Text == bn)
                                {
                                    dt2.Rows.Add();
                                    dt2.Rows[dt2.Rows.Count - 1][objectid] = bl1.ObjectId.Handle.Value.ToString();
                                    dt2.Rows[dt2.Rows.Count - 1][blockname] = bn;
                                    Point3d ptins = bl1.Position;
                                    dt2.Rows[dt2.Rows.Count - 1][X] = Math.Round(ptins.X, 3);
                                    dt2.Rows[dt2.Rows.Count - 1][Y] = Math.Round(ptins.Y, 3);
                                }
                            }
                        }

                        dt1 = Functions.Sort_data_table_2_columns(dt1, STA1, blockname);
                        dt2 = Functions.Sort_data_table_2_columns(dt2, Y, X);
                        if (dt1 != null && dt1.Rows.Count > 1)
                        {

                            string mat1 = "";
                            int index1 = -1;
                            if (dt1.Rows[0][mat] != DBNull.Value)
                            {
                                mat1 = Convert.ToString(dt1.Rows[0][mat]);
                                if (dt_t.Columns.Contains(mat1) == false)
                                {
                                    mat1 = "";
                                }
                                else
                                {
                                    index1 = dt_t.Columns.IndexOf(mat1);
                                }
                            }
                            for (int i = 1; i < dt1.Rows.Count; i++)
                            {
                                string mat2 = "";
                                if (dt1.Rows[i][mat] != DBNull.Value)
                                {
                                    mat2 = Convert.ToString(dt1.Rows[i][mat]);
                                }
                                if (mat1 != "" && mat2 != "")
                                {
                                    if (mat1 != mat2)
                                    {
                                        int index2 = -1;
                                        if (dt_t.Columns.Contains(mat2) == false)
                                        {
                                            mat2 = "";
                                        }
                                        else
                                        {
                                            index2 = dt_t.Columns.IndexOf(mat2);
                                        }
                                        if (index1 != -1 && index2 != -1 && dt_t.Rows[index1][index2] != DBNull.Value)
                                        {
                                            double x2 = Convert.ToDouble(dt1.Rows[i][X]);
                                            for (int j = 0; j < dt2.Rows.Count; j++)
                                            {
                                                double x3 = Convert.ToDouble(dt2.Rows[j][X]);
                                                string Tassigned = "";
                                                if (dt2.Rows[j][hasT] != DBNull.Value)
                                                {
                                                    Tassigned = Convert.ToString(dt2.Rows[j][hasT]);
                                                }

                                                if (x2 == x3 && Tassigned == "")
                                                {
                                                    dt1.Rows[i][hasT] = "T";
                                                    dt2.Rows[j][hasT] = "YES";
                                                }
                                            }
                                        }
                                        mat1 = mat2;
                                        index1 = index2;
                                    }
                                }
                            }

                            for (int j = 0; j < dt2.Rows.Count; j++)
                            {
                                System.Data.DataRow row1 = dt1.NewRow();
                                row1.ItemArray = dt2.Rows[j].ItemArray;
                                dt1.Rows.Add(row1);
                            }

                            for (int j = 0; j < dt2.Rows.Count; j++)
                            {
                                if (dt2.Rows[j][hasT] == DBNull.Value)
                                {
                                    ObjectId id1 = Functions.GetObjectId(ThisDrawing.Database, Convert.ToString(dt2.Rows[j][objectid]));
                                    BlockReference bl1 = Trans1.GetObject(id1, OpenMode.ForWrite) as BlockReference;
                                    if (bl1 != null)
                                    {
                                        bl1.ScaleFactors = new Scale3d(10, 10, 10);
                                    }
                                }
                            }

                            if (dt1.Rows[0][mat] != DBNull.Value)
                            {
                                mat1 = Convert.ToString(dt1.Rows[0][mat]);
                            }
                            else
                            {
                                mat1 = "";
                            }


                        }
                        Trans1.Commit();

                        string nume1 = System.DateTime.Today.Year + "-" + System.DateTime.Today.Month + "-" + System.DateTime.Today.Day + "_" + System.DateTime.Now.Hour + "hr" + System.DateTime.Now.Minute + "min" + System.DateTime.Now.Second + "sec";
                        List<string> lista_col = new List<string>();
                        List<double> lista_width = new List<double>();
                        lista_col.Add("A");
                        lista_width.Add(9);
                        lista_col.Add("B");
                        lista_width.Add(12);
                        lista_col.Add("C");
                        lista_width.Add(10);
                        lista_col.Add("D");
                        lista_width.Add(10);
                        lista_col.Add("E");
                        lista_width.Add(9);
                        lista_col.Add("F");
                        lista_width.Add(5);
                        lista_col.Add("G");
                        lista_width.Add(9);
                        lista_col.Add("H");
                        lista_width.Add(9);
                        lista_col.Add("I");
                        lista_width.Add(9);
                        lista_col.Add("J");
                        lista_width.Add(70);
                        Functions.Transfer_datatable_to_new_excel_spreadsheet(dt1, nume1, lista_col, lista_width);
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

        private void comboBox_xl_DropDown(object sender, EventArgs e)
        {
            ComboBox combo1 = sender as ComboBox;
            Functions.Load_opened_workbooks_to_combobox(combo1);
            combo1.DropDownWidth = Functions.get_dropdown_width(combo1);
        }

        private void button_load_transition_Click(object sender, EventArgs e)
        {
            string filename = comboBox_xl_canada.Text;
            dt_t = new System.Data.DataTable();
            string tab_transition = "Transition Table";
            try
            {
                if (filename.Length > 0)
                {
                    set_enable_false();
                    Microsoft.Office.Interop.Excel.Worksheet W11 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, tab_transition);
                    if (W11 != null)
                    {
                        string h1a = "pipe";
                        string h1b = "type";
                        string value1 = W11.Range["E2"].Value2;
                        if (value1 == null) value1 = "";

                        value1 = value1.ToLower();
                        if (value1.Contains(h1a) == false || value1.Contains(h1b) == false)
                        {
                            MessageBox.Show(h1a + " " + h1b + " is not found. Check E2 on " + tab_transition + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }

                        object[,] values1 = new object[1, 21];
                        values1 = W11.Range["F2:Z2"].Value2;
                        for (int j = 1; j <= values1.Length; ++j)
                        {
                            object Valoare1 = values1[1, j];
                            if (Valoare1 != null)
                            {
                                string mat1 = Convert.ToString(Valoare1);
                                if (dt_t.Columns.Contains(mat1) == false)
                                {
                                    dt_t.Columns.Add(mat1, typeof(string));
                                    dt_t.Rows.Add();
                                }

                            }
                            else
                            {
                                j = values1.Length;
                            }
                        }

                        object[,] values2 = new object[dt_t.Rows.Count, dt_t.Columns.Count];
                        string last_col = Functions.get_excel_column_letter(5 + dt_t.Columns.Count);

                        values2 = W11.Range["F3:" + last_col + Convert.ToString(dt_t.Rows.Count + 2)].Value2;

                        for (int i = 1; i <= dt_t.Rows.Count; ++i)
                        {
                            for (int j = 1; j <= dt_t.Columns.Count; ++j)
                            {
                                object Valoare2 = values2[i, j];
                                if (Valoare2 != null)
                                {
                                    dt_t.Rows[i - 1][j - 1] = "T";

                                }

                            }
                        }
                        //Functions.Transfer_datatable_to_new_excel_spreadsheet(dt_t);
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                set_enable_true();
                return;
            }
            set_enable_true();
        }

        private void comboBox_block_to_be_placed_DropDown(object sender, EventArgs e)
        {
            Functions.Incarca_existing_Blocks_to_combobox(comboBox_block_to_be_placed);
        }

        private void comboBox_block_to_be_placed_SelectedIndexChanged(object sender, EventArgs e)
        {
            Functions.Incarca_existing_Atributes_to_combobox(comboBox_block_to_be_placed.Text, comboBox_atr);
        }

        private void button_place_blocks_on_band_Click(object sender, EventArgs e)
        {
            if (Functions.IsNumeric(textBox_deltaY.Text) == false)
            {
                MessageBox.Show("No deltaY specified");
                return;
            }

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
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult rez1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez1 = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez1.MessageForAdding = "\nSelect the blocks:";
                        Prompt_rez1.SingleOnly = false;
                        this.MdiParent.WindowState = FormWindowState.Minimized;

                        rez1 = ThisDrawing.Editor.GetSelection(Prompt_rez1);

                        if (rez1.Status != PromptStatus.OK)
                        {
                            this.MdiParent.WindowState = FormWindowState.Normal;
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        System.Data.DataTable dt1 = new System.Data.DataTable();
                        dt1.Columns.Add("objectid", typeof(string));
                        dt1.Columns.Add("sta", typeof(double));



                        for (int i = 0; i < rez1.Value.Count; i++)
                        {
                            BlockReference block1 = Trans1.GetObject(rez1.Value[i].ObjectId, OpenMode.ForRead) as BlockReference;
                            if (block1 != null)
                            {
                                string nume_block = Functions.get_block_name(block1);
                                if (nume_block == comboBox_block_to_be_placed.Text)
                                {
                                    if (block1.AttributeCollection.Count > 0)
                                    {
                                        block1.UpgradeOpen();
                                        foreach (ObjectId id1 in block1.AttributeCollection)
                                        {
                                            AttributeReference atr1 = Trans1.GetObject(id1, OpenMode.ForRead) as AttributeReference;
                                            if (atr1 != null)
                                            {
                                                if (atr1.Tag == comboBox_atr.Text)
                                                {
                                                    string val_str = atr1.TextString;
                                                    if (Functions.IsNumeric(val_str.Replace("+", "")) == true)
                                                    {
                                                        double val1 = Convert.ToDouble(val_str.Replace("+", ""));
                                                        dt1.Rows.Add();
                                                        dt1.Rows[dt1.Rows.Count - 1][0] = block1.ObjectId.Handle.Value.ToString();
                                                        dt1.Rows[dt1.Rows.Count - 1][1] = val1;
                                                    }
                                                }
                                            }

                                        }
                                    }
                                }
                            }
                        }

                        if (dt1.Rows.Count > 0)
                        {

                            Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rez.MessageForAdding = "\nSelect the bands:";
                            Prompt_rez.SingleOnly = false;
                            Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                            if (Rezultat1.Status != PromptStatus.OK)
                            {
                                this.MdiParent.WindowState = FormWindowState.Normal;
                                dt1 = null;
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }

                            for (int i = 0; i < Rezultat1.Value.Count; i++)
                            {
                                BlockReference block1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as BlockReference;
                                if (block1 != null)
                                {
                                    //string nume_block = Functions.get_block_name(block1);
                                    //if (nume_block == comboBox_block_to_be_placed.Text)
                                    {
                                        if (block1.AttributeCollection.Count > 0)
                                        {
                                            block1.UpgradeOpen();
                                            double sta1 = 0;
                                            double sta2 = 0;

                                            foreach (ObjectId id1 in block1.AttributeCollection)
                                            {
                                                AttributeReference atr1 = Trans1.GetObject(id1, OpenMode.ForRead) as AttributeReference;
                                                if (atr1 != null)
                                                {
                                                    if (atr1.Tag == "STA1")
                                                    {
                                                        string val_str = atr1.TextString;
                                                        if (Functions.IsNumeric(val_str.Replace("+", "")) == true)
                                                        {
                                                            sta1 = Convert.ToDouble(val_str.Replace("+", ""));

                                                        }
                                                    }
                                                    if (atr1.Tag == "STA2")
                                                    {
                                                        string val_str = atr1.TextString;
                                                        if (Functions.IsNumeric(val_str.Replace("+", "")) == true)
                                                        {
                                                            sta2 = Convert.ToDouble(val_str.Replace("+", ""));

                                                        }
                                                    }
                                                }

                                            }
                                            if (sta1 != sta2)
                                            {
                                                double DeltaY = Convert.ToDouble(textBox_deltaY.Text);
                                                for (int j = dt1.Rows.Count - 1; j >= 0; --j)
                                                {
                                                    double sta = Convert.ToDouble(dt1.Rows[j][1]);
                                                    if (sta >= sta1 && sta <= sta2)
                                                    {
                                                        ObjectId id1 = Functions.GetObjectId(ThisDrawing.Database, Convert.ToString(dt1.Rows[j][0]));
                                                        BlockReference T = Trans1.GetObject(id1, OpenMode.ForWrite) as BlockReference;
                                                        if (T != null)
                                                        {
                                                            double dist = Functions.Get_Param_Value_block(block1, "Distance1");
                                                            double deltaX = ((sta - sta1) * dist) / (sta2 - sta1);
                                                            if (radioButton_right_left.Checked == true)
                                                            {
                                                                T.Position = new Point3d(block1.Position.X - deltaX, block1.Position.Y + DeltaY, 0);
                                                            }
                                                            else
                                                            {
                                                                T.Position = new Point3d(block1.Position.X + deltaX, block1.Position.Y + DeltaY, 0);
                                                            }
                                                        }
                                                        dt1.Rows[j].Delete();
                                                    }
                                                }
                                            }
                                        }
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
            set_enable_true();
            this.MdiParent.WindowState = FormWindowState.Normal;


        }
    }
}
