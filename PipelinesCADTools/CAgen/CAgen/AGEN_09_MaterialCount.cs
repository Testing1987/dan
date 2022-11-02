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
    public partial class AGEN_MaterialCount : Form
    {

        public AGEN_MaterialCount()
        {
            InitializeComponent();
        }

        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_calculate_totals);
            lista_butoane.Add(button_show_mat_draw);



            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {

                bt1.Enabled = false;

            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_calculate_totals);
            lista_butoane.Add(button_show_mat_draw);



            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }


        private void button_show_mat_draw_Click(object sender, EventArgs e)
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

            _AGEN_mainform.tpage_owner_draw.Hide();
            _AGEN_mainform.tpage_mat.Hide();
            _AGEN_mainform.tpage_cust_scan.Hide();
            _AGEN_mainform.tpage_cust_draw.Hide();
            _AGEN_mainform.tpage_sheet_gen.Hide();
          

            _AGEN_mainform.tpage_owner_scan.Hide();
            _AGEN_mainform.tpage_mat_count.Hide();
            _AGEN_mainform.tpage_mat.Show();


        }



        private void button_calculate_totals_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dtmc = new System.Data.DataTable();
            dtmc.Columns.Add("band", typeof(string));

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
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;

                        this.MdiParent.WindowState = FormWindowState.Minimized;

                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez.MessageForAdding = "\nSelect the items:";
                        Prompt_rez.SingleOnly = false;
                        Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            this.MdiParent.WindowState = FormWindowState.Normal;
                            set_enable_true();
                            ThisDrawing.Editor.WriteMessage("\nCommand:");
                            return;
                        }

                        this.MdiParent.WindowState = FormWindowState.Normal;
                        List<ObjectId> lista_processed = new List<ObjectId>();
                        if (Rezultat1.Value.Count > 0)
                        {
                            for (int i = 0; i < Rezultat1.Value.Count; ++i)
                            {
                                Polyline rect_green = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as Polyline;
                                if (rect_green != null && rect_green.Layer == "Agen_no_plot_mat" && rect_green.ColorIndex == 3)
                                {
                                    lista_processed.Add(rect_green.ObjectId);
                                    for (int j = 0; j < Rezultat1.Value.Count; ++j)
                                    {
                                        ObjectId id2 = Rezultat1.Value[j].ObjectId;
                                        if (lista_processed.Contains(id2) == false)
                                        {
                                            MText label1 = Trans1.GetObject(id2, OpenMode.ForRead) as MText;
                                            if (label1 != null && label1.Layer == "Agen_no_plot_mat")
                                            {
                                                if (Math.Abs(label1.Location.X - rect_green.GetPointAtParameter(0).X) < 1 && label1.Location.Y < rect_green.GetPointAtParameter(0).Y && label1.Location.Y > rect_green.GetPointAtParameter(3).Y)
                                                {
                                                    lista_processed.Add(id2);
                                                    dtmc.Rows.Add();
                                                    dtmc.Rows[dtmc.Rows.Count - 1][0] = label1.Text;
                                                    string band = label1.Text;
                                                    for (int k = 0; k < Rezultat1.Value.Count; ++k)
                                                    {
                                                        ObjectId id3 = Rezultat1.Value[k].ObjectId;
                                                        BlockReference block1 = Trans1.GetObject(id3, OpenMode.ForRead) as BlockReference;

                                                        if (lista_processed.Contains(id3) == false && block1 != null)
                                                        {

                                                            double x0 = rect_green.GetPointAtParameter(0).X;
                                                            double x1 = rect_green.GetPointAtParameter(1).X;
                                                            double y0 = rect_green.GetPointAtParameter(0).Y;
                                                            double y3 = rect_green.GetPointAtParameter(3).Y;

                                                            if (block1.Position.X >= x0 && block1.Position.X <= x1 &&
                                                            block1.Position.Y >= y3-0.5 && block1.Position.Y <= y0)
                                                            {

                                                                string blockname = Functions.get_block_name(block1);
                                                                if (block1.AttributeCollection.Count > 0)
                                                                {
                                                                    Autodesk.AutoCAD.DatabaseServices.AttributeCollection attColl = block1.AttributeCollection;
                                                                    bool contains_mat = false;
                                                                    bool contains_len = false;
                                                                    string col_string = "XXX";

                                                                    for (int m = 0; m < attColl.Count; ++m)
                                                                    {
                                                                        DBObject ent = Trans1.GetObject(attColl[m], Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                                                                        if (ent is AttributeReference)
                                                                        {
                                                                            AttributeReference atr1 = ent as AttributeReference;
                                                                            if (atr1 != null)
                                                                            {
                                                                                if (atr1.Tag.ToUpper() == "MAT")
                                                                                {
                                                                                    col_string = Convert.ToString(atr1.TextString);
                                                                                    if (dtmc.Columns.Contains(col_string) == false)
                                                                                    {
                                                                                        dtmc.Columns.Add(col_string, typeof(double));
                                                                                    }
                                                                                    contains_mat = true;
                                                                                    m = attColl.Count;
                                                                                }
                                                                            }
                                                                        }
                                                                    }

                                                                    if (contains_mat == true)
                                                                    {
                                                                        for (int m = 0; m < attColl.Count; ++m)
                                                                        {
                                                                            DBObject ent = Trans1.GetObject(attColl[m], Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                                                                            if (ent is AttributeReference)
                                                                            {
                                                                                AttributeReference atr1 = ent as AttributeReference;
                                                                                if (atr1 != null)
                                                                                {
                                                                                    if (atr1.Tag.ToUpper() == "LEN" || atr1.Tag.ToUpper() == "LENGTH" || atr1.Tag.ToUpper() == "QTY")
                                                                                    {
                                                                                        string val_string = Convert.ToString(atr1.TextString).Replace("'", "");

                                                                                        contains_len = true;
                                                                                        if (Functions.IsNumeric(val_string) == true)
                                                                                        {
                                                                                            double ex_value = 0;
                                                                                            double new_value = Math.Round(Convert.ToDouble(val_string),2);
                                                                                            if (dtmc.Rows[dtmc.Rows.Count - 1][col_string] != DBNull.Value) ex_value = Convert.ToDouble(dtmc.Rows[dtmc.Rows.Count - 1][col_string]);
                                                                                            dtmc.Rows[dtmc.Rows.Count - 1][col_string] = Math.Round(ex_value + new_value,2);
                                                                                        }
                                                                                        else
                                                                                        {
                                                                                            double ex_value = 0;
                                                                                            if (dtmc.Rows[dtmc.Rows.Count - 1][col_string] != DBNull.Value) ex_value = Convert.ToDouble(dtmc.Rows[dtmc.Rows.Count - 1][col_string]);
                                                                                            dtmc.Rows[dtmc.Rows.Count - 1][col_string] = ++ex_value;
                                                                                        }
                                                                                        m = attColl.Count;
                                                                                    }
                                                                                }
                                                                            }
                                                                        }

                                                                        if (contains_len == false)
                                                                        {

                                                                            bool has_sta1 = false;
                                                                            bool has_sta2 = false;

                                                                            double sta1 = 0;
                                                                            double sta2 = 1;

                                                                            for (int m = 0; m < attColl.Count; ++m)
                                                                            {
                                                                                DBObject ent = Trans1.GetObject(attColl[m], Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                                                                                if (ent is AttributeReference)
                                                                                {
                                                                                    AttributeReference atr1 = ent as AttributeReference;
                                                                                    if (atr1 != null)
                                                                                    {
                                                                                        if (atr1.Tag.ToUpper() == "STA1")
                                                                                        {
                                                                                            string val_string = Convert.ToString(atr1.TextString).Replace("+", "");
                                                                                            has_sta1 = true;
                                                                                            if (Functions.IsNumeric(val_string) == true)
                                                                                            {
                                                                                                sta1 = Convert.ToDouble(val_string);
                                                                                                has_sta1 = true;
                                                                                            }
                                                                                        }

                                                                                        if (atr1.Tag.ToUpper() == "STA2")
                                                                                        {
                                                                                            string val_string = Convert.ToString(atr1.TextString).Replace("+", "");
                                                                                            has_sta2 = true;
                                                                                            if (Functions.IsNumeric(val_string) == true)
                                                                                            {
                                                                                                sta2 = Convert.ToDouble(val_string);
                                                                                                has_sta2 = true;
                                                                                            }
                                                                                        }
                                                                                    }
                                                                                }
                                                                            }

                                                                            if (has_sta2 == true && has_sta1 == true)
                                                                            {

                                                                                double ex_value = 0;
                                                                                if (dtmc.Rows[dtmc.Rows.Count - 1][col_string] != DBNull.Value) ex_value = Convert.ToDouble(dtmc.Rows[dtmc.Rows.Count - 1][col_string]);
                                                                                dtmc.Rows[dtmc.Rows.Count - 1][col_string] = Math.Round(ex_value + Math.Round(sta2,2) - Math.Round(sta1,2),2);
                                                                            }
                                                                            else
                                                                            {

                                                                                double ex_value = 0;
                                                                                if (dtmc.Rows[dtmc.Rows.Count - 1][col_string] != DBNull.Value) ex_value = Convert.ToDouble(dtmc.Rows[dtmc.Rows.Count - 1][col_string]);
                                                                                dtmc.Rows[dtmc.Rows.Count - 1][col_string] = ++ex_value;
                                                                            }
                                                                        }
                                                                    }
                                                                    else
                                                                    {
                                                                        if (dtmc.Columns.Contains(blockname) == false)
                                                                        {
                                                                            dtmc.Columns.Add(blockname, typeof(double));
                                                                        }

                                                                        for (int m = 0; m < attColl.Count; ++m)
                                                                        {
                                                                            DBObject ent = Trans1.GetObject(attColl[m], Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                                                                            if (ent is AttributeReference)
                                                                            {
                                                                                AttributeReference atr1 = ent as AttributeReference;
                                                                                if (atr1 != null)
                                                                                {
                                                                                    if (atr1.Tag.ToUpper() == "LEN" || atr1.Tag.ToUpper() == "LENGTH" || atr1.Tag.ToUpper() == "QTY")
                                                                                    {
                                                                                        string val_string = Convert.ToString(atr1.TextString).Replace("'", "");

                                                                                        contains_len = true;
                                                                                        if (Functions.IsNumeric(val_string) == true)
                                                                                        {
                                                                                            double ex_value = 0;
                                                                                            double new_value = Math.Round(Convert.ToDouble(val_string),2);
                                                                                            if (dtmc.Rows[dtmc.Rows.Count - 1][blockname] != DBNull.Value) ex_value = Convert.ToDouble(dtmc.Rows[dtmc.Rows.Count - 1][blockname]);
                                                                                            dtmc.Rows[dtmc.Rows.Count - 1][blockname] = Math.Round(ex_value + new_value,2);
                                                                                        }
                                                                                        else
                                                                                        {
                                                                                            double ex_value = 0;
                                                                                            if (dtmc.Rows[dtmc.Rows.Count - 1][blockname] != DBNull.Value) ex_value = Convert.ToDouble(dtmc.Rows[dtmc.Rows.Count - 1][blockname]);
                                                                                            dtmc.Rows[dtmc.Rows.Count - 1][blockname] = ++ex_value;
                                                                                        }

                                                                                        m = attColl.Count;
                                                                                    }
                                                                                }
                                                                            }
                                                                        }

                                                                        if (contains_len == false)
                                                                        {
                                                                            bool has_sta1 = false;
                                                                            bool has_sta2 = false;
                                                                            double sta1 = 0;
                                                                            double sta2 = 1;

                                                                            for (int m = 0; m < attColl.Count; ++m)
                                                                            {
                                                                                DBObject ent = Trans1.GetObject(attColl[m], Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                                                                                if (ent is AttributeReference)
                                                                                {
                                                                                    AttributeReference atr1 = ent as AttributeReference;
                                                                                    if (atr1 != null)
                                                                                    {
                                                                                        if (atr1.Tag.ToUpper() == "STA1")
                                                                                        {
                                                                                            string val_string = Convert.ToString(atr1.TextString).Replace("+", "");
                                                                                            has_sta1 = true;
                                                                                            if (Functions.IsNumeric(val_string) == true)
                                                                                            {
                                                                                                sta1 = Convert.ToDouble(val_string);
                                                                                                has_sta1 = true;
                                                                                            }
                                                                                        }

                                                                                        if (atr1.Tag.ToUpper() == "STA2")
                                                                                        {
                                                                                            string val_string = Convert.ToString(atr1.TextString).Replace("+", "");
                                                                                            has_sta2 = true;
                                                                                            if (Functions.IsNumeric(val_string) == true)
                                                                                            {
                                                                                                sta2 = Convert.ToDouble(val_string);
                                                                                                has_sta2 = true;
                                                                                            }
                                                                                        }
                                                                                    }
                                                                                }
                                                                            }

                                                                            if (has_sta2 == true && has_sta1 == true)
                                                                            {
                                                                                double ex_value = 0;
                                                                                if (dtmc.Rows[dtmc.Rows.Count - 1][blockname] != DBNull.Value) ex_value = Convert.ToDouble(dtmc.Rows[dtmc.Rows.Count - 1][blockname]);
                                                                                dtmc.Rows[dtmc.Rows.Count - 1][blockname] = Math.Round(ex_value + Math.Round(sta2,2) - Math.Round(sta1,2),2);
                                                                            }
                                                                            else
                                                                            {
                                                                                double ex_value = 0;
                                                                                if (dtmc.Rows[dtmc.Rows.Count - 1][blockname] != DBNull.Value) ex_value = Convert.ToDouble(dtmc.Rows[dtmc.Rows.Count - 1][blockname]);
                                                                                dtmc.Rows[dtmc.Rows.Count - 1][blockname] = ++ex_value;
                                                                            }

                                                                        }
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    if (dtmc.Columns.Contains(blockname) == false)
                                                                    {
                                                                        dtmc.Columns.Add(blockname, typeof(double));
                                                                    }
                                                                    double ex_value = 0;
                                                                    if (dtmc.Rows[dtmc.Rows.Count - 1][blockname] != DBNull.Value) ex_value = Convert.ToDouble(dtmc.Rows[dtmc.Rows.Count - 1][blockname]);
                                                                    dtmc.Rows[dtmc.Rows.Count - 1][blockname] = ++ex_value;
                                                                }
                                                                lista_processed.Add(id3);
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            if (dtmc != null && dtmc.Rows.Count > 0)
                            {
                                double[] count1 = new double[dtmc.Columns.Count - 1];
                                for (int i = 0; i < dtmc.Rows.Count; ++i)
                                {
                                    for (int j = 1; j < dtmc.Columns.Count; ++j)
                                    {
                                        if (dtmc.Rows[i][j] != DBNull.Value)
                                        {
                                            double ex = Math.Round(count1[j - 1],2);
                                            double new1 =Math.Round( Convert.ToDouble(dtmc.Rows[i][j]),2);
                                            count1[j - 1] = ex + new1;
                                        }
                                    }

                                }
                                dtmc.Rows.Add();
                                dtmc.Rows[dtmc.Rows.Count - 1][0] = "Total:";
                                for (int j = 1; j < dtmc.Columns.Count; ++j)
                                {
                                    dtmc.Rows[dtmc.Rows.Count - 1][j] = count1[j - 1];
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
            dtmc = Functions.Sort_data_table(dtmc, "band");
            if (dtmc.Columns.Count > 1)
            {
                List<string> lista1 = new List<string>();
                for (int j = 1; j < dtmc.Columns.Count ; ++j)
                {

                    lista1.Add(dtmc.Columns[j].ColumnName);
                }
                lista1.Sort();
                for (int i = 0; i < lista1.Count; ++i)
                {
                    dtmc.Columns[lista1[i]].SetOrdinal(dtmc.Columns.Count - 1);
                }

            }

            Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dtmc);
            dataGridView_mat_totals.DataSource = dtmc;
            dataGridView_mat_totals.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            dataGridView_mat_totals.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_mat_totals.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView_mat_totals.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_mat_totals.DefaultCellStyle.ForeColor = Color.White;
            dataGridView_mat_totals.EnableHeadersVisualStyles = false;
            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            set_enable_true();

        }

        private void button_counts2TBLK_Click(object sender, EventArgs e)
        {
            int start1 = 3;
            int end1 = 1000;
            string col_um = "units";
            string col_align = "Alignment";
            string tab_materials = "Materials";
            string col_pipe_type = "Pipe Type";
            string col_descr = "Description";

            System.Data.DataTable dt_materials = new System.Data.DataTable();
            dt_materials.Columns.Add(col_pipe_type, typeof(string));
            dt_materials.Columns.Add(col_descr, typeof(string));
            dt_materials.Columns.Add(col_um, typeof(string));

            string filename = comboBox_excel_files.Text;


            System.Data.DataTable dt_mat_counts = new System.Data.DataTable();
            string tab_matcounts = "MatCounts";
            List<string> lista_mat = new List<string>();
            try
            {


                if (filename.Length > 0)
                {
                    set_enable_false();
                    Microsoft.Office.Interop.Excel.Worksheet W1 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, tab_materials);
                    if (W1 != null)
                    {

                        string xl_pipe_type = "E";
                        string xl_description = "F";
                        string xl_um = "G";

                        List<string> lista_col = new List<string>();
                        List<string> lista_colxl = new List<string>();

                        lista_col.Add(col_pipe_type);
                        lista_col.Add(col_descr);
                        lista_col.Add(col_um);

                        lista_colxl.Add(xl_pipe_type);
                        lista_colxl.Add(xl_description);
                        lista_colxl.Add(xl_um);

                        dt_materials = Functions.build_data_table_from_excel(dt_materials, W1,
                                                start1, end1, lista_col, lista_colxl);

                        if (dt_materials.Rows.Count == 0)
                        {
                            MessageBox.Show("No data found in" + tab_materials + "\r\noperation aborted");
                            set_enable_true();
                            return;
                        }

                        dt_mat_counts.Columns.Add(col_align, typeof(string));
                        for (int i = 0; i < dt_materials.Rows.Count; ++i)
                        {
                            string mat1 = Convert.ToString(dt_materials.Rows[i][col_pipe_type]);
                            if (lista_mat.Contains(mat1) == false)
                            {
                                lista_mat.Add(mat1);
                                dt_mat_counts.Columns.Add(mat1, typeof(double));
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show(tab_materials + " not found\r\noperation aborted");
                        set_enable_true();
                        return;
                    }



                    Microsoft.Office.Interop.Excel.Worksheet W2 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, tab_matcounts);
                    if (W2 != null)
                    {
                        start1 = 2;

                        List<string> lista_col = new List<string>();
                        List<string> lista_colxl = new List<string>();

                        int col = 1;

                        for (int i = 0; i < dt_mat_counts.Columns.Count; ++i)
                        {
                            lista_col.Add(dt_mat_counts.Columns[i].ColumnName);
                            lista_colxl.Add(Functions.get_excel_column_letter(col));
                            ++col;
                        }

                        dt_mat_counts = Functions.build_data_table_from_excel(dt_mat_counts, W2, start1, end1, lista_col, lista_colxl);

                    }
                    else
                    {
                        MessageBox.Show(tab_matcounts + " not found\r\noperation aborted");
                        set_enable_true();
                        return;
                    }

                    System.Data.DataTable dt1 = new System.Data.DataTable();
                    dt1.Columns.Add(col_align, typeof(string));

                    for (int i = 0; i < dt_materials.Rows.Count; ++i)
                    {
                        dt1.Columns.Add("MatNo" + (i + 1).ToString(), typeof(string));
                        dt1.Columns.Add("MatDescr" + (i + 1).ToString(), typeof(string));
                        dt1.Columns.Add("MatQty" + (i + 1).ToString(), typeof(double));
                        dt1.Columns.Add("MatUnits" + (i + 1).ToString(), typeof(string));
                    }

                    for (int i = 0; i < dt_mat_counts.Rows.Count; ++i)
                    {
                        if (dt_mat_counts.Rows[i][0] != DBNull.Value)
                        {
                            dt1.Rows.Add();
                            dt1.Rows[dt1.Rows.Count - 1][0] = dt_mat_counts.Rows[i][0];
                            for (int j = 1; j < dt_mat_counts.Columns.Count; ++j)
                            {
                                if (dt_mat_counts.Rows[i][j] != DBNull.Value)
                                {
                                    string mat1 = dt_mat_counts.Columns[j].ColumnName;

                                    double qty1 = Convert.ToDouble(dt_mat_counts.Rows[i][j]);
                                    string descr1 = "";
                                    string um1 = "";
                                    for (int k = 0; k < dt_materials.Rows.Count; ++k)
                                    {
                                        string mat2 = Convert.ToString(dt_materials.Rows[k][0]);
                                        if (mat1 == mat2)
                                        {
                                            descr1 = Convert.ToString(dt_materials.Rows[k][1]);
                                            um1 = Convert.ToString(dt_materials.Rows[k][2]);
                                            k = dt_materials.Rows.Count;
                                        }
                                    }

                                    for (int k = 1; k < dt1.Columns.Count; ++k)
                                    {
                                        if (dt1.Rows[dt1.Rows.Count - 1][k] == DBNull.Value)
                                        {
                                            dt1.Rows[dt1.Rows.Count - 1][k] = mat1;
                                            dt1.Rows[dt1.Rows.Count - 1][k + 1] = descr1;
                                            dt1.Rows[dt1.Rows.Count - 1][k + 2] = qty1;
                                            dt1.Rows[dt1.Rows.Count - 1][k + 3] = um1;
                                            k = dt1.Columns.Count;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    Functions.Transfer_datatable_to_new_excel_spreadsheet(dt1, "General");
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            set_enable_true();
        }

        private void comboBox_excel_files_DropDown(object sender, EventArgs e)
        {
            ComboBox combo1 = sender as ComboBox;
            Functions.Load_opened_workbooks_to_combobox(combo1);
            combo1.DropDownWidth = Functions.get_dropdown_width(combo1);
        }

        private void panel7_Click(object sender, EventArgs e)
        {
            if(comboBox_excel_files.Visible==false)
            {
                comboBox_excel_files.Visible = true;
                label_open_files_canada.Visible = true;
                button_counts2TBLK.Visible = true;
            }
            else
            {
                comboBox_excel_files.Visible = false;
                label_open_files_canada.Visible = false;
                button_counts2TBLK.Visible = false;


            }
        }
    }
}



