using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace Alignment_mdi
{
    public partial class wksp_tool
    {
        private void attach_dt_stg_to_datagridview(System.Data.DataTable dt1)
        {
            if (dt1 != null && dt1.Rows.Count > 0)
            {
                dataGridView_stg_data.DataSource = dt1;
                dataGridView_stg_data.Columns[stg_sta1_column].Width = 100;
                dataGridView_stg_data.Columns[stg_sta2_column].Width = 100;
                dataGridView_stg_data.Columns[stg_area_column].Width = 60;
                dataGridView_stg_data.Columns[stg_justification_column].Width = 250;
                dataGridView_stg_data.Columns[stg_justification_column].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                dataGridView_stg_data.Columns[stg_handle_column].Width = 2;

                dataGridView_stg_data.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_stg_data.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                dataGridView_stg_data.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Padding newpadding = new Padding(4, 0, 0, 0);
                dataGridView_stg_data.ColumnHeadersDefaultCellStyle.Padding = newpadding;
                dataGridView_stg_data.RowHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_stg_data.DefaultCellStyle.BackColor = Color.FromArgb(51, 51, 55);
                dataGridView_stg_data.DefaultCellStyle.ForeColor = Color.White;
                dataGridView_stg_data.EnableHeadersVisualStyles = false;
            }
            else
            {
                dataGridView_stg_data.DataSource = null;
            }
        }

        private void format_and_transfer_dt_stg_to_excel(Worksheet W10)
        {
            if (W10 != null && dt_stg != null && dt_stg.Rows.Count > 0)
            {
                W10.Range["A:C"].ColumnWidth = 15;
                W10.Range["D:D"].ColumnWidth = 85;
                W10.Range["E:E"].ColumnWidth = 12;
                W10.Range["A1:E1"].VerticalAlignment = XlVAlign.xlVAlignCenter;
                W10.Range["A1:E1"].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                W10.Range["A1:E30000"].ClearContents();
                W10.Range["A2:E" + Convert.ToString(1 + dt_stg.Rows.Count)].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                W10.Range["A2:E" + Convert.ToString(1 + dt_stg.Rows.Count)].VerticalAlignment = XlVAlign.xlVAlignCenter;
                Range range1 = W10.Range["A1:E1"];
                Functions.Color_border_range_inside(range1, 41); //blue
                range1.Font.ColorIndex = 2;
                range1.Font.Size = 11;
                range1.Font.Bold = true;

                Functions.Transfer_datatable_to_excel_spreadsheet(W10, dt_stg, 1, true);
                range1 = W10.Range["A2:B" + Convert.ToString(dt_stg.Rows.Count + 1)];
                Functions.Color_border_range_inside(range1, 43); //light green
                range1.Font.ColorIndex = 1;//black
                range1.Font.Size = 11;
                range1.Font.Bold = true;
                range1 = W10.Range["C2:C" + Convert.ToString(dt_stg.Rows.Count + 1)];
                Functions.Color_border_range_inside(range1, 44); //orange
                range1.Font.ColorIndex = 1;//black
                range1.Font.Size = 11;
                range1.Font.Bold = true;
                range1 = W10.Range["D2:D" + Convert.ToString(dt_stg.Rows.Count + 1)];
                Functions.Color_border_range_inside(range1, 43); //light green
                range1.Font.ColorIndex = 1;//black
                range1.Font.Size = 11;
                range1.Font.Bold = true;
                range1 = W10.Range["E2:E" + Convert.ToString(dt_stg.Rows.Count + 1)];
                Functions.Color_border_range_inside(range1, 44); //orange
                range1.Font.ColorIndex = 1;//black
                range1.Font.Size = 11;
                range1.Font.Bold = true;
            }
            else
            {
                W10.Range["A1:D30000"].ClearContents();
            }
        }


        private System.Data.DataTable build_dt_stg_from_config_excel(Worksheet W1)
        {
            System.Data.DataTable dt1 = get_dt_stg_structure();
            string Col1 = "E";

            Range range2 = W1.Range[Col1 + "2:" + Col1 + "30002"];
            object[,] values2 = new object[30000, 1];
            values2 = range2.Value2;

            bool is_data = false;
            for (int i = 1; i <= values2.Length; ++i)
            {
                object Valoare2 = values2[i, 1];
                if (Valoare2 != null)
                {
                    dt1.Rows.Add();
                    is_data = true;
                }
                else
                {
                    i = values2.Length + 1;
                }
            }

            if (is_data == false)
            {
                return null;
            }

            int NrR = dt1.Rows.Count;

            Range range1 = W1.Range["A2:E" + Convert.ToString(NrR + 1)];
            object[,] values = new object[NrR, 5];
            values = range1.Value2;

            for (int i = 0; i < dt1.Rows.Count; ++i)
            {
                for (int j = 0; j < dt1.Columns.Count; ++j)
                {
                    object val = values[i + 1, j + 1];
                    if (val == null) val = DBNull.Value;

                    dt1.Rows[i][j] = val;
                }
            }

            return dt1;
        }






        private void button_stg_select_drafted_Click(object sender, EventArgs e)
        {
            if (checkBox_use_od.Checked == true)
            {
                if (textBox_stg_justification.Text == "")
                {
                    MessageBox.Show("no justification specified\r\noperation aborted");
                    return;
                }
            }
            if (dt_cl == null || dt_cl.Rows.Count < 2)
            {
                MessageBox.Show("no centerline loaded\r\noperation aborted");
                return;
            }

            ObjectId[] Empty_array = null;
            System.Data.DataTable dt_new_stg = new System.Data.DataTable();
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {
                set_enable_false();
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    this.WindowState = FormWindowState.Minimized;

                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult rezultat_stg;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions prompt_stg;
                        prompt_stg = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the staging area polyline:");
                        prompt_stg.SetRejectMessage("\nSelect a polyline!");
                        prompt_stg.AllowNone = true;
                        prompt_stg.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        rezultat_stg = ThisDrawing.Editor.GetEntity(prompt_stg);

                        if (rezultat_stg.Status != PromptStatus.OK)
                        {
                            set_enable_true();
                            this.WindowState = FormWindowState.Normal;

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Polyline poly1 = null;
                        try
                        {
                            poly1 = Trans1.GetObject(rezultat_stg.ObjectId, OpenMode.ForRead) as Polyline;
                            dt_new_stg = build_data_table_from_poly(poly1);
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception)
                        {


                        }




                        Trans1.Commit();
                    }

                    if (dt_new_stg != null && dt_new_stg.Rows.Count > 0 && comboBox_layer_stg.Text != "")
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                            Functions.Creaza_layer(comboBox_layer_stg.Text, 3, true);
                            List<ObjectId> lista_od_stg_object_id = new List<ObjectId>();
                            List<string> lista_od_stg_justif = new List<string>();

                            Polyline polyCL = new Polyline();
                            for (int i = 0; i < dt_cl.Rows.Count; ++i)
                            {
                                polyCL.AddVertexAt(i, (Point2d)dt_cl.Rows[i][0], 0, 0, 0);
                            }
                            polyCL.Elevation = 0;

                            Polyline poly_stg = new Polyline();
                            if (dt_new_stg != null && dt_new_stg.Rows.Count > 0)
                            {
                                for (int i = 0; i < dt_new_stg.Rows.Count; ++i)
                                {
                                    poly_stg.AddVertexAt(i, (Point2d)dt_new_stg.Rows[i][0], 0, 0, 0);
                                }
                            }

                            if (poly_stg.Length > 0.01)
                            {
                                poly_stg.Layer = comboBox_layer_stg.Text;
                                poly_stg.ColorIndex = 256;
                                poly_stg.Linetype = "ByLayer";
                                poly_stg.Closed = true;
                                BTrecord.AppendEntity(poly_stg);
                                Trans1.AddNewlyCreatedDBObject(poly_stg, true);

                                using (DrawOrderTable DrawOrderTable1 = Trans1.GetObject(BTrecord.DrawOrderTableId, OpenMode.ForWrite) as DrawOrderTable)
                                {
                                    ObjectIdCollection col1 = new ObjectIdCollection();
                                    col1.Add(poly_stg.ObjectId);
                                    DrawOrderTable1.MoveToTop(col1);
                                }
                                string handle1 = poly_stg.ObjectId.Handle.Value.ToString();

                                if (dt_stg == null)
                                {
                                    dt_stg = get_dt_stg_structure();
                                }

                                double sta_min = 1000000000;
                                double sta_max = -1;
                                double d_min = 1000000000;

                                System.Data.DataTable dt1 = new System.Data.DataTable();
                                dt1.Columns.Add("sta", typeof(double));
                                dt1.Columns.Add("dist", typeof(double));

                                for (int k = 0; k < poly_stg.NumberOfVertices; ++k)
                                {
                                    Point3d pt1 = poly_stg.GetPointAtParameter(k);
                                    pt1 = new Point3d(pt1.X, pt1.Y, 0);

                                    Point3d pt2 = polyCL.GetClosestPointTo(pt1, Vector3d.ZAxis, false);
                                    pt2 = new Point3d(pt2.X, pt2.Y, 0);

                                    double d1 = Math.Pow(Math.Pow(pt1.X - pt2.X, 2) + Math.Pow(pt1.Y - pt2.Y, 2), 0.5);
                                    double sta1 = polyCL.GetDistAtPoint(pt2);

                                    dt1.Rows.Add();
                                    dt1.Rows[dt1.Rows.Count - 1][0] = sta1;
                                    dt1.Rows[dt1.Rows.Count - 1][1] = d1;
                                }

                                for (int k = 0; k < dt1.Rows.Count; ++k)
                                {
                                    double d1 = Convert.ToDouble(dt1.Rows[k][1]);
                                    if (d1 < d_min)
                                    {
                                        d_min = d1;
                                    }
                                }

                                for (int k = 0; k < dt1.Rows.Count; ++k)
                                {
                                    double d1 = Convert.ToDouble(dt1.Rows[k][1]);
                                    if (Math.Round(d1, 2) == Math.Round(d_min, 2))
                                    {
                                        double sta1 = Math.Round(Convert.ToDouble(dt1.Rows[k][0]), 2);
                                        if (sta1 < sta_min)
                                        {
                                            sta_min = sta1;
                                        }
                                        if (sta1 > sta_max)
                                        {
                                            sta_max = sta1;
                                        }
                                    }
                                }
                                dt1 = null;


                                bool is_found = false;
                                int index1 = -1;
                                for (int k = 0; k < dt_stg.Rows.Count; ++k)
                                {
                                    if (dt_stg.Rows[k][stg_handle_column] != DBNull.Value)
                                    {
                                        string handle2 = Convert.ToString(dt_stg.Rows[k][stg_handle_column]);
                                        if (handle1.ToLower() == handle2.ToLower())
                                        {
                                            is_found = true;
                                            index1 = k;
                                        }
                                    }
                                }

                                if (dt_manual_stg == null)
                                {
                                    dt_manual_stg = new System.Data.DataTable();
                                }

                                if (dt_manual_stg.Columns.Contains(handle1) == false) dt_manual_stg.Columns.Add(handle1, typeof(Point2d));

                                for (int n = 0; n < poly_stg.NumberOfVertices; ++n)
                                {
                                    if (dt_manual_stg.Rows.Count < n + 1)
                                    {
                                        dt_manual_stg.Rows.Add();
                                    }
                                    dt_manual_stg.Rows[n][handle1] = poly_stg.GetPoint2dAt(n);
                                }

                                lista_od_stg_object_id.Add(poly_stg.ObjectId);
                                lista_od_stg_justif.Add(textBox_stg_justification.Text);

                                if (is_found == false)
                                {
                                    dt_stg.Rows.Add();
                                    index1 = dt_stg.Rows.Count - 1;
                                }

                                dt_stg.Rows[index1][stg_sta1_column] = sta_min;
                                dt_stg.Rows[index1][stg_sta2_column] = sta_max;
                                dt_stg.Rows[index1][stg_justification_column] = textBox_stg_justification.Text;
                                dt_stg.Rows[index1][stg_area_column] = poly_stg.Area;
                                dt_stg.Rows[index1][stg_handle_column] = handle1;
                            }


                            attach_od_to_stg(lista_od_stg_object_id, lista_od_stg_justif);
                            attach_dt_stg_to_datagridview(dt_stg);
                            Trans1.Commit();
                        }
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
            this.WindowState = FormWindowState.Normal;

        }

        private void button_zoom_to_stg_Click(object sender, EventArgs e)
        {
            try
            {
                if (dt_stg != null && dt_stg.Rows.Count > 0)
                {
                    set_enable_false();
                    int row_idx = dataGridView_stg_data.SelectedCells[0].RowIndex;
                    if (row_idx >= 0)
                    {
                        string handle1 = Convert.ToString(dataGridView_stg_data.Rows[row_idx].Cells[stg_handle_column].Value);

                        Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                        using (DocumentLock lock1 = ThisDrawing.LockDocument())
                        {
                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                            {
                                ObjectId id1 = Functions.GetObjectId(ThisDrawing.Database, handle1);
                                if (id1 != ObjectId.Null)
                                {
                                    Functions.zoom_to_object(id1);
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
            set_enable_true();
        }

        private void button_stg_update_justification_Click(object sender, EventArgs e)
        {
            if (dt_stg == null || dt_stg.Rows.Count == 0)
            {
                MessageBox.Show("no stg data present\r\noperation abborted");
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

                        List<ObjectId> lista_od_stg_object_id = new List<ObjectId>();
                        List<string> lista_od_stg_justif = new List<string>();



                        for (int i = 0; i < dt_stg.Rows.Count; ++i)
                        {
                            if (dt_stg.Rows[i][stg_handle_column] != DBNull.Value && dt_stg.Rows[i][stg_justification_column] != DBNull.Value)
                            {
                                string handle1 = Convert.ToString(dt_stg.Rows[i][stg_handle_column]);
                                string justif1 = Convert.ToString(dt_stg.Rows[i][stg_justification_column]);

                                ObjectId id1 = Functions.GetObjectId(ThisDrawing.Database, handle1);

                                if (id1 != ObjectId.Null)
                                {
                                    lista_od_stg_justif.Add(justif1);
                                    lista_od_stg_object_id.Add(id1);
                                }
                            }
                        }

                        attach_od_to_stg(lista_od_stg_object_id, lista_od_stg_justif);

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

        }

        private void button_stg_out_Click(object sender, EventArgs e)
        {
            if (dt_stg == null || dt_stg.Rows.Count == 0)
            {
                MessageBox.Show("no stg data present\r\noperation abborted");
                return;
            }

            set_enable_false();
            try
            {
                if (dt_stg != null && dt_stg.Rows.Count > 0)
                {
                    Worksheet W1 = Functions.get_worksheet_W1(true, stg_data_tab);
                    format_and_transfer_dt_stg_to_excel(W1);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            set_enable_true();
        }

        private void button_stg_in_Click(object sender, EventArgs e)
        {
            try
            {

                Microsoft.Office.Interop.Excel.Application Excel1 = null;
                bool is_found = false;
                try
                {
                    Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");

                    if (Excel1 != null)
                    {

                        foreach (Workbook workbook1 in Excel1.Workbooks)
                        {
                            if (is_found == false)
                            {
                                foreach (Worksheet W1 in workbook1.Worksheets)
                                {
                                    if (is_found == false && W1.Name == stg_data_tab)
                                    {
                                        is_found = true;
                                        dt_stg = build_dt_stg_from_config_excel(W1);
                                        attach_dt_stg_to_datagridview(dt_stg);
                                    }
                                }
                            }

                        }
                    }



                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("no excel found");

                }


            }
            catch (System.Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }


        private void button_highlight_staging_Click(object sender, EventArgs e)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            if (dataGridView_stg_data.Rows.Count > 0)
            {
                set_enable_false();

                try
                {
                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                    using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);

                            bool ask_for_selection = false;
                            Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_object = (Autodesk.AutoCAD.EditorInput.PromptSelectionResult)Editor1.SelectImplied();

                            if (Rezultat_object.Status == PromptStatus.OK)
                            {
                                if (Rezultat_object.Value.Count == 0)
                                {
                                    ask_for_selection = true;
                                }
                                if (Rezultat_object.Value.Count > 1)
                                {
                                    MessageBox.Show("There is more than one object selected," + "\r\n" + "the first object in selection will be the one that will be current in table");
                                    ask_for_selection = false;
                                }
                            }
                            else ask_for_selection = true;



                            if (ask_for_selection == true)
                            {
                                this.WindowState = FormWindowState.Minimized;
                                Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_object = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                                Prompt_object.MessageForAdding = "\nSelect an ATWS";
                                Prompt_object.SingleOnly = true;
                                Rezultat_object = Editor1.GetSelection(Prompt_object);

                            }


                            if (Rezultat_object.Status != PromptStatus.OK)
                            {
                                this.WindowState = FormWindowState.Normal;

                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                set_enable_true();
                                return;
                            }
                            this.WindowState = FormWindowState.Normal;


                            Entity Ent1 = (Entity)Trans1.GetObject(Rezultat_object.Value[0].ObjectId, OpenMode.ForRead);
                            string handle1 = Ent1.ObjectId.Handle.Value.ToString();

                            for (int i = 0; i < dataGridView_stg_data.Rows.Count; ++i)
                            {
                                string handle2 = Convert.ToString(dataGridView_stg_data.Rows[i].Cells[atws_handle_column].Value);
                                if (handle1 == handle2)
                                {
                                    dataGridView_stg_data.CurrentCell = dataGridView_stg_data.Rows[i].Cells[0];
                                    i = dataGridView_stg_data.Rows.Count;
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

            set_enable_true();
        }

        private void comboBox_stg_od_name_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                Functions.load_object_data_fieds_to_combobox(comboBox_stg_od_name, comboBox_stg_od_field);
            }
            catch (System.Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void comboBox_stg_od_name_DropDown(object sender, EventArgs e)
        {
            try
            {
                Functions.load_object_data_table_name_to_combobox(comboBox_stg_od_name);
                if (comboBox_stg_od_name.Items.Count == 1)
                {
                    comboBox_stg_od_field.Items.Clear();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private System.Data.DataTable get_dt_stg_structure()
        {
            System.Data.DataTable dt1 = new System.Data.DataTable();
            dt1.Columns.Add(stg_sta1_column, typeof(double));
            dt1.Columns.Add(stg_sta2_column, typeof(double));
            dt1.Columns.Add(stg_area_column, typeof(double));
            dt1.Columns.Add(stg_justification_column, typeof(string));
            dt1.Columns.Add(stg_handle_column, typeof(string));
            return dt1;
        }

        private void update_dt_stg_handles_from_dwg(object sender, EventArgs e)
        {
            if (dt_stg != null && dt_stg.Rows.Count > 0)
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.Database.TransactionManager.StartTransaction())
                    {
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;
                        for (int i = dt_stg.Rows.Count - 1; i >= 0; --i)
                        {
                            if (dt_stg.Rows[i][stg_handle_column] != DBNull.Value)
                            {
                                string handle1 = Convert.ToString(dt_stg.Rows[i][stg_handle_column]);
                                ObjectId id1 = Functions.GetObjectId(ThisDrawing.Database, handle1);
                                if (id1 == ObjectId.Null)
                                {
                                    dt_stg.Rows[i].Delete();
                                    if (dt_manual_stg != null && dt_manual_stg.Columns.Count > 0)
                                    {
                                        for (int k = dt_manual_stg.Columns.Count - 1; k >= 0; --k)
                                        {
                                            if (dt_manual_stg.Columns[k].ColumnName.ToLower().Replace("x ", "").Replace("y ", "") == handle1.ToLower())
                                            {
                                                dt_manual_stg.Columns.RemoveAt(k);
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    Polyline ar1 = null;
                                    try
                                    {
                                        ar1 = Trans1.GetObject(id1, OpenMode.ForRead) as Polyline;
                                    }
                                    catch (Autodesk.AutoCAD.Runtime.Exception)
                                    {


                                    }

                                    if (ar1 == null)
                                    {
                                        dt_stg.Rows[i].Delete();
                                        if (dt_manual_stg != null && dt_manual_stg.Columns.Count > 0)
                                        {
                                            for (int k = dt_manual_stg.Columns.Count - 1; k >= 0; --k)
                                            {
                                                if (dt_manual_stg.Columns[k].ColumnName.ToLower().Replace("x ", "").Replace("y ", "") == handle1.ToLower())
                                                {
                                                    dt_manual_stg.Columns.RemoveAt(k);
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (ar1.Layer != comboBox_layer_stg.Text)
                                        {
                                            dt_stg.Rows[i].Delete();
                                            if (dt_manual_stg != null && dt_manual_stg.Columns.Count > 0)
                                            {
                                                for (int k = dt_manual_stg.Columns.Count - 1; k >= 0; --k)
                                                {
                                                    if (dt_manual_stg.Columns[k].ColumnName.ToLower().Replace("x ", "").Replace("y ", "") == handle1.ToLower())
                                                    {
                                                        dt_manual_stg.Columns.RemoveAt(k);
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
        }

        private void attach_od_to_stg(List<ObjectId> lista_od_stg_object_id, List<string> lista_od_stg_justif)
        {
            if (checkBox_use_od.Checked == false) return;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                if (lista_od_stg_object_id.Count > 0)
                {
                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                    BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;



                    List<string> lista_field_name = Functions.get_object_data_table_field_names(Tables1, comboBox_stg_od_name.Text);

                    if (lista_field_name == null || lista_field_name.Count == 0 || lista_field_name.Contains(comboBox_stg_od_field.Text) == false)
                    {
                        MessageBox.Show("Issue with stg data table found in the drawing");
                        set_enable_true();
                        return;
                    }

                    List<Autodesk.Gis.Map.Constants.DataType> lista_types = Functions.get_object_data_table_data_types(Tables1, comboBox_stg_od_name.Text);

                    for (int i = 0; i < lista_od_stg_object_id.Count; ++i)
                    {

                        Polyline stg1 = Trans1.GetObject(lista_od_stg_object_id[i], OpenMode.ForWrite) as Polyline;

                        List<object> lista_val = new List<object>();

                        for (int k = 0; k < lista_field_name.Count; ++k)
                        {
                            if (lista_field_name[k] == comboBox_stg_od_field.Text)
                            {
                                if (lista_types[k] == Autodesk.Gis.Map.Constants.DataType.Character)
                                {
                                    lista_val.Add(lista_od_stg_justif[i]);
                                }
                                else
                                {
                                    MessageBox.Show("The Object Data field " + comboBox_stg_od_field.Text + " is not defined as character field.\r\nPlease make sure you selected the correct field.\r\nOperation aborted");
                                    Entity ent1 = Trans1.GetObject(lista_od_stg_object_id[i], OpenMode.ForWrite) as Entity;
                                    ent1.Erase();
                                    Trans1.Commit();
                                    set_enable_true();
                                    return;
                                }
                            }
                            else
                            {
                                lista_val.Add(null);
                            }
                        }
                        Functions.Populate_object_data_table_from_objectid(lista_od_stg_object_id[i], comboBox_stg_od_name.Text, lista_val, lista_types);
                    }
                }
                Trans1.Commit();
            }

        }

        private void build_dt_manual_stg_from_config_excel(Worksheet W1)
        {
            if (dt_stg == null || dt_stg.Rows.Count == 0)
            {
                dt_manual_stg = null;
                return;
            }

            List<string> lista1 = new List<string>();

            for (int i = 0; i < dt_stg.Rows.Count; ++i)
            {
                if (dt_stg.Rows[i][stg_handle_column] != DBNull.Value)
                {
                    lista1.Add(Convert.ToString(dt_stg.Rows[i][stg_handle_column]));
                }
            }
            dt_manual_stg = new System.Data.DataTable();
            for (int j = 0; j < lista1.Count; ++j)
            {
                dt_manual_stg.Columns.Add(lista1[j], typeof(Point2d));
            }

            string last_col = Functions.get_excel_column_letter(10000);

            Range range1 = W1.Range["A1:" + last_col + "1"];
            object[,] values1 = new object[1, 10000];
            values1 = range1.Value2;


            for (int j = 1; j <= values1.Length; j += 2)
            {
                object col_name_val1 = values1[1, j];
                object col_name_val2 = values1[1, j + 1];
                if (col_name_val1 != null && col_name_val2 != null)
                {
                    string dtcol1 = Convert.ToString(col_name_val1).Replace("X ", "");
                    string dtcol2 = Convert.ToString(col_name_val2).Replace("Y ", "");

                    if (dtcol1 == dtcol2 && dt_manual_stg.Columns.Contains(Convert.ToString(dtcol1)) == true)
                    {
                        string col1 = Functions.get_excel_column_letter(j);
                        string col2 = Functions.get_excel_column_letter(j + 1);

                        Range range2 = W1.Range[col1 + "2:" + col2 + "30001"];
                        object[,] values2 = new object[30000, 2];
                        values2 = range2.Value2;

                        for (int i = 1; i <= values2.Length / 2; ++i)
                        {
                            object obj_x = values2[i, 1];
                            object obj_y = values2[i, 2];
                            if (obj_x != null && obj_y != null)
                            {
                                if (Functions.IsNumeric(Convert.ToString(obj_x)) == true && Functions.IsNumeric(Convert.ToString(obj_y)) == true)
                                {
                                    double x1 = Convert.ToDouble(obj_x);
                                    double y1 = Convert.ToDouble(obj_y);

                                    if (dt_manual_stg.Rows.Count < i)
                                    {
                                        dt_manual_stg.Rows.Add();
                                    }


                                    dt_manual_stg.Rows[i - 1][dtcol1] = new Point2d(x1, y1);

                                }
                            }
                            else
                            {
                                i = values2.Length;
                            }

                        }
                    }
                    else
                    {
                        MessageBox.Show("there are issues with column names into staging area geometry tab");
                        j = values1.Length + 1;
                    }




                }
                else
                {
                    j = values1.Length + 1;
                }
            }


            //Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt_manual_stg);

        }

    }
}
