using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Microsoft.Office.Interop.Excel;


namespace Alignment_mdi
{
    public partial class Toolz_form : Form
    {
        System.Data.DataTable dt_layout = null;

        public Toolz_form()
        {
            InitializeComponent();

            button_cl_l.Visible = false;
            button_cl_nl.Visible = true;

        }

        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();

            lista_butoane.Add(button_load_dwgs_from_excel);
            lista_butoane.Add(button_rename_layout);
            lista_butoane.Add(dataGridView_layout);
            lista_butoane.Add(button_select_drawings);




            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = false;
            }
        }
        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();

            lista_butoane.Add(button_load_dwgs_from_excel);
            lista_butoane.Add(button_rename_layout);
            lista_butoane.Add(dataGridView_layout);
            lista_butoane.Add(button_select_drawings);


            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }

        private void comboBox_xl_DropDown(object sender, EventArgs e)
        {
            Functions.Load_opened_worksheets_to_combobox(comboBox_xl);
        }

        private void button_load_dwgs_from_excel_Click(object sender, EventArgs e)
        {
            try
            {

                set_enable_false();
                dt_layout = get_dt_layout_structure();

                if (comboBox_xl.Text != "")
                {
                    string string1 = comboBox_xl.Text;
                    if (string1.Contains("[") == true && string1.Contains("]") == true)
                    {
                        string filename = string1.Substring(string1.IndexOf("]") + 4, string1.Length - (string1.IndexOf("]") + 4));

                        string sheet_name = string1.Substring(1, string1.IndexOf("]") - 1);
                        if (filename.Length > 0 && sheet_name.Length > 0)
                        {
                            set_enable_false();
                            Microsoft.Office.Interop.Excel.Worksheet W1 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, sheet_name);
                            if (W1 != null)
                            {

                                Range range1 = W1.Range["A1:C30000"];
                                object[,] values1 = new object[30000, 3];
                                values1 = range1.Value2;

                                for (int i = 2; i <= 30000; ++i)
                                {
                                    object val1 = values1[i, 1];
                                    object val2 = values1[i, 2];
                                    object val3 = values1[i, 3];
                                    if (val1 != null && val2 != null && val3 != null && Functions.IsNumeric(val2.ToString()) == true)
                                    {
                                        dt_layout.Rows.Add();
                                        dt_layout.Rows[dt_layout.Rows.Count - 1][0] = Convert.ToString(val1);
                                        dt_layout.Rows[dt_layout.Rows.Count - 1][1] = Convert.ToInt32(val2);
                                        dt_layout.Rows[dt_layout.Rows.Count - 1][2] = Convert.ToString(val3);
                                    }
                                    else
                                    {
                                        i = values1.Length + 1;
                                    }
                                }

                                dataGridView_layout.DataSource = dt_layout;
                                dataGridView_layout.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                                dataGridView_layout.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                                dataGridView_layout.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                                dataGridView_layout.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                                dataGridView_layout.DefaultCellStyle.ForeColor = Color.White;
                                dataGridView_layout.EnableHeadersVisualStyles = false;
                                button_cl_l.Visible = true;
                                button_cl_nl.Visible = false;
                            }
                        }
                    }
                }




            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                button_cl_l.Visible = false;
                button_cl_nl.Visible = true;
            }
            set_enable_true();
        }

        private System.Data.DataTable get_dt_layout_structure()
        {
            System.Data.DataTable dt1 = new System.Data.DataTable();
            dt1.Columns.Add("Dwg", typeof(string));
            dt1.Columns.Add("Layout Index", typeof(int));
            dt1.Columns.Add("New Name", typeof(string));

            return dt1;
        }

        private void button_rename_layout_Click(object sender, EventArgs e)
        {
            try
            {
                set_enable_false();
                if (dt_layout != null && dt_layout.Rows.Count > 0)
                {

                    DocumentCollection DocumentManager1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
                    if (DocumentManager1.Count == 0)
                    {
                        string strTemplatePath = "acad.dwt";
                        Document acDoc = DocumentManager1.Add(strTemplatePath);
                        DocumentManager1.MdiActiveDocument = acDoc;
                    }

                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {

                            for (int i = 0; i < dt_layout.Rows.Count; i++)
                            {
                                if (dt_layout.Rows[i][0] != DBNull.Value && dt_layout.Rows[i][1] != DBNull.Value && dt_layout.Rows[i][2] != DBNull.Value)
                                {
                                    if (Functions.IsNumeric(Convert.ToString(dt_layout.Rows[i][1])) == true)
                                    {
                                        string file1 = Convert.ToString(dt_layout.Rows[i][0]);
                                        int index1 = Convert.ToInt32(dt_layout.Rows[i][1]);
                                        string new_name = Convert.ToString(dt_layout.Rows[i][2]);
                                        if (System.IO.File.Exists(file1) == true)
                                        {
                                            if (System.IO.File.Exists(file1) == true)
                                            {

                                                bool is_opened = false;
                                                DocumentCollection document_collection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;

                                                foreach (Document opened_dwg in document_collection)
                                                {

                                                    string file2 = opened_dwg.Database.OriginalFileName;
                                                    if (file1 == file2)
                                                    {
                                                        HostApplicationServices.WorkingDatabase = opened_dwg.Database;
                                                        document_collection.MdiActiveDocument = opened_dwg;
                                                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans3 = opened_dwg.TransactionManager.StartTransaction())
                                                        {
                                                            List<string> lista_layout = Functions.get_layout_names(Trans3, opened_dwg.Database);
                                                            if (lista_layout.Contains(new_name) == false)
                                                            {
                                                                BlockTableRecord BtrecordPS = Functions.get_layout_as_paperspace(Trans3, opened_dwg.Database, index1);
                                                                BtrecordPS.UpgradeOpen();
                                                                Layout Layout1 = Functions.get_layout(Trans3, opened_dwg.Database, index1);
                                                                Layout1.UpgradeOpen();
                                                                Layout1.LayoutName = new_name;
                                                                Trans3.Commit();
                                                                opened_dwg.Editor.Regen();
                                                                dataGridView_layout.Rows[i].Cells[0].Style.Font = new System.Drawing.Font(dataGridView_layout.Font, FontStyle.Bold);
                                                                dataGridView_layout.Rows[i].Cells[0].Style.ForeColor = Color.FromArgb(0, 0, 0);
                                                                dataGridView_layout.Rows[i].Cells[0].Style.BackColor = Color.FromArgb(255, 219, 88);

                                                                //
                                                                //LayoutManager LayoutManager1 = (LayoutManager)Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current;
                                                                //LayoutManager1.RenameLayout("80294-03-ML-02-001_B", new_name);
                                                            }


                                                            is_opened = true;
                                                        }
                                                        HostApplicationServices.WorkingDatabase = ThisDrawing.Database;
                                                    }
                                                }

                                                if (is_opened == false)
                                                {
                                                    using (Database Database2 = new Database(false, true))
                                                    {
                                                        Database2.ReadDwgFile(file1, FileOpenMode.OpenForReadAndAllShare, true, "");
                                                        //System.IO.FileShare.ReadWrite, false, null);
                                                         Database2.CloseInput(true);
                                                        HostApplicationServices.WorkingDatabase = Database2;
                                                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans2 = Database2.TransactionManager.StartTransaction())
                                                        {


                                                            List<string> lista_layout = Functions.get_layout_names(Trans2, Database2);
                                                            if (lista_layout.Contains(new_name) == false)
                                                            {
                                                                BlockTableRecord BtrecordPS = Functions.get_layout_as_paperspace(Trans2, Database2, index1);
                                                                BtrecordPS.UpgradeOpen();
                                                                Layout Layout1 = Functions.get_layout(Trans2, Database2, index1);
                                                                Layout1.UpgradeOpen();
                                                                Layout1.LayoutName = new_name;

                                                                Trans2.Commit();

                                                                dataGridView_layout.Rows[i].Cells[0].Style.Font = new System.Drawing.Font(dataGridView_layout.Font, FontStyle.Bold);
                                                                dataGridView_layout.Rows[i].Cells[0].Style.ForeColor = Color.FromArgb(0, 0, 0);
                                                                dataGridView_layout.Rows[i].Cells[0].Style.BackColor = Color.FromArgb(255, 219, 88);
                                                            }
                                                        }

                                                        HostApplicationServices.WorkingDatabase = ThisDrawing.Database;
                                                        Database2.SaveAs(file1, true, DwgVersion.Current, Database2.SecurityParameters);
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
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            set_enable_true();
        }

        private void button_select_drawings_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog fbd = new OpenFileDialog())
            {
                fbd.Multiselect = true;
                fbd.Filter = "Autocad files (*.dwg)|*.dwg";

                List<string> drawing_list = new List<string>();

                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    if (dt_layout == null)
                    {
                        dt_layout = get_dt_layout_structure();
                    }
                    else
                    {
                        for (int i = 0; i < dt_layout.Rows.Count; i++)
                        {
                            if (dt_layout.Rows[i][0] != DBNull.Value)
                            {
                                drawing_list.Add(Convert.ToString(dt_layout.Rows[i][0]));
                            }
                        }
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
                                foreach (string file1 in fbd.FileNames)
                                {
                                    if (drawing_list.Contains(file1) == false)
                                    {
                                        drawing_list.Add(file1);


                                        using (Database Database2 = new Database(false, true))
                                        {
                                            Database2.ReadDwgFile(file1, FileOpenMode.OpenForReadAndAllShare, true, "");
                                            //System.IO.FileShare.ReadWrite, false, null);
                                            Database2.CloseInput(true);
                                            HostApplicationServices.WorkingDatabase = Database2;
                                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans2 = Database2.TransactionManager.StartTransaction())
                                            {
                                                List<string> lista_layout = Functions.get_layout_names(Trans2, Database2);
                                                if(lista_layout.Contains("Model")==true)
                                                {
                                                    lista_layout.Remove("Model");
                                                }

                                                if(lista_layout.Count>0)
                                                {
                                                    for (int i = 0; i < lista_layout.Count; i++)
                                                    {
                                                        dt_layout.Rows.Add();
                                                        dt_layout.Rows[dt_layout.Rows.Count - 1][0] = file1;
                                                        dt_layout.Rows[dt_layout.Rows.Count - 1][1] = i+1;
                                                        dt_layout.Rows[dt_layout.Rows.Count - 1][2] = lista_layout[i];
                                                    }
                                                }

                                            }

                                            HostApplicationServices.WorkingDatabase = ThisDrawing.Database;

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

                    Editor1.SetImpliedSelection(Empty_array);
                    Editor1.WriteMessage("\nCommand:");
                    set_enable_true();


                    Functions.Transfer_datatable_to_new_excel_spreadsheet(dt_layout);



                }
            }
        }
    }



}
