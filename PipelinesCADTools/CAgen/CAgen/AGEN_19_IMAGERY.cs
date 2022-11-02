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
using System.IO;

namespace Alignment_mdi
{
    public partial class image_form : Form
    {
        System.Data.DataTable dt_layout = null;
        List<Polyline> lista_poly = null;
        List<string> lista_dwg = null;
        List<string> lista_layout = null;

        private ContextMenuStrip ContextMenuStrip_layout;
        System.Data.DataTable dt_image = null;

        public image_form()
        {
            InitializeComponent();

            var toolStripMenuItem8 = new ToolStripMenuItem { Text = "Remove" };
            toolStripMenuItem8.Click += Unselect_cell_Click;

            var toolStripMenuItem9 = new ToolStripMenuItem { Text = "Remove All" };
            toolStripMenuItem9.Click += Unselect_all_cells_Click;

            ContextMenuStrip_layout = new ContextMenuStrip();
            ContextMenuStrip_layout.Items.AddRange(new ToolStripItem[] { toolStripMenuItem8, toolStripMenuItem9 });
        }


        private void Unselect_cell_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView_layout.RowCount > 0)
                {
                    int idx1 = dataGridView_layout.CurrentCell.RowIndex;

                    if (idx1 >= 0)
                    {
                        string dwg1 = Convert.ToString(dataGridView_layout.Rows[idx1].Cells[0].Value);

                        for (int i = dt_layout.Rows.Count - 1; i >= 0; --i)
                        {
                            if (dt_layout.Rows[i][0] != DBNull.Value)
                            {
                                string dwg2 = Convert.ToString(dt_layout.Rows[i][0]);
                                if (dwg1 == dwg2)
                                {
                                    dt_layout.Rows[i].Delete();
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

        private void Unselect_all_cells_Click(object sender, EventArgs e)
        {
            set_enable_false();
            try
            {
                if (dataGridView_layout.RowCount > 0)
                {
                    dt_layout.Rows.Clear();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            set_enable_true();
        }

        private void dataGridView_layout_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right && e.RowIndex >= 0)
            {
                dataGridView_layout.CurrentCell = dataGridView_layout.Rows[e.RowIndex].Cells[e.ColumnIndex];
                ContextMenuStrip_layout.Show(Cursor.Position);
                ContextMenuStrip_layout.Visible = true;
            }
            else
            {
                ContextMenuStrip_layout.Visible = false;
            }


        }

        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();

            lista_butoane.Add(button_attach_images);
            lista_butoane.Add(button_load_dwg_and_images);
            lista_butoane.Add(button_remove_all_images);
            lista_butoane.Add(button_select_drawings);
            lista_butoane.Add(button_set_image);
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


            lista_butoane.Add(button_attach_images);
            lista_butoane.Add(button_load_dwg_and_images);
            lista_butoane.Add(button_remove_all_images);
            lista_butoane.Add(button_select_drawings);
            lista_butoane.Add(button_set_image);
            lista_butoane.Add(dataGridView_layout);
            lista_butoane.Add(button_select_drawings);


            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }



        private System.Data.DataTable get_dt_dwg_structure()
        {
            System.Data.DataTable dt1 = new System.Data.DataTable();
            dt1.Columns.Add("Dwg", typeof(string));
            return dt1;
        }

        private void button_set_imageframe_to_zero_Click(object sender, EventArgs e)
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

                        lista_dwg = new List<string>();
                        lista_layout = new List<string>();
                        lista_poly = new List<Polyline>();
                        List<ObjectId> lista_objid = new List<ObjectId>();

                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {

                            for (int i = 0; i < dt_layout.Rows.Count; i++)
                            {

                                if (dt_layout.Rows[i][0] != DBNull.Value)
                                {
                                    string file1 = Convert.ToString(dt_layout.Rows[i][0]);

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
                                                    DBDictionary theNOD = Trans2.GetObject(Database2.NamedObjectsDictionaryId, OpenMode.ForRead) as DBDictionary;
                                                    RasterVariables rasterVars = null;
                                                    string kImageVars = "ACAD_IMAGE_VARS";
                                                    if (theNOD.Contains(kImageVars) == true)
                                                    {
                                                        ObjectId rastVarsId = theNOD.GetAt(kImageVars);
                                                        rasterVars = Trans2.GetObject(rastVarsId, OpenMode.ForWrite) as RasterVariables;
                                                    }
                                                    else
                                                    {
                                                        rasterVars = new RasterVariables();
                                                        theNOD.UpgradeOpen();
                                                        theNOD.SetAt(kImageVars, rasterVars);
                                                        Trans2.AddNewlyCreatedDBObject(rasterVars, true);
                                                    }
                                                    rasterVars.ImageFrame = FrameSetting.ImageFrameOff;
                                                    Trans2.Commit();
                                                    Database2.SaveAs(file1, true, DwgVersion.Current, ThisDrawing.Database.SecurityParameters);
                                                }
                                                HostApplicationServices.WorkingDatabase = ThisDrawing.Database;
                                                Database2.Dispose();
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


                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    if (dt_layout == null)
                    {
                        dt_layout = get_dt_dwg_structure();
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
                                    dt_layout.Rows.Add();
                                    dt_layout.Rows[dt_layout.Rows.Count - 1][0] = file1;


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


                    dataGridView_layout.DataSource = dt_layout;
                    dataGridView_layout.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                    dataGridView_layout.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                    dataGridView_layout.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                    dataGridView_layout.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                    dataGridView_layout.DefaultCellStyle.ForeColor = Color.White;
                    dataGridView_layout.EnableHeadersVisualStyles = false;

                }
            }
        }


        private void Create_rectangle_object_data_table()
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

                            List1.Add("Drawing");
                            List2.Add("Drawing");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            List1.Add("LayoutName");
                            List2.Add("layout name");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);


                            List1.Add("UserName");
                            List2.Add("Generated by");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            List1.Add("Date");
                            List2.Add("Date and Time");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);


                            Functions.Get_object_data_table("ODYXX", "Generated by Profiler", List1, List2, List3);

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

        private void Append_object_data_to_ODYXX(List<ObjectId> lista1, List<string> drawing_name, List<string> layout_name)
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
                            Lista_val.Add(drawing_name[i]);
                            Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                            Lista_val.Add(layout_name[i]);
                            Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                            Lista_val.Add(Environment.UserName.ToUpper());
                            Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            Lista_val.Add(System.DateTime.Today.Year + "-" + System.DateTime.Today.Month + "-" + System.DateTime.Today.Day + " at " + System.DateTime.Now.Hour + ":" + System.DateTime.Now.Minute);
                            Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);


                            Functions.Populate_object_data_table_from_objectid(Tables1, id1, "ODYXX", Lista_val, Lista_type);
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

        private void button_load_dwg_and_images_Click(object sender, EventArgs e)
        {
            try
            {
                string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                {
                    ProjFolder = ProjFolder + "\\";
                }

                if (System.IO.Directory.Exists(ProjFolder) == true)
                {
                    string fisier_image = ProjFolder + _AGEN_mainform.imagery_excel_name;

                    if (System.IO.File.Exists(fisier_image) == true)
                    {

                        Microsoft.Office.Interop.Excel.Worksheet W1 = null;
                        Microsoft.Office.Interop.Excel.Application Excel1 = null;

                        try
                        {
                            Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                        }
                        catch (System.Exception ex)
                        {
                            Excel1 = new Microsoft.Office.Interop.Excel.Application();

                        }

                        set_enable_false();

                        if (Excel1.Workbooks.Count == 0) Excel1.Visible = _AGEN_mainform.ExcelVisible;
                        Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(fisier_image);

                        try
                        {
                            Workbook1 = Excel1.Workbooks.Open(fisier_image);
                            W1 = Workbook1.Worksheets[1];

                            Build_Data_table_imagery_from_excel(W1, 9);

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
                            if (W1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W1);
                            if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                            if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                        }


                        if (dt_image != null && dt_image.Rows.Count > 0)
                        {

                            for (int i = 0; i < dt_image.Rows.Count; i++)
                            {
                                string file1 = Convert.ToString(dt_image.Rows[i]["World_File"]);
                                using (System.IO.StreamReader reader1 = new StreamReader(file1))
                                {
                                    int nr = 1;
                                    while (reader1.Peek() > 0)
                                    {
                                        string line1 = reader1.ReadLine();
                                        if (Functions.IsNumeric(line1.Replace(" ", "")) == true)
                                        {

                                            if (nr == 1)
                                            {
                                                dt_image.Rows[i]["Width"] = Convert.ToDouble(line1.Replace(" ", ""));
                                            }
                                            if (nr == 4)
                                            {
                                                dt_image.Rows[i]["Height"] = Convert.ToDouble(line1.Replace(" ", "").Replace("-", ""));
                                            }
                                            if (nr == 5)
                                            {
                                                dt_image.Rows[i]["X"] = Convert.ToDouble(line1.Replace(" ", ""));
                                            }
                                            if (nr == 6)
                                            {
                                                dt_image.Rows[i]["Y"] = Convert.ToDouble(line1.Replace(" ", ""));
                                            }
                                        }
                                        ++nr;

                                    }
                                }
                            }

                            dataGridView_layout.DataSource = dt_image;
                            dataGridView_layout.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                            dataGridView_layout.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                            dataGridView_layout.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                            dataGridView_layout.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                            dataGridView_layout.DefaultCellStyle.ForeColor = Color.White;
                            dataGridView_layout.EnableHeadersVisualStyles = false;
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


        private void button_attach_images_Click(object sender, EventArgs e)
        {
            if (dt_image == null || dt_image.Rows.Count == 0) return;

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

                        List<string> lista_dwg_completed = new List<string>();

                        int counter = 1;

                        for (int i = dt_image.Rows.Count - 1; i >= 0; --i)
                        {

                            string file1 = Convert.ToString(dt_image.Rows[i]["Drawing_File"]);
                            if (System.IO.File.Exists(file1) == true)
                            {
                                if (lista_dwg_completed.Contains(file1) == false)
                                {
                                    lista_dwg_completed.Add(file1);
                                    List<double> lista_h1 = new List<double>();
                                    List<double> lista_w1 = new List<double>();
                                    List<double> lista_x1 = new List<double>();
                                    List<double> lista_y1 = new List<double>();
                                    List<string> lista_image_filename = new List<string>();
                                    List<string> lista_Xref_Name = new List<string>();

                                    lista_image_filename.Add(Convert.ToString(dt_image.Rows[i]["Image_File"]));
                                    lista_Xref_Name.Add("AGEN_image_" + Convert.ToString(counter));
                                    ++counter;

                                    lista_h1.Add(Convert.ToDouble(dt_image.Rows[i]["Height"]));
                                    lista_w1.Add(Convert.ToDouble(dt_image.Rows[i]["Width"]));
                                    lista_x1.Add(Convert.ToDouble(dt_image.Rows[i]["X"]));
                                    lista_y1.Add(Convert.ToDouble(dt_image.Rows[i]["Y"]));

                                    for (int j = i - 1; j >= 0; --j)
                                    {
                                        string file2 = Convert.ToString(dt_image.Rows[j]["Drawing_File"]);
                                        if (file1 == file2)
                                        {
                                            lista_image_filename.Add(Convert.ToString(dt_image.Rows[j]["Image_File"]));
                                            lista_Xref_Name.Add("AGEN_image_" + Convert.ToString(counter));
                                            ++counter;
                                            lista_h1.Add(Convert.ToDouble(dt_image.Rows[j]["Height"]));
                                            lista_w1.Add(Convert.ToDouble(dt_image.Rows[j]["Width"]));
                                            lista_x1.Add(Convert.ToDouble(dt_image.Rows[j]["X"]));
                                            lista_y1.Add(Convert.ToDouble(dt_image.Rows[j]["Y"]));
                                        }

                                    }

                                    counter = 1;

                                    for (int k = lista_image_filename.Count - 1; k >= 0; --k)
                                    {
                                        string image_filename = lista_image_filename[k];
                                        if (System.IO.File.Exists(image_filename) == false)
                                        {
                                            lista_image_filename.RemoveAt(k);
                                            lista_Xref_Name.RemoveAt(k);
                                            lista_h1.RemoveAt(k);
                                            lista_w1.RemoveAt(k);
                                            lista_x1.RemoveAt(k);
                                            lista_y1.RemoveAt(k);
                                        }
                                    }









                                    using (Database Database2 = new Database(false, true))
                                    {
                                        Database2.ReadDwgFile(file1, FileOpenMode.OpenForReadAndWriteNoShare, true, "");
                                        //System.IO.FileShare.ReadWrite, false, null);
                                        Database2.CloseInput(true);
                                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans2 = Database2.TransactionManager.StartTransaction())
                                        {

                                            BlockTableRecord BtrecordMS2 = Functions.get_modelspace(Trans2, Database2);
                                            BtrecordMS2.UpgradeOpen();

                                            BlockTable BlockTable2 = Trans2.GetObject(Database2.BlockTableId, OpenMode.ForRead) as BlockTable;
                                            RasterImageDef Raster_def;
                                            bool bRasterDefCreated = false;
                                            ObjectId acImgDefId;

                                            // Get the image dictionary
                                            ObjectId acImgDctID = RasterImageDef.GetImageDictionary(Database2);

                                            // Check to see if the dictionary does not exist, it not then create it
                                            if (acImgDctID.IsNull)
                                            {
                                                acImgDctID = RasterImageDef.CreateImageDictionary(Database2);
                                            }
                                            // Open the image dictionary
                                            DBDictionary acImgDict = Trans2.GetObject(acImgDctID, OpenMode.ForRead) as DBDictionary;

                                            for (int k = 0; k < lista_image_filename.Count; k++)
                                            {
                                                string Xref_Name = lista_Xref_Name[k];
                                                string image_filename = lista_image_filename[k];
                                                double x1 = lista_x1[k];
                                                double y1 = lista_y1[k];
                                                double h1 = lista_h1[k];
                                                double w1 = lista_w1[k];




                                                // Check to see if the image definition already exists
                                                if (acImgDict.Contains(Xref_Name) == true)
                                                {
                                                    int duplicate = 1;
                                                    string orig = Xref_Name;
                                                    do
                                                    {
                                                        Xref_Name = orig + duplicate.ToString();
                                                        ++duplicate;
                                                    } while (acImgDict.Contains(Xref_Name) == true);

                                                    // acImgDefId = acImgDict.GetAt(Xref_Name);

                                                    // Raster_def = Trans2.GetObject(acImgDefId, OpenMode.ForWrite) as RasterImageDef;
                                                }

                                                // Create a raster image definition
                                                RasterImageDef acRasterDefNew = new RasterImageDef();

                                                // Set the source for the image file
                                                acRasterDefNew.SourceFileName = image_filename;

                                                // Load the image into memory
                                                acRasterDefNew.Load();

                                                // Add the image definition to the dictionary
                                                acImgDict.UpgradeOpen();
                                                acImgDefId = acImgDict.SetAt(Xref_Name, acRasterDefNew);

                                                Trans2.AddNewlyCreatedDBObject(acRasterDefNew, true);

                                                Raster_def = acRasterDefNew;

                                                bRasterDefCreated = true;



                                                // Create the new image and assign it the image definition
                                                using (RasterImage acRaster = new RasterImage())
                                                {
                                                    acRaster.ImageDefId = acImgDefId;

                                                    // Use ImageWidth and ImageHeight to get the size of the image in pixels (1024 x 768).
                                                    // Use ResolutionMMPerPixel to determine the number of millimeters in a pixel so you 
                                                    // can convert the size of the drawing into other units or millimeters based on the 
                                                    // drawing units used in the current drawing.

                                                    // Define the width and height of the image
                                                    Vector3d Vector_width;
                                                    Vector3d Vector_height;

                                                    // Check to see if the measurement is set to English (Imperial) or Metric units
                                                    if (Database2.Measurement == MeasurementValue.English)
                                                    {

                                                    }

                                                    Vector_width = new Vector3d(w1 * acRaster.ImageWidth, 0, 0);
                                                    Vector_height = new Vector3d(0, h1 * acRaster.ImageHeight, 0);

                                                    // Define the position for the image 
                                                    Point3d insPt = new Point3d(x1, y1 - Vector_height.Length, 0.0);

                                                    // Define and assign a coordinate system for the image's orientation
                                                    CoordinateSystem3d coordinateSystem = new CoordinateSystem3d(insPt, Vector_width, Vector_height);
                                                    acRaster.Orientation = coordinateSystem;

                                                    // Set the rotation angle for the image
                                                    acRaster.Rotation = 0;

                                                    // Add the new object to the block table record and the transaction
                                                    BtrecordMS2.AppendEntity(acRaster);
                                                    Trans2.AddNewlyCreatedDBObject(acRaster, true);

                                                    ObjectIdCollection oBJiD_COL = new ObjectIdCollection();
                                                    oBJiD_COL.Add(acRaster.ObjectId);
                                                    DrawOrderTable DrawOrderTable2 = Trans2.GetObject(BtrecordMS2.DrawOrderTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite) as DrawOrderTable;
                                                    DrawOrderTable2.MoveToBottom(oBJiD_COL);

                                                    // Connect the raster definition and image together so the definition
                                                    // does not appear as "unreferenced" in the External References palette.
                                                    RasterImage.EnableReactors(true);
                                                    acRaster.AssociateRasterDef(Raster_def);

                                                    if (bRasterDefCreated)
                                                    {
                                                        Raster_def.Dispose();
                                                    }



                                                }
                                            }

                                            Trans2.Commit();
                                        }
                                        HostApplicationServices.WorkingDatabase = ThisDrawing.Database;
                                        Database2.SaveAs(file1, true, DwgVersion.Current, ThisDrawing.Database.SecurityParameters);
                                    }

                                }
                            }

                        }
                        Trans1.Commit();
                        MessageBox.Show("done");
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

        public static System.Data.DataTable Creaza_dt_image_datatable_structure()
        {

            System.Data.DataTable dt1 = new System.Data.DataTable();
            dt1.Columns.Add("Drawing_File", typeof(string));
            dt1.Columns.Add("Image_File", typeof(string));
            dt1.Columns.Add("World_File", typeof(string));
            dt1.Columns.Add("Width", typeof(double));
            dt1.Columns.Add("Height", typeof(double));
            dt1.Columns.Add("X", typeof(double));
            dt1.Columns.Add("Y", typeof(double));
            return dt1;
        }

        private System.Data.DataTable Build_Data_table_imagery_from_excel(Microsoft.Office.Interop.Excel.Worksheet W1, int Start_row)
        {

            dt_image = Creaza_dt_image_datatable_structure();

            string Col1 = "B";
            string Col2 = "C";
            string Col3 = "D";

            Range range1 = W1.Range[Col1 + Start_row.ToString() + ":" + Col1 + "30000"];
            object[,] values1 = new object[30000, 1];
            values1 = range1.Value2;

            Range range2 = W1.Range[Col2 + Start_row.ToString() + ":" + Col2 + "30000"];
            object[,] values2 = new object[30000, 1];
            values2 = range2.Value2;

            Range range3 = W1.Range[Col3 + Start_row.ToString() + ":" + Col3 + "30000"];
            object[,] values3 = new object[30000, 1];
            values3 = range3.Value2;

            bool is_data = false;
            for (int i = 1; i <= values1.Length; ++i)
            {
                object Valoare1 = values1[i, 1];
                object Valoare2 = values2[i, 1];
                object Valoare3 = values3[i, 1];
                if (Valoare1 != null && Valoare2 != null && Valoare3 != null)
                {
                    dt_image.Rows.Add();
                    is_data = true;
                }
                else
                {
                    i = values1.Length + 1;
                }
            }


            if (is_data == false)
            {
                return dt_image;
            }

            int NrR = dt_image.Rows.Count;


            Microsoft.Office.Interop.Excel.Range range_val = W1.Range[W1.Cells[Start_row, 2], W1.Cells[NrR + Start_row - 1, 4]];

            object[,] values = new object[NrR - 1, 3];

            values = range_val.Value2;

            for (int i = 0; i < dt_image.Rows.Count; ++i)
            {
                for (int j = 0; j < 3; ++j)
                {
                    object Valoare = values[i + 1, j + 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt_image.Rows[i][j] = Valoare;
                }
            }




            return dt_image;

        }

        private void button_remove_all_images_Click(object sender, EventArgs e)
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

                        lista_dwg = new List<string>();
                        lista_layout = new List<string>();
                        lista_poly = new List<Polyline>();
                        List<ObjectId> lista_objid = new List<ObjectId>();

                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {

                            for (int i = 0; i < dt_layout.Rows.Count; i++)
                            {

                                if (dt_layout.Rows[i][0] != DBNull.Value)
                                {
                                    string file1 = Convert.ToString(dt_layout.Rows[i][0]);

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

                                                    BlockTableRecord BtrecordMS2 = Functions.get_modelspace(Trans2, Database2);
                                                    BtrecordMS2.UpgradeOpen();

                                                    BlockTable BlockTable2 = Trans2.GetObject(Database2.BlockTableId, OpenMode.ForRead) as BlockTable;
                                                    RasterImageDef Raster_def;
                                                    // Get the image dictionary
                                                    ObjectId acImgDctID = RasterImageDef.GetImageDictionary(Database2);
                                                    if (acImgDctID.IsNull == false)
                                                    {

                                                       // Open the image dictionary
                                                        DBDictionary acImgDict = Trans2.GetObject(acImgDctID, OpenMode.ForWrite) as DBDictionary;

                                                        foreach (DBDictionaryEntry ob1 in acImgDict)
                                                        {
                                                            ObjectId id1 = ob1.Value;
                                                            Raster_def = Trans2.GetObject(id1, OpenMode.ForWrite) as RasterImageDef;
                                                            Raster_def.Erase();
                                                            
                                                        }
                                                    }

                                                    Trans2.Commit();
                                                    Database2.SaveAs(file1, true, DwgVersion.Current, ThisDrawing.Database.SecurityParameters);
                                                }
                                                HostApplicationServices.WorkingDatabase = ThisDrawing.Database;
                                                Database2.Dispose();
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

        public void change_image_contrast(Database Database2, double contrast, double brightness)
        {
            HostApplicationServices.WorkingDatabase = Database2;
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans2 = Database2.TransactionManager.StartTransaction())
            {
                Functions.make_ms_active(Trans2, Database2);
                BlockTable BlockTable1 = Database2.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                BlockTableRecord BTrecord_MS = Trans2.GetObject(BlockTable1[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                foreach (ObjectId id1 in BTrecord_MS)
                {
                    RasterImage imag1 = Trans2.GetObject(id1, OpenMode.ForRead) as RasterImage;
                    if (imag1 != null)
                    {
                        imag1.UpgradeOpen();

                    }
                }
                Trans2.Commit();
            }
        }

        private void button_adjust_images_Click(object sender, EventArgs e)
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

                                if (dt_layout.Rows[i][0] != DBNull.Value)
                                {
                                    string file1 = Convert.ToString(dt_layout.Rows[i][0]);

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

                                                    BlockTableRecord BtrecordMS2 = Functions.get_modelspace(Trans2, Database2);
                                                    BtrecordMS2.UpgradeOpen();


                                                        foreach (ObjectId id1 in BtrecordMS2)
                                                        {
                                                            
                                                          RasterImage    Agen_image = Trans2.GetObject(id1, OpenMode.ForWrite) as RasterImage;
                                                            double contrast = 50;
                                                            double brightness = 50;
                                                            if (Functions.IsNumeric(textBox_image_brightness.Text) == true) brightness = Convert.ToDouble(textBox_image_brightness.Text);
                                                            if (Functions.IsNumeric(textBox_image_contrast.Text) == true) contrast = Convert.ToDouble(textBox_image_contrast.Text);
                                                            if(Agen_image!=null)
                                                        {
                                                            Agen_image.Contrast = Convert.ToByte(contrast);
                                                            Agen_image.Brightness = Convert.ToByte(brightness);
                                                        }
                                                        }
                                                  

                                                    Trans2.Commit();
                                                    Database2.SaveAs(file1, true, DwgVersion.Current, ThisDrawing.Database.SecurityParameters);
                                                }
                                                HostApplicationServices.WorkingDatabase = ThisDrawing.Database;
                                                Database2.Dispose();
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
    }
}
