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
    public partial class laypurge_form : Form
    {
        System.Data.DataTable dt_layout = null;
        List<Polyline> lista_poly = null;
        List<string> lista_dwg = null;
        List<string> lista_layout = null;

        private ContextMenuStrip ContextMenuStrip_layout;


        public laypurge_form()
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

            lista_butoane.Add(button_remove_empty_layers_list);
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


            lista_butoane.Add(button_remove_empty_layers_list);
            lista_butoane.Add(dataGridView_layout);
            lista_butoane.Add(button_select_drawings);


            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }



        private System.Data.DataTable get_dt_layout_structure()
        {
            System.Data.DataTable dt1 = new System.Data.DataTable();
            dt1.Columns.Add("Dwg", typeof(string));
            dt1.Columns.Add("Layout Index", typeof(int));
            dt1.Columns.Add("Layout Name", typeof(string));

            return dt1;
        }

        private void button_vp2poly_Click(object sender, EventArgs e)
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
                                if (dt_layout.Rows[i][0] != DBNull.Value && dt_layout.Rows[i][1] != DBNull.Value && dt_layout.Rows[i][2] != DBNull.Value)
                                {
                                    if (Functions.IsNumeric(Convert.ToString(dt_layout.Rows[i][1])) == true)
                                    {
                                        string file1 = Convert.ToString(dt_layout.Rows[i][0]);
                                        int index1 = Convert.ToInt32(dt_layout.Rows[i][1]);
                                        string nume1 = Convert.ToString(dt_layout.Rows[i][2]);
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
                                                            if (lista_layout.Contains(nume1) == false)
                                                            {
                                                                BlockTableRecord BtrecordPS = Functions.get_layout_as_paperspace(Trans3, opened_dwg.Database, index1);

                                                                foreach (ObjectId id1 in BtrecordPS)
                                                                {
                                                                    Viewport vp1 = Trans3.GetObject(id1, OpenMode.ForRead) as Viewport;
                                                                    if (vp1 != null)
                                                                    {
                                                                        lista_dwg.Add(file1);
                                                                        lista_layout.Add(nume1);

                                                                        double rot1 = -vp1.TwistAngle;
                                                                        double len1 = vp1.Width;
                                                                        double height1 = vp1.Height;
                                                                        double scale1 = vp1.CustomScale;

                                                                        Point3d pt1 = new Point3d(vp1.ViewCenter.X, vp1.ViewCenter.Y, 0);
                                                                        Matrix3d transMatrx = Matrix3d.WorldToPlane(vp1.ViewDirection);
                                                                        transMatrx = Matrix3d.Displacement(vp1.ViewTarget - Point3d.Origin) * transMatrx;
                                                                        transMatrx = Matrix3d.Rotation(-vp1.TwistAngle, vp1.ViewDirection, vp1.ViewTarget) * transMatrx;
                                                                        pt1 = pt1.TransformBy(transMatrx);

                                                                        Polyline poly1 = new Polyline();
                                                                        poly1.AddVertexAt(0, new Point2d(pt1.X - (len1 / 2) / scale1, pt1.Y + (height1 / 2) / scale1), 0, 0, 0);
                                                                        poly1.AddVertexAt(1, new Point2d(pt1.X + (len1 / 2) / scale1, pt1.Y + (height1 / 2) / scale1), 0, 0, 0);
                                                                        poly1.AddVertexAt(2, new Point2d(pt1.X + (len1 / 2) / scale1, pt1.Y - (height1 / 2) / scale1), 0, 0, 0);
                                                                        poly1.AddVertexAt(3, new Point2d(pt1.X - (len1 / 2) / scale1, pt1.Y - (height1 / 2) / scale1), 0, 0, 0);
                                                                        poly1.Closed = true;
                                                                        poly1.TransformBy(Matrix3d.Rotation(rot1, Vector3d.ZAxis, new Point3d(pt1.X, pt1.Y, 0)));
                                                                        lista_poly.Add(poly1);

                                                                    }
                                                                }

                                                                dataGridView_layout.Rows[i].Cells[0].Style.Font = new System.Drawing.Font(dataGridView_layout.Font, FontStyle.Bold);
                                                                dataGridView_layout.Rows[i].Cells[0].Style.ForeColor = Color.FromArgb(0, 0, 0);
                                                                dataGridView_layout.Rows[i].Cells[0].Style.BackColor = Color.FromArgb(255, 219, 88);
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
                                                            BlockTableRecord BtrecordPS = Functions.get_layout_as_paperspace(Trans2, Database2, index1);

                                                            foreach (ObjectId id1 in BtrecordPS)
                                                            {
                                                                Viewport vp1 = Trans2.GetObject(id1, OpenMode.ForRead) as Viewport;
                                                                if (vp1 != null)
                                                                {
                                                                    lista_dwg.Add(file1);
                                                                    lista_layout.Add(nume1);

                                                                    double rot1 = -vp1.TwistAngle;
                                                                    double len1 = vp1.Width;
                                                                    double height1 = vp1.Height;
                                                                    double scale1 = vp1.CustomScale;

                                                                    Point3d pt1 = new Point3d(vp1.ViewCenter.X, vp1.ViewCenter.Y, 0);
                                                                    Matrix3d transMatrx = Matrix3d.WorldToPlane(vp1.ViewDirection);
                                                                    transMatrx = Matrix3d.Displacement(vp1.ViewTarget - Point3d.Origin) * transMatrx;
                                                                    transMatrx = Matrix3d.Rotation(-vp1.TwistAngle, vp1.ViewDirection, vp1.ViewTarget) * transMatrx;
                                                                    pt1 = pt1.TransformBy(transMatrx);


                                                                    Polyline poly1 = new Polyline();
                                                                    poly1.AddVertexAt(0, new Point2d(pt1.X - (len1 / 2) / scale1, pt1.Y + (height1 / 2) / scale1), 0, 0, 0);
                                                                    poly1.AddVertexAt(1, new Point2d(pt1.X + (len1 / 2) / scale1, pt1.Y + (height1 / 2) / scale1), 0, 0, 0);
                                                                    poly1.AddVertexAt(2, new Point2d(pt1.X + (len1 / 2) / scale1, pt1.Y - (height1 / 2) / scale1), 0, 0, 0);
                                                                    poly1.AddVertexAt(3, new Point2d(pt1.X - (len1 / 2) / scale1, pt1.Y - (height1 / 2) / scale1), 0, 0, 0);
                                                                    poly1.Closed = true;
                                                                    poly1.TransformBy(Matrix3d.Rotation(rot1, Vector3d.ZAxis, new Point3d(pt1.X, pt1.Y, 0)));
                                                                    lista_poly.Add(poly1);

                                                                }
                                                            }

                                                            dataGridView_layout.Rows[i].Cells[0].Style.Font = new System.Drawing.Font(dataGridView_layout.Font, FontStyle.Bold);
                                                            dataGridView_layout.Rows[i].Cells[0].Style.ForeColor = Color.FromArgb(0, 0, 0);
                                                            dataGridView_layout.Rows[i].Cells[0].Style.BackColor = Color.FromArgb(255, 219, 88);
                                                        }
                                                        HostApplicationServices.WorkingDatabase = ThisDrawing.Database;
                                                        Database2.Dispose();
                                                    }
                                                }
                                            }
                                        }
                                    }

                                }
                            }


                            if (lista_poly.Count > 0)
                            {
                                Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                                Create_rectangle_object_data_table();
                                Functions.Creaza_layer("VP_YXX", 40, false);

                                for (int i = 0; i < lista_poly.Count; i++)
                                {
                                    Polyline poly1 = lista_poly[i];
                                    poly1.Layer = "VP_YXX";
                                    poly1.ColorIndex = 256;
                                    BTrecord.AppendEntity(poly1);
                                    Trans1.AddNewlyCreatedDBObject(poly1, true);
                                    lista_objid.Add(poly1.ObjectId);
                                }
                                Append_object_data_to_ODYXX(lista_objid, lista_dwg, lista_layout);
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
                        dt_layout = get_dt_layout_structure();
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
                                    using (Database Database2 = new Database(false, true))
                                    {
                                        Database2.ReadDwgFile(file1, FileOpenMode.OpenForReadAndAllShare, true, "");
                                        //System.IO.FileShare.ReadWrite, false, null);
                                        Database2.CloseInput(true);
                                        HostApplicationServices.WorkingDatabase = Database2;
                                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans2 = Database2.TransactionManager.StartTransaction())
                                        {
                                            List<string> lista_layout = Functions.get_layout_names(Trans2, Database2);
                                            if (lista_layout.Contains("Model") == true)
                                            {
                                                lista_layout.Remove("Model");
                                            }

                                            if (lista_layout.Count > 0)
                                            {
                                                for (int i = 0; i < lista_layout.Count; i++)
                                                {
                                                    dt_layout.Rows.Add();
                                                    dt_layout.Rows[dt_layout.Rows.Count - 1][0] = file1;
                                                    dt_layout.Rows[dt_layout.Rows.Count - 1][1] = i + 1;
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

        private void button_remove_empty_layers_current_dwg_Click(object sender, EventArgs e)
        {

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
                        BlockTableRecord BTrecord_PS = Trans1.GetObject(BlockTable1[BlockTableRecord.PaperSpace], OpenMode.ForRead) as BlockTableRecord;
                        BlockTableRecord BTrecord_MS = Trans1.GetObject(BlockTable1[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;
                        LayerTable LayerTable1 = Trans1.GetObject(ThisDrawing.Database.LayerTableId, OpenMode.ForWrite) as LayerTable;

                        List<string> lista1 = new List<string>();

                        foreach (ObjectId obid in BTrecord_PS)
                        {
                            Entity ent1 = Trans1.GetObject(obid, OpenMode.ForRead) as Entity;
                            BlockReference block1 = Trans1.GetObject(obid, OpenMode.ForRead) as BlockReference;
                            if (ent1 != null)
                            {
                                if (lista1.Contains(ent1.Layer) == false) lista1.Add(ent1.Layer);

                                if (block1!=null)
                                {
                                    string nume_block = Functions.get_block_name(block1);
                                    foreach(ObjectId id2 in BlockTable1)
                                    {
                                        BlockTableRecord btr = Trans1.GetObject(id2, OpenMode.ForRead) as BlockTableRecord;
                                        if(nume_block==btr.Name)
                                        {

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

        }
    }
}
