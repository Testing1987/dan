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
    public partial class AGEN_tools : Form
    {

        bool Freeze_operations = false;

        System.Data.DataTable dt_vp = _AGEN_mainform.dt_vp;

        _AGEN_mainform Ag = null;

        System.Data.DataTable Display_dataTable = null;
        System.Data.DataTable dt_atr = null;



        private ContextMenuStrip ContextMenuStrip_open_alignment;

        string Excel_tblk_atr = "";


        public AGEN_tools()
        {
            InitializeComponent();

            if (dt_vp == null)
            {
                dt_vp = new System.Data.DataTable();
                dt_vp.Columns.Add("dwg", typeof(string));
                dt_vp.Columns.Add("polyline", typeof(Polyline));
            }

            var toolStripMenuItem1 = new ToolStripMenuItem { Text = "Open selected drawing" };
            toolStripMenuItem1.Click += open_DWG_Click;


            var toolStripMenuItem3 = new ToolStripMenuItem { Text = "Remove drawing" };
            toolStripMenuItem3.Click += remove_selected_dwg_Click;

            var toolStripMenuItem4 = new ToolStripMenuItem { Text = "Clear drawing list" };
            toolStripMenuItem4.Click += remove_all_dwg_Click;

            ContextMenuStrip_open_alignment = new ContextMenuStrip();
            ContextMenuStrip_open_alignment.Items.AddRange(new ToolStripItem[] { toolStripMenuItem1, toolStripMenuItem3, toolStripMenuItem4 });

        }



        private void remove_selected_dwg_Click(object sender, EventArgs e)
        {
            if (dataGridView_drawings.RowCount > 0)
            {
                int Index1 = dataGridView_drawings.CurrentCell.RowIndex;
                if (Index1 == -1)
                {
                    return;
                }

                string val1 = Convert.ToString(dataGridView_drawings.Rows[Index1].Cells[0].Value);
                dataGridView_drawings.Rows.RemoveAt(Index1);

                dt_vp.Rows[Index1].Delete();

            }
        }

        private void remove_all_dwg_Click(object sender, EventArgs e)
        {
            Display_dataTable = null;

            dataGridView_drawings.DataSource = "";
            dt_vp = new System.Data.DataTable();
            dt_vp.Columns.Add("dwg", typeof(string));
            dt_vp.Columns.Add("polyline", typeof(Polyline));
        }

        private void open_DWG_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView_drawings.RowCount > 0)
                {

                    int Index1 = dataGridView_drawings.CurrentCell.RowIndex;
                    if (Display_dataTable != null)
                    {
                        if (Display_dataTable.Rows.Count - 1 >= Index1)
                        {
                            string path0 = Display_dataTable.Rows[Index1][0].ToString();



                            if (System.IO.File.Exists(path0) == true)
                            {

                                bool is_opened = false;
                                DocumentCollection DocumentManager1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
                                foreach (Document Doc in DocumentManager1)
                                {
                                    if (Doc.Name == path0)
                                    {
                                        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument = Doc;
                                        is_opened = true;

                                    }

                                }

                                if (is_opened == false)
                                {
                                    DocumentCollectionExtension.Open(DocumentManager1, path0, false);
                                }

                            }
                            else
                            {
                                MessageBox.Show("file not found");
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

        private void dataGridView_drawings_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right && e.RowIndex >= 0)
            {
                dataGridView_drawings.CurrentCell = dataGridView_drawings.Rows[e.RowIndex].Cells[e.ColumnIndex];
                ContextMenuStrip_open_alignment.Show(Cursor.Position);
                ContextMenuStrip_open_alignment.Visible = true;
            }
            else
            {
                ContextMenuStrip_open_alignment.Visible = false;
            }
        }


        private void dataGridView_drawings_Click(object sender, EventArgs e)
        {
            Type t = e.GetType();
            if (t.Equals(typeof(MouseEventArgs)))
            {
                MouseEventArgs mouse = (MouseEventArgs)e;
                if (mouse.Button == MouseButtons.Right)
                {
                    ContextMenuStrip_open_alignment.Show(Cursor.Position);
                    ContextMenuStrip_open_alignment.Visible = true;
                }
            }
            else
            {
                ContextMenuStrip_open_alignment.Visible = false;
            }
        }


        private void button_select_viewport_Click(object sender, EventArgs e)
        {

            if (Freeze_operations == false)
            {
                ObjectId[] Empty_array = null;
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                try
                {
                    Freeze_operations = true;
                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;


                            Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rez.MessageForAdding = "\nSelect the viewport:";
                            Prompt_rez.SingleOnly = false;
                            Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                            if (Rezultat1.Status != PromptStatus.OK)
                            {
                                Freeze_operations = false;
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }
                            ObjectId vpId = ObjectId.Null;

                            Point3dCollection colpt = new Point3dCollection();

                            for (int i = 0; i < Rezultat1.Value.Count; ++i)
                            {
                                Viewport vp0 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as Viewport;
                                Polyline polyclip = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as Polyline;


                                if (vp0 != null)
                                {
                                    vpId = vp0.ObjectId;

                                    Point3d pt0 = new Point3d(vp0.CenterPoint.X - vp0.Width / 2, vp0.CenterPoint.Y - vp0.Height / 2, 0);
                                    Point3d pt1 = new Point3d(vp0.CenterPoint.X - vp0.Width / 2, vp0.CenterPoint.Y + vp0.Height / 2, 0);
                                    Point3d pt2 = new Point3d(vp0.CenterPoint.X + vp0.Width / 2, vp0.CenterPoint.Y + vp0.Height / 2, 0);
                                    Point3d pt3 = new Point3d(vp0.CenterPoint.X + vp0.Width / 2, vp0.CenterPoint.Y - vp0.Height / 2, 0);

                                    colpt.Add(pt0);
                                    colpt.Add(pt1);
                                    colpt.Add(pt2);
                                    colpt.Add(pt3);

                                    i = Rezultat1.Value.Count;
                                }

                                if (vp0 == null && polyclip != null)
                                {

                                    if (LayoutManager.Current.GetNonRectangularViewportIdFromClipId(polyclip.ObjectId) != ObjectId.Null)
                                    {
                                        vpId = LayoutManager.Current.GetNonRectangularViewportIdFromClipId(polyclip.ObjectId);
                                        for (int j = 0; j < polyclip.NumberOfVertices; ++j)
                                        {
                                            colpt.Add(polyclip.GetPointAtParameter(j));
                                        }

                                        i = Rezultat1.Value.Count;
                                    }

                                }
                            }
                            if (vpId != ObjectId.Null && colpt.Count > 0)
                            {
                                Viewport vp1 = Trans1.GetObject(vpId, OpenMode.ForRead) as Viewport;

                                if (vp1 != null)
                                {
                                    Matrix3d matrix1 = Functions.PaperToModel(vp1);
                                    Polyline poly_vp = new Polyline();

                                    for (int i = 0; i < colpt.Count; ++i)
                                    {
                                        Point3d pt0 = colpt[i].TransformBy(matrix1);
                                        poly_vp.AddVertexAt(i, new Point2d(pt0.X, pt0.Y), 0, 0, 0);
                                    }

                                    string filename = ThisDrawing.Database.OriginalFileName;

                                    if (System.IO.File.Exists(filename) == true)
                                    {
                                        filename = System.IO.Path.GetFileNameWithoutExtension(filename);
                                    }

                                    poly_vp.Closed = true;
                                    dt_vp.Rows.Add();
                                    dt_vp.Rows[dt_vp.Rows.Count - 1][0] = filename;
                                    dt_vp.Rows[dt_vp.Rows.Count - 1][1] = poly_vp;

                                    System.Data.DataTable dt_display = new System.Data.DataTable();
                                    dt_display.Columns.Add("DWG name", typeof(string));

                                    for (int i = 0; i < dt_vp.Rows.Count; ++i)
                                    {
                                        dt_display.Rows.Add();
                                        dt_display.Rows[dt_display.Rows.Count - 1][0] = dt_vp.Rows[i][0];
                                    }

                                    dataGridView_drawings.DataSource = dt_display;
                                    dataGridView_drawings.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                                    dataGridView_drawings.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                                    dataGridView_drawings.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                                    dataGridView_drawings.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                                    dataGridView_drawings.DefaultCellStyle.ForeColor = Color.White;
                                    dataGridView_drawings.EnableHeadersVisualStyles = false;


                                    Trans1.Commit();
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
                Freeze_operations = false;
            }

        }

        private void button_refresh_layers_Click(object sender, EventArgs e)
        {
            Functions.Incarca_existing_layers_to_combobox(comboBox_layers);
        }

        private void button_draw_rectangles_Click(object sender, EventArgs e)
        {

            if (Freeze_operations == false && dt_vp.Rows.Count > 0 && comboBox_layers.Text != "")
            {
                string nume_layer = comboBox_layers.Text;

                ObjectId[] Empty_array = null;
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                try
                {
                    Freeze_operations = true;
                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                            Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                            Functions.Create_vp_grab_od_table(nume_layer);

                            for (int i = 0; i < dt_vp.Rows.Count; ++i)
                            {
                                if (dt_vp.Rows[i][0] != DBNull.Value && dt_vp.Rows[i][1] != DBNull.Value)
                                {
                                    string dwg_name = Convert.ToString(dt_vp.Rows[i][0]);
                                    Polyline poly1 = new Polyline();
                                    poly1 = (Polyline)dt_vp.Rows[i][1];

                                    Polyline poly0 = new Polyline();
                                    for (int k = 0; k < poly1.NumberOfVertices; ++k)
                                    {
                                        poly0.AddVertexAt(k, poly1.GetPoint2dAt(k), 0, 0, 0);
                                    }
                                    poly0.Closed = true;
                                    poly0.Elevation = 0;
                                    poly0.Layer = nume_layer;
                                    poly0.ColorIndex = 256;
                                    poly0.Linetype = "BYLAYER";
                                    poly0.LineWeight = LineWeight.ByLayer;

                                    BTrecord.AppendEntity(poly0);
                                    Trans1.AddNewlyCreatedDBObject(poly0, true);


                                    List<object> Lista_val = new List<object>();
                                    List<Autodesk.Gis.Map.Constants.DataType> Lista_type = new List<Autodesk.Gis.Map.Constants.DataType>();
                                    Lista_val.Add(dwg_name);
                                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                                    Lista_val.Add("");
                                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                    Lista_val.Add(System.DateTime.Today.Year + "-" + System.DateTime.Today.Month + "-" + System.DateTime.Today.Day + " at " + System.DateTime.Now.Hour + ":" + System.DateTime.Now.Minute);
                                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                    Functions.Populate_object_data_table_from_objectid(Tables1, poly0.ObjectId, nume_layer, Lista_val, Lista_type);

                                    MText mtext1 = new MText();
                                    mtext1.Contents = dwg_name;
                                    mtext1.TextHeight = 10;
                                    mtext1.Attachment = AttachmentPoint.MiddleCenter;
                                    mtext1.Location = poly0.GetPointAtParameter(1);
                                    mtext1.Layer = nume_layer;
                                    mtext1.ColorIndex = 256;
                                    BTrecord.AppendEntity(mtext1);
                                    Trans1.AddNewlyCreatedDBObject(mtext1, true);
                                    Functions.Populate_object_data_table_from_objectid(Tables1, mtext1.ObjectId, nume_layer, Lista_val, Lista_type);
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
                Freeze_operations = false;
            }

        }
    }
}
