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
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Microsoft.Office.Interop.Excel;
using Autodesk.Gis.Map;
using Autodesk.Gis.Map.ImportExport;
using Autodesk.Gis.Map.ObjectData;
using Autodesk.Gis.Map.Project;
using System.Runtime.InteropServices;

namespace Alignment_mdi
{
    public partial class SGEN_Shape_Export : Form
    {

        System.Data.DataTable dt_layer = null;
        System.Data.DataTable dt_od = null;



        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_export);
            lista_butoane.Add(button_load_layers);
            lista_butoane.Add(button_browse_select_output_folder);


            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {

                bt1.Enabled = false;

            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();

            lista_butoane.Add(button_export);
            lista_butoane.Add(button_load_layers);
            lista_butoane.Add(button_browse_select_output_folder);
            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }

        public SGEN_Shape_Export()
        {
            InitializeComponent();
            if (Functions.is_dan_popescu() == true) textBox_output_folder.Text = "C:\\Users\\pop70694\\Documents\\Work Files\\2021-05-09 SGEN\\shp";
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

        private void button_load_layers_Click(object sender, EventArgs e)
        {
            List<string> lista1 = get_layers_from_dwg();
            dt_layer = new System.Data.DataTable();
            dt_layer.Columns.Add("Select", typeof(bool));
            dt_layer.Columns.Add("Name", typeof(string));
            dt_layer.Columns.Add("Export as Polygon", typeof(bool));
            dt_layer.Columns.Add("Export PVcase Blocks (ONLY)", typeof(bool));

            for (int i = 0; i < lista1.Count; ++i)
            {
                dt_layer.Rows.Add();
                dt_layer.Rows[dt_layer.Rows.Count - 1]["Select"] = false;
                dt_layer.Rows[dt_layer.Rows.Count - 1]["Name"] = lista1[i];
                dt_layer.Rows[dt_layer.Rows.Count - 1]["Export as Polygon"] = false;
                dt_layer.Rows[dt_layer.Rows.Count - 1]["Export PVcase Blocks (ONLY)"] = false;

            }

            if (lista1.Count > 0)
            {
                dataGridView_prop.DataSource = dt_layer;
                dataGridView_prop.Columns[2].Visible = false;
                dataGridView_prop.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                dataGridView_prop.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_prop.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                dataGridView_prop.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_prop.DefaultCellStyle.ForeColor = Color.White;
                dataGridView_prop.EnableHeadersVisualStyles = false;
            }
        }

        private List<string> get_layers_from_dwg()
        {
            List<string> lista_layers = new List<string>();

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument();
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                Autodesk.AutoCAD.DatabaseServices.LayerTable layer_table = (Autodesk.AutoCAD.DatabaseServices.LayerTable)Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                System.Data.DataTable dt1 = new System.Data.DataTable();
                dt1.Columns.Add("ln", typeof(string));
                foreach (ObjectId Layer_id in layer_table)
                {
                    LayerTableRecord Layer1 = (LayerTableRecord)Trans1.GetObject(Layer_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                    string Name_of_layer = Layer1.Name;
                    if (Name_of_layer.Contains("|") == false && Name_of_layer.Contains("$") == false && Layer1.IsFrozen == false && Layer1.IsOff == false)
                    {
                        dt1.Rows.Add();
                        dt1.Rows[dt1.Rows.Count - 1][0] = Name_of_layer;
                    }
                }
                System.Data.DataTable dt2 = Functions.Sort_data_table(dt1, "ln");
                for (int i = 0; i < dt2.Rows.Count; ++i)
                {
                    lista_layers.Add(dt2.Rows[i][0].ToString());
                }
                Trans1.Commit();
            }
            return lista_layers;
        }

        private void button_export_Click(object sender, EventArgs e)
        {
            List<Entity> lista_delete = new List<Entity>();

            string error_message = "";

            if (dt_layer != null && dt_layer.Rows.Count > 0)
            {
                List<string> lista_selected = new List<string>();
                List<string> lista_polygon_layers = new List<string>();
                List<string> lista_layers_with_blocks = new List<string>();
                for (int i = 0; i < dt_layer.Rows.Count; ++i)
                {
                    if ((bool)dt_layer.Rows[i][0] == true && (bool)dt_layer.Rows[i][2] == false)
                    {
                        lista_selected.Add(Convert.ToString(dt_layer.Rows[i][1]));
                    }
                    if ((bool)dt_layer.Rows[i][0] == true && (bool)dt_layer.Rows[i][2] == true)
                    {
                        lista_polygon_layers.Add(Convert.ToString(dt_layer.Rows[i][1]));
                    }
                    if ((bool)dt_layer.Rows[i][0] == true && (bool)dt_layer.Rows[i][3] == true)
                    {
                        lista_layers_with_blocks.Add(Convert.ToString(dt_layer.Rows[i][1]));
                    }
                }

                if (lista_selected.Count > 0 || lista_polygon_layers.Count > 0 || lista_layers_with_blocks.Count > 0)
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
                                BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;


                                System.Data.DataTable dt1 = new System.Data.DataTable();
                                dt1.Columns.Add("id", typeof(ObjectId));
                                dt1.Columns.Add("layer", typeof(string));

                                System.Data.DataTable dt2 = new System.Data.DataTable();
                                dt2.Columns.Add("id", typeof(ObjectId));
                                dt2.Columns.Add("layer", typeof(string));

                                System.Data.DataTable dt3 = new System.Data.DataTable();
                                dt3.Columns.Add("id", typeof(ObjectId));
                                dt3.Columns.Add("layer", typeof(string));

                                System.Data.DataTable dt4 = new System.Data.DataTable();
                                dt4.Columns.Add("id", typeof(ObjectId));
                                dt4.Columns.Add("blockname", typeof(string));
                                dt4.Columns.Add("layer", typeof(string));




                                Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                                Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                                Prompt_rez.MessageForAdding = "\nSelect the objects:";
                                Prompt_rez.SingleOnly = false;

                                bool open_poly_in_Polygon_layer = false;
                                bool linie_in_Polygon_layer = false;
                                bool point_in_line_layer = false;


                                List<string> lista_polylines_layers = new List<string>();
                                List<string> lista_point_layers = new List<string>();





                                if (Functions.is_dan_popescu() == true && DateTime.Now.Day == 122)
                                {
                                    #region dan popescu
                                    Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);
                                    if (Rezultat1.Status != PromptStatus.OK)
                                    {
                                        foreach (ObjectId id1 in BTrecord)
                                        {
                                            Entity Ent1 = Trans1.GetObject(id1, OpenMode.ForRead) as Entity;
                                            if (Ent1 is Curve && lista_selected.Contains(Ent1.Layer) == true)
                                            {
                                                dt1.Rows.Add();
                                                dt1.Rows[dt1.Rows.Count - 1][0] = id1;
                                                dt1.Rows[dt1.Rows.Count - 1][1] = Ent1.Layer;
                                            }

                                            if ((Ent1 is DBPoint || Ent1 is BlockReference) && lista_selected.Contains(Ent1.Layer) == true)
                                            {
                                                dt2.Rows.Add();
                                                dt2.Rows[dt2.Rows.Count - 1][0] = id1;
                                                dt2.Rows[dt2.Rows.Count - 1][1] = Ent1.Layer;
                                            }

                                        }



                                    }
                                    else
                                    {
                                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                                        {

                                            Entity Ent1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as Entity;
                                            if (Ent1 is Curve && lista_selected.Contains(Ent1.Layer) == true)
                                            {
                                                dt1.Rows.Add();
                                                dt1.Rows[dt1.Rows.Count - 1][0] = Rezultat1.Value[i].ObjectId;
                                                dt1.Rows[dt1.Rows.Count - 1][1] = Ent1.Layer;

                                            }
                                            if ((Ent1 is DBPoint || Ent1 is BlockReference) && lista_selected.Contains(Ent1.Layer) == true)
                                            {
                                                dt2.Rows.Add();
                                                dt2.Rows[dt2.Rows.Count - 1][0] = Rezultat1.Value[i].ObjectId;
                                                dt2.Rows[dt2.Rows.Count - 1][1] = Ent1.Layer;
                                            }

                                        }
                                    }
                                    #endregion
                                }
                                else
                                {
                                    foreach (ObjectId id1 in BTrecord)
                                    {
                                        Entity Ent1 = Trans1.GetObject(id1, OpenMode.ForRead) as Entity;


                                        if (lista_polygon_layers.Contains(Ent1.Layer) == true)
                                        {
                                            if (Ent1 is Polyline)
                                            {
                                                Polyline poly1 = Ent1 as Polyline;
                                                if (poly1.Closed == true)
                                                {
                                                    dt3.Rows.Add();
                                                    dt3.Rows[dt3.Rows.Count - 1][0] = id1;
                                                    dt3.Rows[dt3.Rows.Count - 1][1] = Ent1.Layer;
                                                }
                                                else
                                                {
                                                    open_poly_in_Polygon_layer = true;
                                                }
                                            }
                                            else if (Ent1 is MPolygon)
                                            {
                                                MPolygon mpolyg1 = Ent1 as MPolygon;
                                                dt3.Rows.Add();
                                                dt3.Rows[dt3.Rows.Count - 1][0] = id1;
                                                dt3.Rows[dt3.Rows.Count - 1][1] = Ent1.Layer;
                                            }
                                            else
                                            {
                                                linie_in_Polygon_layer = true;
                                            }
                                        }

                                        if (lista_selected.Contains(Ent1.Layer) == true)
                                        {
                                            if (Ent1 is Curve && lista_point_layers.Contains(Ent1.Layer) == false)
                                            {
                                                dt1.Rows.Add();
                                                dt1.Rows[dt1.Rows.Count - 1][0] = id1;
                                                dt1.Rows[dt1.Rows.Count - 1][1] = Ent1.Layer;
                                                if (lista_polylines_layers.Contains(Ent1.Layer) == false) lista_polylines_layers.Add(Ent1.Layer);
                                            }
                                            else if (Ent1 is Curve && lista_point_layers.Contains(Ent1.Layer) == true)
                                            {
                                                point_in_line_layer = true;
                                            }

                                            if ((Ent1 is DBPoint || Ent1 is BlockReference) && lista_polylines_layers.Contains(Ent1.Layer) == false && lista_layers_with_blocks.Contains(Ent1.Layer) == false)
                                            {
                                                dt2.Rows.Add();
                                                dt2.Rows[dt2.Rows.Count - 1][0] = id1;
                                                dt2.Rows[dt2.Rows.Count - 1][1] = Ent1.Layer;
                                                if (lista_point_layers.Contains(Ent1.Layer) == false) lista_point_layers.Add(Ent1.Layer);
                                            }
                                            else if ((Ent1 is DBPoint || Ent1 is BlockReference) && lista_polylines_layers.Contains(Ent1.Layer) == true && lista_layers_with_blocks.Contains(Ent1.Layer) == false)
                                            {
                                                point_in_line_layer = true;
                                            }

                                        }



                                        using (BlockReference bl2 = Trans1.GetObject(id1, OpenMode.ForRead) as BlockReference)
                                        {
                                            if (bl2 != null)
                                            {
                                                string bn = Functions.get_block_name(bl2);
                                                if (bn != "" && lista_layers_with_blocks.Contains(bl2.Layer) == true)
                                                {

                                                    error_message = bn;
                                                    dt4.Rows.Add();
                                                    dt4.Rows[dt4.Rows.Count - 1][0] = id1;
                                                    dt4.Rows[dt4.Rows.Count - 1][1] = bn;
                                                    dt4.Rows[dt4.Rows.Count - 1][2] = bl2.Layer;

                                                }
                                            }


                                        }
                                        Ent1.Dispose();
                                    }
                                }



                                if (Functions.is_dan_popescu() == true && DateTime.Now.Day == 122)
                                {
                                    #region dan popescu
                                    if (dt1.Rows.Count > 0)
                                    {
                                        for (int i = 0; i < lista_selected.Count; ++i)
                                        {
                                            ObjectIdCollection col_filter_by_layer = new ObjectIdCollection();
                                            for (int j = 0; j < dt1.Rows.Count; ++j)
                                            {
                                                string layerName1 = Convert.ToString(dt1.Rows[j][1]);
                                                if (layerName1 == lista_selected[i])
                                                {
                                                    col_filter_by_layer.Add((ObjectId)dt1.Rows[j][0]);
                                                }
                                            }
                                            string filename = textBox_output_folder.Text;

                                            if (filename.Substring(filename.Length - 1, 1) != "\\")
                                            {
                                                filename = filename + "\\";
                                            }

                                            if (System.IO.Directory.Exists(filename) == true)
                                            {
                                                int incr = 0;
                                                string suff1 = "";
                                                bool exista = true;
                                                do
                                                {

                                                    if (System.IO.File.Exists(filename + lista_selected[i] + suff1 + ".shp") == false)
                                                    {
                                                        filename = filename + lista_selected[i] + suff1 + ".shp";
                                                        exista = false;
                                                    }
                                                    else
                                                    {

                                                        ++incr;
                                                        suff1 = incr.ToString();
                                                    }

                                                } while (exista == true);


                                            }
                                            ExportSHP("SHP", filename, lista_selected[i], true, false, "line", col_filter_by_layer);

                                            for (int j = dt1.Rows.Count - 1; j >= 0; --j)
                                            {
                                                string layerName1 = Convert.ToString(dt1.Rows[j][1]);
                                                if (layerName1 == lista_selected[i])
                                                {
                                                    dt1.Rows[j].Delete();
                                                }
                                            }
                                        }

                                    }

                                    if (dt2.Rows.Count > 0)
                                    {
                                        for (int i = 0; i < lista_selected.Count; ++i)
                                        {
                                            ObjectIdCollection col_filter_by_layer = new ObjectIdCollection();
                                            for (int j = 0; j < dt2.Rows.Count; ++j)
                                            {
                                                string layerName1 = Convert.ToString(dt2.Rows[j][1]);
                                                if (layerName1 == lista_selected[i])
                                                {
                                                    col_filter_by_layer.Add((ObjectId)dt2.Rows[j][0]);
                                                }
                                            }
                                            string filename = textBox_output_folder.Text;

                                            if (filename.Substring(filename.Length - 1, 1) != "\\")
                                            {
                                                filename = filename + "\\";
                                            }

                                            if (System.IO.Directory.Exists(filename) == true)
                                            {
                                                int incr = 0;
                                                string suff1 = "";
                                                bool exista = true;
                                                do
                                                {

                                                    if (System.IO.File.Exists(filename + lista_selected[i] + suff1 + ".shp") == false)
                                                    {
                                                        filename = filename + lista_selected[i] + suff1 + ".shp";
                                                        exista = false;
                                                    }
                                                    else
                                                    {

                                                        ++incr;
                                                        suff1 = incr.ToString();
                                                    }

                                                } while (exista == true);


                                            }
                                            ExportSHP("SHP", filename, lista_selected[i], true, false, "point", col_filter_by_layer);

                                            for (int j = dt2.Rows.Count - 1; j >= 0; --j)
                                            {
                                                string layerName1 = Convert.ToString(dt2.Rows[j][1]);
                                                if (layerName1 == lista_selected[i])
                                                {
                                                    dt2.Rows[j].Delete();
                                                }
                                            }
                                        }

                                    }



                                    #endregion

                                }
                                else
                                {
                                    #region others

                                    if (linie_in_Polygon_layer == false && open_poly_in_Polygon_layer == false && point_in_line_layer == false)
                                    {
                                        if (dt1.Rows.Count > 0)
                                        {
                                            for (int i = 0; i < lista_polylines_layers.Count; ++i)
                                            {
                                                ObjectIdCollection col_filter_by_layer = new ObjectIdCollection();


                                                for (int j = 0; j < dt1.Rows.Count; ++j)
                                                {
                                                    string layerName1 = Convert.ToString(dt1.Rows[j][1]);
                                                    if (layerName1 == lista_polylines_layers[i])
                                                    {
                                                        col_filter_by_layer.Add((ObjectId)dt1.Rows[j][0]);
                                                    }
                                                }
                                                string filename = textBox_output_folder.Text;

                                                if (filename.Substring(filename.Length - 1, 1) != "\\")
                                                {
                                                    filename = filename + "\\";
                                                }

                                                if (System.IO.Directory.Exists(filename) == true)
                                                {
                                                    int incr = 0;
                                                    string suff1 = "";
                                                    bool exista = true;
                                                    do
                                                    {

                                                        if (System.IO.File.Exists(filename + lista_polylines_layers[i] + suff1 + ".shp") == false)
                                                        {
                                                            filename = filename + lista_polylines_layers[i] + suff1 + ".shp";
                                                            exista = false;
                                                        }
                                                        else
                                                        {

                                                            ++incr;
                                                            suff1 = incr.ToString();
                                                        }

                                                    } while (exista == true);


                                                }
                                                ExportSHP("SHP", filename, lista_polylines_layers[i], true, false, "line", col_filter_by_layer);

                                                for (int j = dt1.Rows.Count - 1; j >= 0; --j)
                                                {
                                                    string layerName1 = Convert.ToString(dt1.Rows[j][1]);
                                                    if (layerName1 == lista_polylines_layers[i])
                                                    {
                                                        dt1.Rows[j].Delete();
                                                    }
                                                }
                                            }

                                        }
                                        if (dt2.Rows.Count > 0)
                                        {
                                            for (int i = 0; i < lista_point_layers.Count; ++i)
                                            {
                                                ObjectIdCollection col_filter_by_layer = new ObjectIdCollection();
                                                for (int j = 0; j < dt2.Rows.Count; ++j)
                                                {
                                                    string layerName1 = Convert.ToString(dt2.Rows[j][1]);
                                                    if (layerName1 == lista_point_layers[i])
                                                    {
                                                        col_filter_by_layer.Add((ObjectId)dt2.Rows[j][0]);
                                                    }
                                                }
                                                string filename = textBox_output_folder.Text;

                                                if (filename.Substring(filename.Length - 1, 1) != "\\")
                                                {
                                                    filename = filename + "\\";
                                                }

                                                if (System.IO.Directory.Exists(filename) == true)
                                                {
                                                    int incr = 0;
                                                    string suff1 = "";
                                                    bool exista = true;
                                                    do
                                                    {

                                                        if (System.IO.File.Exists(filename + lista_point_layers[i] + suff1 + ".shp") == false)
                                                        {
                                                            filename = filename + lista_point_layers[i] + suff1 + ".shp";
                                                            exista = false;
                                                        }
                                                        else
                                                        {

                                                            ++incr;
                                                            suff1 = incr.ToString();
                                                        }

                                                    } while (exista == true);


                                                }
                                                ExportSHP("SHP", filename, lista_point_layers[i], true, false, "point", col_filter_by_layer);

                                                for (int j = dt2.Rows.Count - 1; j >= 0; --j)
                                                {
                                                    string layerName1 = Convert.ToString(dt2.Rows[j][1]);
                                                    if (layerName1 == lista_point_layers[i])
                                                    {
                                                        dt2.Rows[j].Delete();
                                                    }
                                                }
                                            }

                                        }
                                        if (dt3.Rows.Count > 0)
                                        {

                                            for (int i = 0; i < lista_polygon_layers.Count; ++i)
                                            {
                                                ObjectIdCollection col_filter_by_layer = new ObjectIdCollection();
                                                for (int j = 0; j < dt3.Rows.Count; ++j)
                                                {
                                                    string layerName1 = Convert.ToString(dt3.Rows[j][1]);
                                                    if (layerName1 == lista_polygon_layers[i])
                                                    {
                                                        col_filter_by_layer.Add((ObjectId)dt3.Rows[j][0]);
                                                    }
                                                }

                                                ObjectIdCollection col_polygon = new ObjectIdCollection();
                                                if (col_filter_by_layer.Count > 0)
                                                {
                                                    for (int k = 0; k < col_filter_by_layer.Count; ++k)
                                                    {
                                                        Polyline polyline2 = Trans1.GetObject(col_filter_by_layer[k], OpenMode.ForRead) as Polyline;
                                                        MPolygon mpolyg1 = Trans1.GetObject(col_filter_by_layer[k], OpenMode.ForRead) as MPolygon;

                                                        if (polyline2 != null)
                                                        {
                                                            MPolygon mpolygon1 = new MPolygon();
                                                            mpolygon1.AppendLoopFromBoundary(polyline2, true, 1e-12);
                                                            mpolygon1.BalanceTree();
                                                            mpolygon1.Elevation = 0;
                                                            mpolygon1.Normal = polyline2.Normal;
                                                            BTrecord.AppendEntity(mpolygon1);
                                                            Trans1.AddNewlyCreatedDBObject(mpolygon1, true);

                                                            copy_od(polyline2.ObjectId, mpolygon1.ObjectId);
                                                            col_polygon.Add(mpolygon1.ObjectId);
                                                        }

                                                        if (mpolyg1 != null)
                                                        {
                                                            col_polygon.Add(mpolyg1.ObjectId);
                                                        }
                                                    }
                                                }


                                                string filename = textBox_output_folder.Text;

                                                if (filename.Substring(filename.Length - 1, 1) != "\\")
                                                {
                                                    filename = filename + "\\";
                                                }

                                                if (System.IO.Directory.Exists(filename) == true)
                                                {
                                                    int incr = 0;
                                                    string suff1 = "";
                                                    bool exista = true;
                                                    do
                                                    {

                                                        if (System.IO.File.Exists(filename + lista_polygon_layers[i] + suff1 + ".shp") == false)
                                                        {
                                                            filename = filename + lista_polygon_layers[i] + suff1 + ".shp";
                                                            exista = false;
                                                        }
                                                        else
                                                        {

                                                            ++incr;
                                                            suff1 = incr.ToString();
                                                        }

                                                    } while (exista == true);


                                                }
                                                ExportSHP("SHP", filename, lista_polygon_layers[i], true, false, "polygon", col_polygon);



                                                for (int j = dt3.Rows.Count - 1; j >= 0; --j)
                                                {
                                                    string layerName1 = Convert.ToString(dt3.Rows[j][1]);
                                                    if (layerName1 == lista_polygon_layers[i])
                                                    {
                                                        dt3.Rows[j].Delete();
                                                    }
                                                }

                                                for (int k = 0; k < col_polygon.Count; ++k)
                                                {
                                                    MPolygon mp1 = Trans1.GetObject(col_polygon[k], OpenMode.ForWrite) as MPolygon;
                                                    mp1.Erase();
                                                }
                                            }



                                        }

                                        if (dt4.Rows.Count > 0)
                                        {


                                            dt_od = new System.Data.DataTable();
                                            dt_od.Columns.Add("id", typeof(ObjectId));
                                            dt_od.Columns.Add("bn", typeof(string));

                                            Create_blockname_object_data();



                                            List<string> lista_layere_from_dt4 = new List<string>();

                                            for (int i = 0; i < dt4.Rows.Count; ++i)
                                            {
                                                string layerName1 = Convert.ToString(dt4.Rows[i][2]);
                                                if (lista_layere_from_dt4.Contains(layerName1) == false) lista_layere_from_dt4.Add(layerName1);
                                            }

                                            for (int k = 0; k < lista_layere_from_dt4.Count; ++k)
                                            {
                                                string layer1 = lista_layere_from_dt4[k];
                                                ObjectIdCollection col_new_linework = new ObjectIdCollection();

                                                for (int i = 0; i < dt4.Rows.Count; ++i)
                                                {
                                                    if (dt4.Rows[i][0] != DBNull.Value && dt4.Rows[i][1] != DBNull.Value && dt4.Rows[i][2] != DBNull.Value)
                                                    {
                                                        ObjectId id1 = (ObjectId)dt4.Rows[i][0];

                                                        string nume1 = Convert.ToString(dt4.Rows[i][1]);
                                                        string layer2 = Convert.ToString(dt4.Rows[i][2]);

                                                        if (layer1 == layer2)
                                                        {
                                                            BlockReference br1 = Trans1.GetObject(id1, OpenMode.ForRead) as BlockReference;

                                                            Extents3d ext1 = br1.Bounds.Value;
                                                            double width = ext1.MaxPoint.X - ext1.MinPoint.X;
                                                            double height = ext1.MaxPoint.Y - ext1.MinPoint.Y;
                                                            Polyline poly1 = new Polyline();

                                                            poly1.AddVertexAt(0, new Point2d(br1.Position.X, br1.Position.Y), 0, 0, 0);
                                                            poly1.AddVertexAt(1, new Point2d(br1.Position.X, br1.Position.Y - height), 0, 0, 0);
                                                            poly1.AddVertexAt(2, new Point2d(br1.Position.X - width, br1.Position.Y - height), 0, 0, 0);
                                                            poly1.AddVertexAt(3, new Point2d(br1.Position.X - width, br1.Position.Y), 0, 0, 0);
                                                            poly1.Closed = true;

                                                            BTrecord.AppendEntity(poly1);
                                                            Trans1.AddNewlyCreatedDBObject(poly1, true);

                                                            col_new_linework.Add(poly1.ObjectId);
                                                            lista_delete.Add(poly1);
                                                            dt_od.Rows.Add();
                                                            dt_od.Rows[dt_od.Rows.Count - 1][0] = poly1.ObjectId;
                                                            dt_od.Rows[dt_od.Rows.Count - 1][1] = nume1;
                                                        }
                                                    }
                                                }

                                                Append_blockname_object_data(dt_od);

                                                string filename = textBox_output_folder.Text;

                                                if (filename.Substring(filename.Length - 1, 1) != "\\")
                                                {
                                                    filename = filename + "\\";
                                                }

                                                if (System.IO.Directory.Exists(filename) == true)
                                                {
                                                    int incr = 0;
                                                    string suff1 = "";
                                                    bool exista = true;
                                                    do
                                                    {
                                                        if (System.IO.File.Exists(filename + layer1 + suff1 + ".shp") == false)
                                                        {
                                                            filename = filename + layer1 + suff1 + ".shp";
                                                            exista = false;
                                                        }
                                                        else
                                                        {
                                                            ++incr;
                                                            suff1 = incr.ToString();
                                                        }

                                                    } while (exista == true);
                                                }
                                                ExportSHP("SHP", filename, "BNAME", true, false, "line", col_new_linework);
                                            }
                                        }
                                    }


                                    if (point_in_line_layer == true)
                                    {
                                        MessageBox.Show("Operation aborted!\r\nyou have lines into the point layers or points inside lines layers");
                                    }


                                    if (open_poly_in_Polygon_layer == true)
                                    {
                                        MessageBox.Show("Operation aborted!\r\nyou have at least one open polyline that you want to export it as a polygon");
                                    }

                                    if (linie_in_Polygon_layer == true)
                                    {
                                        MessageBox.Show("Operation aborted!\r\nyou have at least one item that is not a polyline that you want to export it as a polygon");
                                    }


                                    #endregion
                                }

                                if (lista_delete.Count > 0)
                                {
                                    for (int i = 0; i < lista_delete.Count; ++i)
                                    {
                                        Entity ent1 = lista_delete[i];
                                        ent1.Erase();
                                    }
                                }



                                Trans1.Commit();

                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.Message + "\r\n" + error_message);
                    }

                    Editor1.SetImpliedSelection(Empty_array);
                    Editor1.WriteMessage("\nCommand:");
                    set_enable_true();
                }

            }

        }



        public void ExportSHP(string format, string filename, string shpname, bool isODTable, bool isLinkTemplate, string geomtryType, ObjectIdCollection object_id_col)
        {
            Exporter exporter = null;
            try
            {
                MapApplication mapApp = HostMapApplicationServices.Application;
                exporter = mapApp.Exporter;
                // Initiate the exporter
                exporter.Init(format, filename);
                if (object_id_col != null && object_id_col.Count > 0)
                {
                    exporter.ExportAll = false;

                    exporter.SetSelectionSet(object_id_col);
                }
                else
                {
                    //exporter.ExportAll = true;
                }

                GeometryType geomtryType1 = (GeometryType)Enum.Parse(typeof(GeometryType), geomtryType, true);
                exporter.SetStorageOptions(StorageType.FileOneEntityType, geomtryType1, string.Empty);

                // Get Data mapping object
                ExpressionTargetCollection dataMapping = null;
                dataMapping = exporter.GetExportDataMappings();
                dataMapping.Clear();

                // Set ObjectData data mapping if isODTable is true
                if (isODTable == true && MapODData(dataMapping, shpname) == true)
                {
                    // Reset Data mapping with Object data and Link template keys		
                    exporter.SetExportDataMappings(dataMapping);
                }
                // If layerFilter isn't null, set the layer filter to export layer by layer
                if (null != shpname)
                {
                    exporter.LayerFilter = shpname;
                }

                // Do the exporting and log the result
                ExportResults results;
                results = exporter.Export(true);

                Utility.ShowMsg("\nExporting succeeded.");
            }
            catch (MapException e)
            {

            }
            finally
            {

            }
        }
        public bool MapODData(ExpressionTargetCollection mapping, string tablename)
        {
            MapApplication mapApi = HostMapApplicationServices.Application;
            ProjectModel proj = mapApi.ActiveProject;
            Tables tables = proj.ODTables;
            if (tables.IsTableDefined(tablename) == true)
            {
                try
                {
                    Autodesk.Gis.Map.ObjectData.Table table = tables[tablename];
                    FieldDefinitions definitions = table.FieldDefinitions;
                    for (int j = 0; j < definitions.Count; j++)
                    {
                        FieldDefinition column = null;
                        column = definitions[j];
                        // fieldName is the OD table field name in the data mapping. It should be 
                        // in the format:fieldName&tableName. 
                        // newFieldName is the attribute field name of exported-to file
                        string ODfieldName = ":" + column.Name + "@" + tablename;
                        string shpFieldName = column.Name;

                        mapping.Add(ODfieldName, shpFieldName);
                    }
                }
                catch (MapImportExportException)
                {
                    return false;
                }
                return true;
            }
            else
            {
                return false;
            }

        }

        public void copy_od(ObjectId id1, ObjectId id2)
        {
            System.Data.DataTable Data_table_for_object_data = new System.Data.DataTable();

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                    Autodesk.Gis.Map.ObjectData.Records Records1;
                    Autodesk.Gis.Map.ObjectData.Records Records2;

                    Entity Ent1 = Trans1.GetObject(id1, OpenMode.ForRead) as Entity;
                    Entity Ent2 = Trans1.GetObject(id2, OpenMode.ForRead) as Entity;

                    if (Ent1 != null & Ent2 != null & Ent1.ObjectId != Ent2.ObjectId)
                    {
                        try
                        {
                            using (Records2 = Tables1.GetObjectRecords(Convert.ToUInt32(0), Ent2.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForWrite, false))
                            {
                                if (Records2 != null)
                                {
                                    if (Records2.Count > 0)
                                    {
                                        System.Collections.IEnumerator ie = Records2.GetEnumerator();
                                        while (ie.MoveNext())
                                        {
                                            Records2.RemoveRecord();
                                        }
                                    }
                                }
                            }

                            using (Records1 = Tables1.GetObjectRecords(Convert.ToUInt32(0), Ent1.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, false))
                            {
                                if (Records1 != null)
                                {
                                    if (Records1.Count > 0)
                                    {
                                        Data_table_for_object_data.Rows.Add();
                                        foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                        {
                                            Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Record1.TableName];
                                            Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;
                                            Tabla1.AddRecord(Record1, Ent2.ObjectId);
                                            for (int j = 0; j < Record1.Count; ++j)
                                            {
                                                Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[j];
                                                string Nume_field = Field_def1.Name;
                                                string Valoare_field = (string)Record1[j].StrValue;
                                                if (Data_table_for_object_data.Columns.Contains(Nume_field) == false)
                                                {
                                                    Data_table_for_object_data.Columns.Add(Nume_field, typeof(String));
                                                }
                                                Data_table_for_object_data.Rows[0][Nume_field] = Valoare_field;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        catch (AccessViolationException ex1)
                        {
                            MessageBox.Show(ex1.Message);
                        }
                    }
                    Trans1.Commit();
                }
            }
        }


        private void Create_blockname_object_data()
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


                            List1.Add("BNAME");
                            List2.Add("Name");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);



                            Functions.Get_object_data_table("BNAME", "Generated by SGEN", List1, List2, List3);

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

        private void Append_blockname_object_data(System.Data.DataTable dt1)
        {
            if (dt1 != null && dt1.Rows.Count > 0)
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
                            BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                            for (int i = 0; i < dt1.Rows.Count; ++i)
                            {
                                if (dt1.Rows[i][0] != DBNull.Value && dt1.Rows[i][1] != DBNull.Value)
                                {
                                    ObjectId id_poly = (ObjectId)dt1.Rows[i][0];
                                    string nume_block = Convert.ToString(dt1.Rows[i][1]);
                                    List<object> Lista_val = new List<object>();
                                    List<Autodesk.Gis.Map.Constants.DataType> Lista_type = new List<Autodesk.Gis.Map.Constants.DataType>();
                                    Lista_val.Add(nume_block);
                                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                    Functions.Populate_object_data_table_from_objectid(Tables1, id_poly, "BNAME", Lista_val, Lista_type);
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

        //private Polyline get_outer_poly_from_block(BlockTableRecord)
        //{

        //}


    }
}
