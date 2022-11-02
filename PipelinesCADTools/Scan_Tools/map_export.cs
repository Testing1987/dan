using System;
using System.Collections.Generic;
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
using System.Management;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.Gis.Map.ImportExport;
using Autodesk.Gis.Map;
using Autodesk.Gis.Map.Project;
using Autodesk.Gis.Map.ObjectData;
using System.Collections.Specialized;

namespace Alignment_mdi
{
    public partial class map_export : Form
    {
        System.Data.DataTable dt_layer = null;


        public map_export()
        {
            InitializeComponent();
        }

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







        private List<string> get_od_fields_from_dwg()
        {
            List<string> lista_fields = new List<string>();
            List<string> lista_OD = get_od_tables_from_dwg();

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                    for (int i = 0; i < lista_OD.Count; ++i)
                    {
                        if (Tables1.IsTableDefined(lista_OD[i]) == true)
                        {
                            Autodesk.Gis.Map.ObjectData.Table tabla1 = Tables1[lista_OD[i]];
                            Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = tabla1.FieldDefinitions;
                            for (int j = 0; j < Field_defs1.Count; ++j)
                            {
                                Autodesk.Gis.Map.ObjectData.FieldDefinition fielddef1 = Field_defs1[j];
                                lista_fields.Add(fielddef1.Name);
                            }
                        }
                    }



                    Trans1.Commit();
                }
            }

            return lista_fields;
        }

        private List<string> get_od_tables_from_dwg()
        {
            List<string> lista_OD = new List<string>();


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
                        System.Collections.Specialized.StringCollection Nume_tables = new System.Collections.Specialized.StringCollection();
                        Nume_tables = Tables1.GetTableNames();

                        for (int i = 0; i < Nume_tables.Count; i = i + 1)
                        {
                            lista_OD.Add(Nume_tables[i]);

                        }
                    }
                }

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return lista_OD;

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



        private void button_map_export_to_shp_Click(object sender, EventArgs e)
        {
            if (dt_layer != null && dt_layer.Rows.Count > 0)
            {
                List<string> lista_selected = new List<string>();
                for (int i = 0; i < dt_layer.Rows.Count; ++i)
                {

                    if ((bool)dt_layer.Rows[i][0] == true)
                    {
                        lista_selected.Add(Convert.ToString(dt_layer.Rows[i][1]));
                    }



                }
                if (lista_selected.Count > 0)
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
                                BlockTableRecord BTrecord_MS = Trans1.GetObject(BlockTable1[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;
                                BlockTableRecord BTrecord_PS = Trans1.GetObject(BlockTable1[BlockTableRecord.PaperSpace], OpenMode.ForRead) as BlockTableRecord;
                                BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;
                                LayerTable LayerTable1 = Trans1.GetObject(ThisDrawing.Database.LayerTableId, OpenMode.ForRead) as LayerTable;
                                TextStyleTable Text_style_table1 = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                                Autodesk.Gis.Map.MapApplication mapApp = Autodesk.Gis.Map.HostMapApplicationServices.Application;


                                Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                                Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                                Prompt_rez.MessageForAdding = "\nSelect the objects:";
                                Prompt_rez.SingleOnly = false;
                                Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);




                                System.Data.DataTable dt1 = new System.Data.DataTable();
                                dt1.Columns.Add("id", typeof(ObjectId));
                                dt1.Columns.Add("layer", typeof(string));

                                System.Data.DataTable dt2 = new System.Data.DataTable();
                                dt2.Columns.Add("id", typeof(ObjectId));
                                dt2.Columns.Add("layer", typeof(string));

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

                                        if ((Ent1 is DBPoint||Ent1 is BlockReference) && lista_selected.Contains(Ent1.Layer) == true)
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


        public void ExportSHP(string format, string filename, string layername, bool isODTable, bool isLinkTemplate, string geomtryType, ObjectIdCollection object_id_col)
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
                if (isODTable == true && MapODData(dataMapping, layername) == true)
                {
                    // Reset Data mapping with Object data and Link template keys		
                    exporter.SetExportDataMappings(dataMapping);
                }
                // If layerFilter isn't null, set the layer filter to export layer by layer
                if (null != layername)
                {
                    exporter.LayerFilter = layername;
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
        private void label_for_dan_Click(object sender, EventArgs e)
        {
            if (Functions.is_dan_popescu() == true)
            {

            }
        }



        private void button_load_layers_Click(object sender, EventArgs e)
        {
            List<string> lista1 = get_layers_from_dwg();
            dt_layer = new System.Data.DataTable();
            dt_layer.Columns.Add("Select", typeof(bool));
            dt_layer.Columns.Add("Name", typeof(string));
            for (int i = 0; i < lista1.Count; ++i)
            {
                dt_layer.Rows.Add();
                dt_layer.Rows[dt_layer.Rows.Count - 1]["Select"] = false;
                dt_layer.Rows[dt_layer.Rows.Count - 1]["Name"] = lista1[i];
            }

            if (lista1.Count > 0)
            {
                dataGridView_prop.DataSource = dt_layer;
                dataGridView_prop.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                dataGridView_prop.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_prop.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                dataGridView_prop.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_prop.DefaultCellStyle.ForeColor = Color.White;
                dataGridView_prop.EnableHeadersVisualStyles = false;
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
    }


    public sealed class Utility
    {
        private Utility()
        {
        }

        public static void ShowMsg(string msg)
        {
            AcadEditor.WriteMessage(msg);
        }

        public static Autodesk.AutoCAD.EditorInput.Editor AcadEditor
        {
            get
            {
                return Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            }
        }
    }


}
