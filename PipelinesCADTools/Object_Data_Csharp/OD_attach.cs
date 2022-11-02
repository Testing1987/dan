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

namespace Alignment_mdi
{
    public partial class OD_attach : Form
    {
        public OD_attach()
        {
            InitializeComponent();
        }

        System.Data.DataTable Data_table_layers;
        bool Freeze_operations = false;

        private void button_read_excel_column_Click(object sender, EventArgs e)
        {
            try
            {
                if (textBox_excel_column.Text == "")
                {
                    MessageBox.Show("No Excel Column");
                    return;
                }

                string ColXL = textBox_excel_column.Text.ToUpper();

                string start1_text = textBox_start.Text;

                if (Functions.IsNumeric(start1_text) == false)
                {
                    MessageBox.Show("No Start row");
                    return;
                }
                int Start1 = Convert.ToInt32(textBox_start.Text);

                string end1_text = textBox_end.Text;

                if (Functions.IsNumeric(end1_text) == false)
                {
                    MessageBox.Show("No End row");
                    return;
                }


                int End1 = Convert.ToInt32(textBox_end.Text);

                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                if (Freeze_operations == false)
                {
                    Freeze_operations = true;


                    Data_table_layers = new System.Data.DataTable();
                    Data_table_layers.Columns.Add("NAME", typeof(string));
                    int idx = 0;
                    Worksheet W1 = Alignment_mdi.Functions.Get_active_worksheet_from_Excel();

                    for (int i = Start1; i <= End1; i = i + 1)
                    {
                        String NameLayerExcel = Convert.ToString(W1.Range[ColXL + i].Value2);
                        if (NameLayerExcel != "")
                        {
                            Data_table_layers.Rows.Add();
                            Data_table_layers.Rows[idx]["NAME"] = NameLayerExcel;
                            idx = idx + 1;

                        }

                    }

                    Freeze_operations = false;
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                Freeze_operations = false;
            }
        }

        private void button_ADD_TO_LAYERS_Click(object sender, EventArgs e)
        {

            string Nume_layer123 = "";
            try
            {
                try
                {
                    if (Data_table_layers != null)
                    {
                        if (Data_table_layers.Rows.Count > 0)
                        {
                            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                            using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                            {
                                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                                {
                                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                                    Autodesk.AutoCAD.DatabaseServices.LayerTable layer_table = (Autodesk.AutoCAD.DatabaseServices.LayerTable)Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);

                                    System.Collections.Specialized.StringCollection Colectie_layere = new System.Collections.Specialized.StringCollection();
                                    System.Collections.Specialized.StringCollection Colectie_layere_off = new System.Collections.Specialized.StringCollection();
                                    foreach (ObjectId Layer_id in layer_table)
                                    {
                                        LayerTableRecord Layer1 = (LayerTableRecord)Trans1.GetObject(Layer_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                                        string Nume_layer = Layer1.Name;
                                        bool OffL = true;
                                        for (int j = 0; j < Data_table_layers.Rows.Count; j = j + 1)
                                        {

                                            string NumeLayerExcel = (string)Data_table_layers.Rows[j]["NAME"];
                                            if (Nume_layer.ToUpper().Contains(NumeLayerExcel.ToUpper()) == true)
                                            {
                                                Colectie_layere.Add(Nume_layer);
                                                OffL = false;
                                                j = Data_table_layers.Rows.Count;
                                            }

                                        }

                                        if (OffL == true)
                                        {
                                            if (Nume_layer != "PNTDESC" & Nume_layer != "PNTNO")
                                            {
                                                if (Nume_layer.Length > 4)
                                                {
                                                    if (Nume_layer.Substring(0, 3) != "PT_")
                                                    {
                                                        Layer1.IsOff = true;
                                                    }
                                                }
                                                else
                                                {
                                                    Layer1.IsOff = true;
                                                }
                                            }
                                        }

                                    }

                                    if (Colectie_layere.Count > 0)
                                    {
                                        Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

                                        Autodesk.Gis.Map.ObjectData.Records Records1;
                                        foreach (ObjectId Obj_ID1 in BTrecord)
                                        {
                                            Entity Ent1 = (Entity)Trans1.GetObject(Obj_ID1, OpenMode.ForWrite);


                                            if (Ent1 != null)
                                            {

                                                string Nume_layer = Ent1.Layer;
                                                if (Colectie_layere.Contains(Nume_layer) == true)
                                                {


                                                    try
                                                    {

                                                        using (Records1 = Tables1.GetObjectRecords(Convert.ToUInt32(0), Ent1.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForWrite, true))
                                                        {

                                                            bool Exista_OD = false;
                                                            if (Records1 != null)
                                                            {
                                                                if (Records1.Count > 0)
                                                                {

                                                                    foreach (Autodesk.Gis.Map.ObjectData.Record Record2 in Records1)
                                                                    {
                                                                        Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Record2.TableName];

                                                                        if (Ent1.Layer.ToUpper().Contains(Tabla1.Name.ToUpper()) == false)
                                                                        {
                                                                            System.Collections.IEnumerator ie = Records1.GetEnumerator();
                                                                            while (ie.MoveNext())
                                                                            {
                                                                                Records1.RemoveRecord();
                                                                            }
                                                                        }
                                                                        else
                                                                        {
                                                                            Exista_OD = true;
                                                                        }


                                                                    }
                                                                }

                                                            }

                                                            if (Exista_OD == false)
                                                            {
                                                                Nume_layer123 = Nume_layer;
                                                                Autodesk.Gis.Map.ObjectData.Table Tabla1 = null;

                                                                try
                                                                {
                                                                    Tabla1 = Tables1[Nume_layer];
                                                                }
                                                                catch (Autodesk.Gis.Map.MapException ex)
                                                                {

                                                                }


                                                                if (Tabla1 != null)
                                                                {
                                                                    Autodesk.Gis.Map.ObjectData.Record Record1 = Autodesk.Gis.Map.ObjectData.Record.Create();
                                                                    Tabla1.InitRecord(Record1);
                                                                    Tabla1.AddRecord(Record1, Ent1.ObjectId);

                                                                }


                                                            }


                                                        }




                                                    }
                                                    catch (AccessViolationException ex1)
                                                    {
                                                        MessageBox.Show(ex1.Message);
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
                }

                catch (Autodesk.Gis.Map.MapException ex)
                {
                    MessageBox.Show(ex.Message + "\n" + Nume_layer123);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }


    }
}
