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

namespace Alignment_mdi
{
    public partial class OD_TABLE_form : Form
    {
        int Row_current = 0;
        int Row_previous = 0;
        System.Data.DataTable Data_table_for_object_data;

        public OD_TABLE_form()
        {
            InitializeComponent();
        }

        private void button_LOAD_Click(object sender, EventArgs e)
        {
            try
            {


                DataGrid1.DataSource = null;
                DataGrid1.Refresh();

                Data_table_for_object_data = new System.Data.DataTable();

                Data_table_for_object_data.Columns.Add("OBJECT_ID", typeof(ObjectId));
                Data_table_for_object_data.Columns.Add("OBJECT_TYPE", typeof(String));
                Data_table_for_object_data.Columns.Add("LAYER_NAME", typeof(String));
                Data_table_for_object_data.Columns.Add("BLOCK_NAME", typeof(String));
                Data_table_for_object_data.Columns.Add("TABLE_NAME", typeof(String));

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

                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez.MessageForAdding = "\nSelect an object:";
                        Prompt_rez.SingleOnly = true;
                        Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            DataGrid1.DataSource = null;
                            DataGrid1.Refresh();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Entity Ent_layer = (Entity)Trans1.GetObject(Rezultat1.Value[0].ObjectId, OpenMode.ForRead);

                        int OD_id = 0;

                        foreach (ObjectId Obj_ID1 in BTrecord)
                        {
                            Entity Ent1 = (Entity)Trans1.GetObject(Obj_ID1, OpenMode.ForRead);
                            if (Ent1 != null)
                            {
                                if (Ent1.Layer == Ent_layer.Layer)
                                {
                                    try
                                    {
                                        Data_table_for_object_data.Rows.Add();

                                        if (Ent1 is BlockReference)
                                        {
                                            BlockReference Block1 = (BlockReference)Ent1;
                                            Data_table_for_object_data.Rows[OD_id]["BLOCK_NAME"] = Block1.Name;
                                        }

                                        Data_table_for_object_data.Rows[OD_id]["LAYER_NAME"] = Ent1.Layer;

                                        Data_table_for_object_data.Rows[OD_id]["OBJECT_ID"] = Ent1.ObjectId;

                                        string Type1 = Ent1.GetType().ToString();
                                        Type1 = Type1.Replace("Autodesk.AutoCAD.DatabaseServices.", "");
                                        Data_table_for_object_data.Rows[OD_id]["OBJECT_TYPE"] = Type1;

                                        using (Records1 = Tables1.GetObjectRecords(Convert.ToUInt32(0), Ent1.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, false))
                                        {
                                            if (Records1 != null)
                                            {
                                                if (Records1.Count > 0)
                                                {


                                                    int No_of_tables = 1;

                                                    foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                                    {
                                                        if (No_of_tables > 1)
                                                        {
                                                            Data_table_for_object_data.Rows.Add();
                                                            OD_id = OD_id + 1;
                                                            Data_table_for_object_data.Rows[OD_id]["OBJECT_TYPE"] = Type1;
                                                            Data_table_for_object_data.Rows[OD_id]["LAYER_NAME"] = Data_table_for_object_data.Rows[OD_id - 1]["LAYER_NAME"];
                                                            Data_table_for_object_data.Rows[OD_id]["OBJECT_ID"] = Data_table_for_object_data.Rows[OD_id - 1]["OBJECT_ID"];


                                                            if (Ent1 is BlockReference)
                                                            {
                                                                BlockReference Block1 = (BlockReference)Ent1;
                                                                Data_table_for_object_data.Rows[OD_id]["BLOCK_NAME"] = Data_table_for_object_data.Rows[OD_id - 1]["BLOCK_NAME"];
                                                            }

                                                        }
                                                        Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Record1.TableName];
                                                        Data_table_for_object_data.Rows[OD_id]["TABLE_NAME"] = Tabla1.Name;

                                                        Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;
                                                        for (int j = 0; j < Record1.Count; ++j)
                                                        {
                                                            Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[j];
                                                            string Nume_field = Field_def1.Name;
                                                            string Valoare_field = (string)Record1[j].StrValue;
                                                            if (Data_table_for_object_data.Columns.Contains(Nume_field) == false)
                                                            {
                                                                Data_table_for_object_data.Columns.Add(Nume_field, typeof(String));
                                                            }
                                                            Data_table_for_object_data.Rows[OD_id][Nume_field] = Valoare_field;
                                                        }
                                                        No_of_tables = No_of_tables + 1;
                                                    }

                                                }
                                                else
                                                {

                                                }
                                            }
                                            else
                                            {


                                            }
                                        }

                                        OD_id = OD_id + 1;

                                    }
                                    catch (AccessViolationException ex1)
                                    {
                                        MessageBox.Show(ex1.Message);
                                    }
                                }
                            }
                        }

                        Trans1.Commit();
                    }
                }
                if (Data_table_for_object_data.Rows.Count > 0)
                {

                    DataGrid1.DataSource = Data_table_for_object_data;
                    DataGrid1.Columns[0].ReadOnly = true;
                    DataGrid1.Columns[1].ReadOnly = true;
                    DataGrid1.Columns[2].ReadOnly = true;
                    DataGrid1.Columns[3].ReadOnly = true;
                    DataGrid1.Columns[4].ReadOnly = true;
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Button_Update_object_data_Click(object sender, EventArgs e)
        {
            try
            {
                if (DataGrid1.RowCount > 0)
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

                            int Rand_current = DataGrid1.CurrentCell.RowIndex;

                            ObjectId Id1 = (ObjectId)Data_table_for_object_data.Rows[Rand_current][0];
                            Entity Ent1 = (Entity)Trans1.GetObject(Id1, OpenMode.ForWrite);
                            if (Ent1 != null)
                            {


                                Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                                Autodesk.Gis.Map.ObjectData.Records Records1;

                                if (Data_table_for_object_data.Rows[Rand_current]["TABLE_NAME"] != null)
                                {

                                    string OD_table_name = (String)Data_table_for_object_data.Rows[Rand_current]["TABLE_NAME"];

                                    try
                                    {

                                        using (Records1 = Tables1.GetObjectRecords(Convert.ToUInt32(0), Ent1.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForWrite, true))
                                        {





                                            Autodesk.Gis.Map.ObjectData.Table Tabla1 = null;

                                            try
                                            {
                                                Tabla1 = Tables1[OD_table_name];
                                            }
                                            catch (Autodesk.Gis.Map.MapException ex)
                                            {

                                            }


                                            if (Tabla1 != null)
                                            {
                                                foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                                {
                                                    if (Record1.TableName == OD_table_name)
                                                    {
                                                        int Numar_max_rec = Record1.Count;
                                                        int Index_update = 0;
                                                        Autodesk.Gis.Map.Utilities.MapValue Valoare1;
                                                        for (int i = 5; i < DataGrid1.Columns.Count; i = i + 1)
                                                        {


                                                            try
                                                            {
                                                                if (Index_update < Numar_max_rec)
                                                                {
                                                                    Valoare1 = Record1[Index_update];

                                                                    if (Valoare1.Type == Autodesk.Gis.Map.Constants.DataType.Character)
                                                                    {
                                                                        if (DataGrid1.Rows[Rand_current].Cells[i].Value != null)
                                                                        {
                                                                            Valoare1.Assign(Convert.ToString(DataGrid1.Rows[Rand_current].Cells[i].Value));
                                                                            Records1.UpdateRecord(Record1);
                                                                        }
                                                                    }

                                                                    if (Valoare1.Type == Autodesk.Gis.Map.Constants.DataType.Integer)
                                                                    {
                                                                        if (Functions.IsNumeric(Convert.ToString(DataGrid1.Rows[Rand_current].Cells[i].Value)) == true)
                                                                        {
                                                                            if (DataGrid1.Rows[Rand_current].Cells[i].Value != null)
                                                                            {
                                                                                Valoare1.Assign(Convert.ToInt32(DataGrid1.Rows[Rand_current].Cells[i].Value));
                                                                                Records1.UpdateRecord(Record1);
                                                                            }
                                                                        }
                                                                        else
                                                                        {
                                                                            //string Old_value = (String)Data_table_for_object_data.Rows[Rand_current][DataGrid1.Columns[i].HeaderText];
                                                                            DataGrid1.Rows[Rand_current].Cells[i].Value = "";// Data_table_for_object_data.Rows[Rand_current][Old_value];
                                                                            DataGrid1.Rows[Rand_current].Cells[i].Selected = true;
                                                                            DataGrid1.Refresh();
                                                                        }

                                                                    }
                                                                    if (Valoare1.Type == Autodesk.Gis.Map.Constants.DataType.Real & Functions.IsNumeric(Convert.ToString(DataGrid1.Rows[Rand_current].Cells[i].Value)) == true)
                                                                    {
                                                                        if (DataGrid1.Rows[Rand_current].Cells[i].Value != null)
                                                                        {
                                                                            Valoare1.Assign(Convert.ToDouble(DataGrid1.Rows[Rand_current].Cells[i].Value));
                                                                            Records1.UpdateRecord(Record1);
                                                                        }
                                                                    }

                                                                    Index_update = Index_update + 1;
                                                                }
                                                                else
                                                                {
                                                                    Index_update = 0;
                                                                }
                                                            }
                                                            catch (System.Exception ex)
                                                            {

                                                                if (Index_update < Numar_max_rec)
                                                                {
                                                                    DataGrid1.Rows[Rand_current].Cells[i].Value = "";

                                                                    DataGrid1.Rows[Rand_current].Cells[i].Selected = true;
                                                                    Index_update = Index_update + 1;
                                                                    DataGrid1.Refresh();
                                                                }
                                                                else
                                                                {
                                                                    Index_update = 0;
                                                                }

                                                            }

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
                                else
                                {
                                    MessageBox.Show("NO OBJECT DATA LOADED FOR THE SELECTED OBJECT");
                                }

                            }


                            Trans1.Commit();
                        }
                    }






                }

                else
                {
                    MessageBox.Show("NO OBJECT DATA LOADED");
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button_add_OD_table_as_layer_name_Click(object sender, EventArgs e)
        {
            String Nume_layer123 = "123";

            try
            {
                try
                {
                    if (DataGrid1.RowCount > 0)
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



                                Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

                                System.Collections.Specialized.StringCollection Nume_tables = new System.Collections.Specialized.StringCollection();
                                
                                Nume_tables = Tables1.GetTableNames();
                                for (int k = 0; k < Nume_tables.Count; k = k + 1)
                                {
                                    //MessageBox.Show(Nume_tables[k]);
                                }

                                if (Nume_tables.Count > 0)
                                {

                                    Autodesk.Gis.Map.ObjectData.Records Records1;

                                    int Rand_current = DataGrid1.CurrentCell.RowIndex;
                                    ObjectId Id1 = (ObjectId)Data_table_for_object_data.Rows[Rand_current][0];
                                    Entity Ent1 = (Entity)Trans1.GetObject(Id1, OpenMode.ForWrite);


                                    if (Ent1 != null)
                                    {

                                        string Nume_layer = Ent1.Layer;

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
                                                    String Nume_de_adaugat = "xxxxxx";

                                                    for (int i = 0; i < Nume_tables.Count; i = i + 1)
                                                    {

                                                        string Layer1 = Ent1.Layer.ToUpper();
                                                        String Ttable1 = Nume_tables[i].ToUpper();

                                                        if (Ent1.Layer.ToUpper().Contains(Nume_tables[i].ToUpper()) == true)
                                                        {

                                                            Nume_de_adaugat = Nume_tables[i];
                                                            if (Ent1.Layer.ToUpper() == Nume_tables[i].ToUpper()) i = Nume_tables.Count;
                                                        }
                                                    }

                                                    try
                                                    {
                                                        Tabla1 = Tables1[Nume_de_adaugat];
                                                    }
                                                    catch (Autodesk.Gis.Map.MapException ex)
                                                    {
                                                        MessageBox.Show(ex.Message);
                                                    }


                                                    if (Tabla1 != null)
                                                    {
                                                        Autodesk.Gis.Map.ObjectData.Record Record1 = Autodesk.Gis.Map.ObjectData.Record.Create();
                                                        Tabla1.InitRecord(Record1);
                                                        Tabla1.AddRecord(Record1, Ent1.ObjectId);
                                                        Data_table_for_object_data.Rows[Rand_current]["TABLE_NAME"] = Tabla1.Name;
                                                        Button_Update_object_data_Click(sender, e);

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


                                Trans1.Commit();

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

        private void button_add_OD_table_as_layer_name_entire_drawing_Click(object sender, EventArgs e)
        {
            String Nume_layer123 = "123";

            try
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
                            Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                            Autodesk.AutoCAD.DatabaseServices.LayerTable layer_table = (Autodesk.AutoCAD.DatabaseServices.LayerTable)Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);



                            Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

                            System.Collections.Specialized.StringCollection Nume_tables = new System.Collections.Specialized.StringCollection();
                            Nume_tables = Tables1.GetTableNames();

                            if (Nume_tables.Count > 0)
                            {

                                Autodesk.Gis.Map.ObjectData.Records Records1;

                                foreach (ObjectId Id1 in BTrecord)
                                {


                                    Entity Ent1 = (Entity)Trans1.GetObject(Id1, OpenMode.ForWrite);


                                    if (Ent1 != null)
                                    {

                                        string Nume_layer = Ent1.Layer;

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

                                                            if (Ent1.Layer.ToUpper().Contains(Tabla1.Name.ToUpper()) == true)
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
                                                    String Nume_de_adaugat = "xxxxxx";

                                                    for (int i = 0; i < Nume_tables.Count; i = i + 1)
                                                    {
                                                        if (Ent1.Layer.ToUpper().Contains(Nume_tables[i].ToUpper()) == true)
                                                        {

                                                            Nume_de_adaugat = Nume_tables[i];


                                                            if (Ent1.Layer.ToUpper() == Nume_tables[i].ToUpper()) i = Nume_tables.Count;

                                                        }
                                                    }

                                                    try
                                                    {
                                                        Tabla1 = Tables1[Nume_de_adaugat];
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


                            Trans1.Commit();

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

        private void DataGrid1_CurrentCellChanged(object sender, EventArgs e)
        {
            if (Row_previous != Row_current)
            {
                //button_add_OD_table_as_layer_name_entire_drawing_Click(sender, e);
                Row_previous = Row_current;
            }
        }

        private void DataGrid1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (DataGrid1.CurrentCell.RowIndex != Row_current)
            {
                Row_previous = Row_current;
                Row_current = DataGrid1.CurrentCell.RowIndex;
            }
        }

        private void button_go_to_object_data_Click(object sender, EventArgs e)
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
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_object_poly;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                        Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the object:");
                        Prompt_centerline.SetRejectMessage("\nSelect a polyline!");
                        Prompt_centerline.AllowNone = true;
                        Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rezultat_object_poly = ThisDrawing.Editor.GetEntity(Prompt_centerline);

                        if (Rezultat_object_poly.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Entity Ent1 = (Entity)Trans1.GetObject(Rezultat_object_poly.ObjectId, OpenMode.ForRead);
                        ObjectId Id1 = Ent1.ObjectId;

                        if (Data_table_for_object_data != null)
                        {
                            if (Data_table_for_object_data.Rows.Count > 0)
                            {
                                for (int i = 0; i < Data_table_for_object_data.Rows.Count; i = i + 1)
                                {
                                    if (Data_table_for_object_data.Rows[i]["OBJECT_ID"] != null)
                                    {
                                        ObjectId iD2 = (ObjectId)Data_table_for_object_data.Rows[i]["OBJECT_ID"];
                                        if (Id1 == iD2)
                                        {
                                            DataGrid1.CurrentCell = DataGrid1[0, i];
                                            DataGrid1.Refresh();

                                            i = Data_table_for_object_data.Rows.Count;
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
        }

        private void button_zoom_Click(object sender, EventArgs e)
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
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);

                        ObjectId ObjId = (ObjectId)Data_table_for_object_data.Rows[DataGrid1.CurrentCell.RowIndex][0];
                        Entity Ent1 = (Entity)Trans1.GetObject(ObjId, OpenMode.ForRead);
                        if (Ent1 != null)
                        {
                            Autodesk.AutoCAD.GraphicsSystem.KernelDescriptor kd = new Autodesk.AutoCAD.GraphicsSystem.KernelDescriptor();

                            kd.addRequirement(Autodesk.AutoCAD.UniqueString.Intern("3D Drawing"));   



                            int Cvport = Convert.ToInt32(Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("CVPORT"));

                            Point3d minx = Ent1.GeometricExtents.MinPoint;

                            Point3d maxx = Ent1.GeometricExtents.MaxPoint;

                            Autodesk.AutoCAD.GraphicsSystem.Manager GraphicsManager = ThisDrawing.GraphicsManager;

                            using (GraphicsManager)
                            {
                                Autodesk.AutoCAD.GraphicsSystem.View view = GraphicsManager.ObtainAcGsView(Cvport, kd);

                                if (view != null)
                                {
                                    using (view)
                                    {
                                        view.ZoomExtents(Ent1.GeometricExtents.MaxPoint, Ent1.GeometricExtents.MinPoint);

                                        view.Zoom(0.95);//<--optional 

                                        GraphicsManager.SetViewportFromView(Cvport, view, true, true, false);

                                    }


                                }


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



        }






    }
}
