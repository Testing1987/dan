using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

namespace Alignment_mdi
{
    public partial class OD_layer_change_form : Form
    {

        string[] Colectie_ID;

        string Layer_name = "";
        bool Freeze_operations = false;

        System.Data.DataTable Data_table1 = new System.Data.DataTable();

        public OD_layer_change_form()
        {
            InitializeComponent();
        }

        private void Button_read_OD_Click(object sender, EventArgs e)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            if (Freeze_operations == false)
            {
                Freeze_operations = true;
                try
                {
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();


                    Layer_name = "";

                    Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat1;
                    Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt1;
                    Prompt1 = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect an object containing object data:");
                    Prompt1.SetRejectMessage("\nSelect a polyline!");
                    Prompt1.AllowNone = true;
                    Prompt1.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                    Rezultat1 = ThisDrawing.Editor.GetEntity(Prompt1);

                    if (Rezultat1.Status != PromptStatus.OK)
                    {

                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                        return;
                    }

                    using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);


                            DBObject DBobj1 = Trans1.GetObject(Rezultat1.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                            Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                            ObjectId Id1 = DBobj1.ObjectId;
                            Entity Ent1 = (Entity)DBobj1;

                            using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tables1.GetObjectRecords(Convert.ToUInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, false))
                            {
                                if (Records1 != null)
                                {

                                    if (Records1.Count > 0)
                                    {
                                        if (comboBox_OD1.Items.Count > 0)
                                        {
                                            comboBox_OD1.Items.Clear();
                                        }

                                        if (comboBox_OD2.Items.Count > 0)
                                        {
                                            comboBox_OD2.Items.Clear();
                                        }

                                        Layer_name = Ent1.Layer;
                                        foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                        {
                                            Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Record1.TableName];
                                            Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;
                                            for (int i = 0; i < Record1.Count; ++i)
                                            {
                                                Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[i];
                                                string Nume_field = Field_def1.Name;
                                                //string Valoare_field = Record1[i].StrValue;

                                                if (comboBox_OD1.Items.Contains(Nume_field) == false)
                                                {
                                                    comboBox_OD1.Items.Add(Nume_field);
                                                }

                                                if (comboBox_OD2.Items.Contains(Nume_field) == false)
                                                {
                                                    comboBox_OD2.Items.Add(Nume_field);
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
                catch (System.Exception ex)
                {
                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\n" + "Command:");
                    MessageBox.Show(ex.Message);

                }
                finally
                {

                }
                Freeze_operations = false;

            }

        }

        private void Button_read_Excel_Click(object sender, EventArgs e)
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

            Microsoft.Office.Interop.Excel.Worksheet W1;
            W1 = Functions.Get_active_worksheet_from_Excel();
            if (W1 != null)
            {
                int idx1 = 1;
                for (int i = Start1; i <= End1; ++i)
                {
                    if (W1.Range[ColXL + Convert.ToString(i)].Text != "")
                    {
                        Array.Resize(ref Colectie_ID, idx1);
                        Colectie_ID[idx1 - 1] = Convert.ToString(W1.Range[ColXL + Convert.ToString(i)].Value2);
                        idx1 = idx1 + 1;
                    }
                }

                var Functie = new Functions();
                Functie.Incarca_existing_layers_to_combobox(comboBox_Layers);
                if (comboBox_Layers.Items.Count > 0) comboBox_Layers.SelectedIndex = 0;
            }
        }

        private void Form_Click(object sender, EventArgs e)
        {
            var Functie = new Functions();
            Functie.Incarca_existing_layers_to_combobox(comboBox_Layers);
            if (comboBox_Layers.Items.Count > 0) comboBox_Layers.SelectedIndex = 0;
        }

        private void Change_layer_Click(object sender, EventArgs e)
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

                        if (checkBox_use_null_values.Checked == false)
                        {
                            if (Colectie_ID != null)
                            {
                                if (Colectie_ID.Length > 0)
                                {
                                    if (comboBox_OD1.Text != "" & comboBox_Layers.Text != "")
                                    {

                                        string Layer_name = comboBox_Layers.Text;
                                        string OD_field = comboBox_OD1.Text;


                                        Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                                        Autodesk.Gis.Map.ObjectData.Records Records1;

                                        foreach (ObjectId id1 in BTrecord)
                                        {

                                            Entity Ent1 = (Entity)Trans1.GetObject(id1, OpenMode.ForRead);
                                            if (Ent1 != null)
                                            {

                                                try
                                                {
                                                    using (Records1 = Tables1.GetObjectRecords(Convert.ToUInt32(0), Ent1.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, false))
                                                    {
                                                        if (Records1 != null)
                                                        {
                                                            if (Records1.Count > 0)
                                                            {
                                                                foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                                                {
                                                                    Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Record1.TableName];
                                                                    Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;
                                                                    for (int j = 0; j < Record1.Count; ++j)
                                                                    {
                                                                        Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[j];
                                                                        string Nume_field = Field_def1.Name;
                                                                        if (Nume_field == OD_field)
                                                                        {
                                                                            string Valoare_field = (string)Record1[j].StrValue;
                                                                            foreach (String String1 in Colectie_ID)
                                                                            {
                                                                                if (String1 == Valoare_field)
                                                                                {

                                                                                    if (Ent1 != null)
                                                                                    {
                                                                                        Ent1.UpgradeOpen();
                                                                                        Ent1.Layer = Layer_name;
                                                                                    }
                                                                                }
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
                                        }

                                    }
                                }
                            }
                        }
                        else
                        {

                            string Layer_name = comboBox_Layers.Text;
                            string OD_field = comboBox_OD1.Text;

                            Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                            Autodesk.Gis.Map.ObjectData.Records Records1;

                            foreach (ObjectId id1 in BTrecord)
                            {
                                Entity Ent1 = (Entity)Trans1.GetObject(id1, OpenMode.ForRead);
                                if (Ent1 != null)
                                {
                                    try
                                    {
                                        using (Records1 = Tables1.GetObjectRecords(Convert.ToUInt32(0), Ent1.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, false))
                                        {
                                            if (Records1 != null)
                                            {
                                                if (Records1.Count > 0)
                                                {
                                                    foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                                    {
                                                        Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Record1.TableName];
                                                        Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;
                                                        for (int j = 0; j < Record1.Count; ++j)
                                                        {
                                                            Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[j];
                                                            string Nume_field = Field_def1.Name;
                                                            if (Nume_field == OD_field)
                                                            {
                                                                string Valoare_field = (string)Record1[j].StrValue;

                                                                if (Valoare_field.Replace(" ", "") == "")
                                                                {
                                                                    if (Ent1 != null)
                                                                    {
                                                                        Ent1.UpgradeOpen();
                                                                        Ent1.Layer = Layer_name;
                                                                    }
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

        private void button_load_OD_Click(object sender, EventArgs e)
        {

            if (Freeze_operations == false)
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

                Freeze_operations = true;
                try
                {

                    if (comboBox_OD1.Text != "" & comboBox_OD2.Text != "")
                    {


                        Data_table1 = new System.Data.DataTable();
                        Data_table1.Columns.Add("OD1", typeof(String));
                        Data_table1.Columns.Add("OD2", typeof(String));





                        Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;

                        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                        using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                        {
                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                            {
                                Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                                Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;



                                int Id_od = 0;

                                foreach (ObjectId Obj_ID1 in BTrecord)
                                {
                                    Entity Ent1 = (Entity)Trans1.GetObject(Obj_ID1, OpenMode.ForRead);
                                    if (Ent1 != null)
                                    {
                                        Boolean Add_to_table = false;


                                        if (Ent1.Layer == Layer_name)
                                        {
                                            Add_to_table = true;
                                        }



                                        if (Add_to_table == true)
                                        {

                                            using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tables1.GetObjectRecords(Convert.ToUInt32(0), Ent1.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, false))
                                            {
                                                if (Records1 != null)
                                                {
                                                    if (Records1.Count > 0)
                                                    {


                                                        Data_table1.Rows.Add();
                                                        Data_table1.Rows[Id_od]["OD1"] = Ent1.ObjectId;
                                                        Data_table1.Rows[Id_od]["OD2"] = 0;


                                                        foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                                        {
                                                            Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Record1.TableName];
                                                            Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;

                                                            for (int i = 0; i < Record1.Count; ++i)
                                                            {
                                                                Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[i];
                                                                string Nume_field = Field_def1.Name;
                                                                String Valoare1 = Record1[i].StrValue;
                                                                if (Nume_field == comboBox_OD1.Text) Data_table1.Rows[Id_od]["OD1"] = Valoare1;
                                                                if (Nume_field == comboBox_OD2.Text) Data_table1.Rows[Id_od]["OD2"] = Valoare1;

                                                            }
                                                        }

                                                        Id_od = Id_od + 1;
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
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                MessageBox.Show("OD loaded");
                Freeze_operations = false;

            }


        }

        private void button_blocks_Click(object sender, EventArgs e)
        {
            if (Freeze_operations == false)
            {
                Freeze_operations = true;
                Incarca_existing_blocks_with_attributes_to_combobox(comboBox_blocks);
                Freeze_operations = false;
            }


        }



        static public void Incarca_existing_blocks_with_attributes_to_combobox(System.Windows.Forms.ComboBox Combo_blocks)
        {

            {
                try
                {
                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument();
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);

                        Combo_blocks.Items.Clear();


                        foreach (ObjectId Block_id in BlockTable1)
                        {
                            BlockTableRecord Block1 = (BlockTableRecord)Trans1.GetObject(Block_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);

                            if (Block1.HasAttributeDefinitions == true & Block1.Name.Contains("*") == false)
                            {
                                Combo_blocks.Items.Add(Block1.Name);
                            }

                        }

                        Trans1.Dispose();
                    }
                }
                catch (System.Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);
                }
            }

        }




        private void comboBox_blocks_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Freeze_operations == false)
            {
                Freeze_operations = true;
                Incarca_existing_Atributes_to_combobox(comboBox_blocks.Text, comboBox_block_atr1);
                Freeze_operations = false;
            }
        }



        static public void Incarca_existing_Atributes_to_combobox(string Blockname, System.Windows.Forms.ComboBox Combo_attrib)
        {

            {
                try
                {
                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument();
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);

                        Combo_attrib.Items.Clear();
                        if (Blockname != "")
                        {
                            BlockTableRecord Block1 = (BlockTableRecord)Trans1.GetObject(BlockTable1[Blockname], Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                            foreach (ObjectId Attrib_id in Block1)
                            {
                                Entity Enta1 = (Entity)Trans1.GetObject(Attrib_id, OpenMode.ForRead);

                                AttributeDefinition Atrib1 = null;

                                if (Enta1 is AttributeDefinition)
                                {
                                    Atrib1 = (AttributeDefinition)Trans1.GetObject(Attrib_id, OpenMode.ForRead);
                                }
                                if (Atrib1 != null)
                                {
                                    Combo_attrib.Items.Add(Atrib1.Tag);
                                }

                            }
                        }

                        Trans1.Dispose();
                    }
                }
                catch (System.Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);
                }
            }

        }

        private void button_change_attribute_value_Click(object sender, EventArgs e)
        {
            if (Freeze_operations == false)
            {
                Freeze_operations = true;
                try
                {
                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

                    if (comboBox_OD1.Text != "" & comboBox_OD2.Text != "" & comboBox_blocks.Text != "" & comboBox_block_atr1.Text != "")
                    {

                        if (Data_table1 != null)
                        {
                            if (Data_table1.Rows.Count > 0)
                            {
                                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;

                                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                                {
                                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                                    {
                                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);





                                        foreach (ObjectId Obj_ID1 in BTrecord)
                                        {
                                            Entity Ent1 = (Entity)Trans1.GetObject(Obj_ID1, OpenMode.ForRead);
                                            if (Ent1 != null)
                                            {

                                                if (Ent1 is BlockReference)
                                                {
                                                    BlockReference Block1 = (BlockReference)Ent1;
                                                    String Nume1 = "";


                                                    BlockTableRecord BlockTrec = null;
                                                    if (Block1.IsDynamicBlock == true)
                                                    {
                                                        BlockTrec = (BlockTableRecord)Trans1.GetObject(Block1.DynamicBlockTableRecord, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                                                        Nume1 = BlockTrec.Name;
                                                    }
                                                    else
                                                    {
                                                        BlockTrec = (BlockTableRecord)Trans1.GetObject(Block1.BlockTableRecord, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                                                        Nume1 = BlockTrec.Name;
                                                    }
                                                    if (Nume1 == comboBox_blocks.Text)
                                                    {
                                                        Block1.UpgradeOpen();
                                                        Autodesk.AutoCAD.DatabaseServices.AttributeCollection Col1 = Block1.AttributeCollection;


                                                        foreach (ObjectId IdA1 in Col1)
                                                        {

                                                            AttributeReference Atrib1 = (AttributeReference)Trans1.GetObject(IdA1, OpenMode.ForWrite);
                                                            String Valoare1 = "";
                                                            if (Atrib1.Tag == comboBox_block_atr1.Text)
                                                            {
                                                                if (Atrib1.IsMTextAttribute == true)
                                                                {
                                                                    Valoare1 = Atrib1.MTextAttribute.Text;
                                                                }
                                                                else
                                                                {
                                                                    Valoare1 = Atrib1.TextString;
                                                                }

                                                                for (int i = 0; i < Data_table1.Rows.Count; ++i)
                                                                {
                                                                    if (Data_table1.Rows[i][0] != DBNull.Value)
                                                                    {
                                                                        string OD1 = (string)Data_table1.Rows[i][0];
                                                                        if (OD1 == Valoare1)
                                                                        {

                                                                            string OD2 = "";
                                                                            if (Data_table1.Rows[i][0] != DBNull.Value)
                                                                            {
                                                                                OD2 = (string)Data_table1.Rows[i][1];
                                                                            }

                                                                            Atrib1.TextString = OD2;

                                                                            i = Data_table1.Rows.Count;

                                                                        }


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
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                MessageBox.Show("blocks UPDATED");





                Freeze_operations = false;
            }
        }


    }
}
