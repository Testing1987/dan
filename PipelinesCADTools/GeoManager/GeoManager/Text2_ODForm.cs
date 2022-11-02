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


namespace MMGeoTools
{
    public partial class Text2_ODForm : Form
    {
        bool Freeze_operations = false;


        public Text2_ODForm()
        {
            InitializeComponent();
        }

        private void button_load_OD_Click(object sender, EventArgs e)
        {

            if (Freeze_operations == false)
            {
                Freeze_operations = true;
                try
                {
                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();


                    using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            comboBox_OD_table.Items.Clear();
                            comboBox_OD_field.Items.Clear();

                            Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                            System.Collections.Specialized.StringCollection Nume_tables = new System.Collections.Specialized.StringCollection();
                            Nume_tables = Tables1.GetTableNames();

                            for (int i = 0; i < Nume_tables.Count; i = i + 1)
                            {
                                String Tabla1 = Nume_tables[i];
                                if (comboBox_OD_table.Items.Contains(Tabla1) == false)
                                {
                                    comboBox_OD_table.Items.Add(Tabla1);
                                }
                            }
                            List<string> List2 = new List<string>();

                            for (int i = 0; i < comboBox_OD_table.Items.Count; ++i)
                            {

                                string item2 = comboBox_OD_table.Items[i].ToString();
                                List2.Add(item2);

                            }

                            this.Refresh();
                        }
                    }

                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                Freeze_operations = false;
            }

        }

        private void button_tranfer_text_Click(object sender, EventArgs e)
        {
            if (Freeze_operations == false)
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

                if (comboBox_OD_field.Text == "" | comboBox_OD_table.Text == "")
                {
                    MessageBox.Show("specify the field and table name");

                    return;
                }


                Freeze_operations = true;
                try
                {

                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                    using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                            Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_text;
                            Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_text;
                            Prompt_text = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the text object:");
                            Prompt_text.SetRejectMessage("\nSelect a text ar mtext!");
                            Prompt_text.AllowNone = true;
                            Prompt_text.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.DBText), false);
                            Prompt_text.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.MText), false);
                            Rezultat_text = ThisDrawing.Editor.GetEntity(Prompt_text);

                            if (Rezultat_text.Status != PromptStatus.OK)
                            {

                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }


                            Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_destination;
                            Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_destination;
                            Prompt_destination = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the destination:");
                            Prompt_destination.SetRejectMessage("\nSelect a entity!");
                            Prompt_destination.AllowNone = true;
                            Prompt_destination.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Entity), false);
                            Rezultat_destination = ThisDrawing.Editor.GetEntity(Prompt_destination);

                            if (Rezultat_destination.Status != PromptStatus.OK)
                            {

                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }




                            ObjectId Id1 = (ObjectId)Rezultat_text.ObjectId;

                            Entity Ent1 = null;
                            try
                            {
                                Ent1 = (Entity)Trans1.GetObject(Id1, OpenMode.ForRead);
                            }
                            catch (System.Exception ex)
                            {

                                MessageBox.Show("The object to be updated was deleted" + "\r\nRefresh!");
                                Freeze_operations = false;
                                return;
                            }


                            ObjectId Id2 = (ObjectId)Rezultat_destination.ObjectId;

                            Entity Ent2 = null;
                            try
                            {
                                Ent2 = (Entity)Trans1.GetObject(Id2, OpenMode.ForWrite);
                            }
                            catch (System.Exception ex)
                            {

                                MessageBox.Show("The object to be updated was deleted" + "\r\nRefresh!");
                                Freeze_operations = false;
                                return;
                            }

                            if (Ent1 != null && Ent2 != null)
                            {


                                string textstring1 = "";
                                if (Ent1 is DBText)
                                {
                                    DBText text1 = (DBText)Ent1;
                                    textstring1 = text1.TextString;
                                }
                                if (Ent1 is MText)
                                {
                                    MText text1 = (MText)Ent1;
                                    textstring1 = text1.Text;
                                }

                                Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                                Autodesk.Gis.Map.ObjectData.Records Records1;
                                if (Tables1.IsTableDefined(comboBox_OD_table.Text) == true)
                                {
                                    Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[comboBox_OD_table.Text];
                                    Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;
                                    for (int i = 0; i < Field_defs1.Count; ++i)
                                    {
                                        Autodesk.Gis.Map.ObjectData.FieldDefinition fielddef1 = Field_defs1[i];
                                        if (fielddef1.Name == comboBox_OD_field.Text)
                                        {
                                            using (Records1 = Tabla1.GetObjectTableRecords(Convert.ToUInt32(0), Ent2.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForWrite, true))
                                            {
                                                if (Records1.Count == 0)
                                                {

                                                    using (Autodesk.Gis.Map.ObjectData.Record Record1 = Autodesk.Gis.Map.ObjectData.Record.Create())
                                                    {
                                                        Tabla1.InitRecord(Record1);
                                                        Autodesk.Gis.Map.Utilities.MapValue Valoare1 = Record1[i];
                                                        
                                                        if (Valoare1.Type == Autodesk.Gis.Map.Constants.DataType.Character)
                                                        {
                                                            Valoare1.Assign(textstring1);
                                                        }
                                                        if (Valoare1.Type == Autodesk.Gis.Map.Constants.DataType.Integer)
                                                        {
                                                            if (Alignment_mdi.Functions.IsNumeric(textstring1) == true)
                                                            {
                                                                Valoare1.Assign(Convert.ToInt32(textstring1));
                                                            }
                                                        }
                                                        if (Valoare1.Type == Autodesk.Gis.Map.Constants.DataType.Real)
                                                        {
                                                            if (Alignment_mdi.Functions.IsNumeric(textstring1) == true)
                                                            {
                                                                Valoare1.Assign(Convert.ToDouble(textstring1));
                                                            }
                                                        }
                                                        Tabla1.AddRecord(Record1, Ent2.ObjectId);

                                                    }

                                                }
                                                else
                                                {
                                                    foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                                    {
                                                        Autodesk.Gis.Map.Utilities.MapValue Valoare1 = Record1[i];



                                                        if (Valoare1.Type == Autodesk.Gis.Map.Constants.DataType.Character)
                                                        {
                                                            Valoare1.Assign(textstring1);
                                                        }
                                                        if (Valoare1.Type == Autodesk.Gis.Map.Constants.DataType.Integer)
                                                        {
                                                            if (Alignment_mdi.Functions.IsNumeric(textstring1) == true)
                                                            {
                                                                Valoare1.Assign(Convert.ToInt32(textstring1));
                                                            }
                                                        }
                                                        if (Valoare1.Type == Autodesk.Gis.Map.Constants.DataType.Real)
                                                        {
                                                            if (Alignment_mdi.Functions.IsNumeric(textstring1) == true)
                                                            {
                                                                Valoare1.Assign(Convert.ToDouble(textstring1));
                                                            }
                                                        }



                                                        Records1.UpdateRecord(Record1);

                                                    }
                                                }



                                            }



                                        }
                                    }



                                }
                                else
                                {

                                    MessageBox.Show("The table not found");
                                    Freeze_operations = false;
                                    return;
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
                Freeze_operations = false;
            }
        }

        private void comboBox_OD_table_SelectedIndexChanged(object sender, EventArgs e)
        {
            Alignment_mdi.Functions.add_OD_fieds_to_combobox(comboBox_OD_table, comboBox_OD_field);
        }

    }
}
