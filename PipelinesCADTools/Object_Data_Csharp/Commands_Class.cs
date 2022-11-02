using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System.Management;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;


using Autodesk.AutoCAD.EditorInput;



namespace Alignment_mdi
{
    public class Commands_Class
    {
        public bool isSECURE()
        {

            string number_drive = GetHDDSerialNumber("C");

            switch (number_drive)
            {
                case "8CDA6CE3":
                    return true;
                case "36D79DE5":
                    return true;
                case "FEA3192C":
                    return true;
                case "B454BD5B":
                    return true;
                case "6E40460D":
                    return true;
                case "0892E01D":
                    return true;
                case "4ED21ABF":
                    return true;
                case "56766C69":
                    return true;
                case "DA214366":
                    return true;
                case "3CF68AF2":
                    return true;
                case "389A2249":
                    return true;
                case "AED6B68E":
                    return true;
                case "8C040338":
                    return true;
                case "8CD08F48":
                    return true;
                case "0E26E402":
                    return true;
                case "4A123A50":
                    return true;

                case "98D9B617":
                    return true;
                case "B838FEB4":
                    return true;
                case "1AE1721C":
                    return true;
                case "CA9E6FFE":
                    return true;
                case "DE281128":
                    return true;
                case "FC7C4F1":
                    return true;
                case "B67EC134":
                    return true;
                case "E64DBF0A":
                    return true;
                case "561F1509":
                    return true;
                case "B63AD3F6":
                    return true;

                case "120E4B54":
                    return true;
                case "F6633173":
                    return true;
                case "40D6BDCB":
                    return true;
                case "18399D24":
                    return true;

                default:
                    try
                    {
                        string UserDNS = Environment.GetEnvironmentVariable("USERDNSDOMAIN");
                        if (UserDNS == "HMMG.CC" | UserDNS.ToLower() == "mottmac.group.int")
                        {
                            return true;
                        }
                        else
                        {
                            return false;
                        }
                    }
                    catch (System.Exception ex)
                    {
                        return false;
                    }
            }
        }


        public string GetHDDSerialNumber(string drive)
        {
            //check to see if the user provided a drive letter
            //if not default it to "C"
            if (drive == "" || drive == null)
            {
                drive = "C";
            }
            //create our ManagementObject, passing it the drive letter to the
            //DevideID using WQL
            ManagementObject disk = new ManagementObject("Win32_LogicalDisk.DeviceID=\"" + drive + ":\"");
            //bind our management object
            disk.Get();
            //return the serial number
            return disk["VolumeSerialNumber"].ToString();
        }

        [CommandMethod("OD_QUERY")]
        public void Show_od_layer_form()
        {
            if (isSECURE() == true)
            {
                foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
                {
                    if (Forma1 is OD_layer_change_form)
                    {
                        Forma1.Focus();
                        Forma1.WindowState = System.Windows.Forms.FormWindowState.Normal;
                        return;
                    }
                }

                try
                {
                    OD_layer_change_form forma2 = new OD_layer_change_form();
                    Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma2);
                }
                catch (System.Exception EX)
                {
                    MessageBox.Show(EX.Message);
                }


            }
            else
            {
                return;
            }
        }

        [CommandMethod("PPL_WOD_XL")]
        public void write_OD_to_Excel_and_highlight_OBJ()
        {
            if (isSECURE() == true)
            {
                try
                {

                    System.Data.DataTable Data_table_for_object_data = new System.Data.DataTable();

                    Data_table_for_object_data.Columns.Add("LAYER", typeof(String));
                    Data_table_for_object_data.Columns.Add("X", typeof(double));
                    Data_table_for_object_data.Columns.Add("Y", typeof(double));
                    Data_table_for_object_data.Columns.Add("XY", typeof(string));
                    Data_table_for_object_data.Columns.Add("OBJECT_TYPE", typeof(String));
                    Data_table_for_object_data.Columns.Add("CLOSED", typeof(String));

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
                            int Color_index = -1;
                            int OD_id = 0;
                            Autodesk.AutoCAD.EditorInput.PromptIntegerOptions Line_prompt00 = new Autodesk.AutoCAD.EditorInput.PromptIntegerOptions(
                                                           "\n" + "Specify color index for objects with no object data attached (if you press escape no object will be changed)|0=ByBlock|.......|256=ByLayer|:");
                            Line_prompt00.AllowNegative = false;
                            Line_prompt00.AllowZero = true;
                            Line_prompt00.AllowNone = true;
                            Autodesk.AutoCAD.EditorInput.PromptIntegerResult Rezultat00 = Editor1.GetInteger(Line_prompt00);

                            if (Rezultat00.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                            {
                                Color_index = Rezultat00.Value;
                            }

                            if (Color_index > 256) Color_index = -1;
                            if (Color_index < 0) Color_index = -1;

                            foreach (ObjectId Obj_ID1 in BTrecord)
                            {
                                Entity Ent1 = (Entity)Trans1.GetObject(Obj_ID1, OpenMode.ForRead);
                                if (Ent1 != null)
                                {
                                    if (Ent1 is Curve | Ent1 is DBPoint)
                                    {

                                        try
                                        {
                                            using (Records1 = Tables1.GetObjectRecords(Convert.ToUInt32(0), Ent1.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, false))
                                            {
                                                if (Records1 != null)
                                                {
                                                    if (Records1.Count > 0)
                                                    {
                                                        Data_table_for_object_data.Rows.Add();

                                                        Data_table_for_object_data.Rows[OD_id]["LAYER"] = Ent1.Layer;

                                                        if (Ent1 is Polyline)
                                                        {
                                                            Polyline Poly1 = (Polyline)Ent1;
                                                            Data_table_for_object_data.Rows[OD_id]["X"] = Poly1.StartPoint.X;
                                                            Data_table_for_object_data.Rows[OD_id]["Y"] = Poly1.StartPoint.Y;
                                                            Data_table_for_object_data.Rows[OD_id]["XY"] = Poly1.StartPoint.X + "," + Poly1.StartPoint.Y;
                                                            Data_table_for_object_data.Rows[OD_id]["OBJECT_TYPE"] = "Polyline";
                                                            if (Poly1.Closed == true)
                                                            {
                                                                Data_table_for_object_data.Rows[OD_id]["CLOSED"] = "CLOSED";
                                                            }
                                                            else
                                                            {
                                                                Data_table_for_object_data.Rows[OD_id]["CLOSED"] = "OPEN";
                                                            }
                                                        }

                                                        if (Ent1 is Autodesk.AutoCAD.DatabaseServices.Line)
                                                        {
                                                            Autodesk.AutoCAD.DatabaseServices.Line Line1 = (Autodesk.AutoCAD.DatabaseServices.Line)Ent1;
                                                            Data_table_for_object_data.Rows[OD_id]["X"] = Line1.StartPoint.X;
                                                            Data_table_for_object_data.Rows[OD_id]["Y"] = Line1.StartPoint.Y;
                                                            Data_table_for_object_data.Rows[OD_id]["XY"] = Line1.StartPoint.X + "," + Line1.StartPoint.Y;
                                                            Data_table_for_object_data.Rows[OD_id]["OBJECT_TYPE"] = "Line";
                                                            Data_table_for_object_data.Rows[OD_id]["CLOSED"] = "n/a";
                                                        }

                                                        if (Ent1 is Autodesk.AutoCAD.DatabaseServices.Circle)
                                                        {
                                                            Autodesk.AutoCAD.DatabaseServices.Circle Circle1 = (Autodesk.AutoCAD.DatabaseServices.Circle)Ent1;
                                                            Data_table_for_object_data.Rows[OD_id]["X"] = Circle1.Center.X;
                                                            Data_table_for_object_data.Rows[OD_id]["Y"] = Circle1.Center.Y;
                                                            Data_table_for_object_data.Rows[OD_id]["XY"] = Circle1.Center.X + "," + Circle1.Center.Y;
                                                            Data_table_for_object_data.Rows[OD_id]["OBJECT_TYPE"] = "Circle";
                                                            Data_table_for_object_data.Rows[OD_id]["CLOSED"] = "n/a";
                                                        }


                                                        if (Ent1 is Polyline3d)
                                                        {
                                                            Polyline3d Poly1 = (Polyline3d)Ent1;
                                                            Data_table_for_object_data.Rows[OD_id]["X"] = Poly1.StartPoint.X;
                                                            Data_table_for_object_data.Rows[OD_id]["Y"] = Poly1.StartPoint.Y;
                                                            Data_table_for_object_data.Rows[OD_id]["XY"] = Poly1.StartPoint.X + "," + Poly1.StartPoint.Y;
                                                            Data_table_for_object_data.Rows[OD_id]["OBJECT_TYPE"] = "3D Polyline";
                                                            Data_table_for_object_data.Rows[OD_id]["CLOSED"] = "n/a";
                                                        }


                                                        if (Ent1 is DBPoint)
                                                        {
                                                            DBPoint Point1 = (DBPoint)Ent1;
                                                            Data_table_for_object_data.Rows[OD_id]["X"] = Point1.Position.X;
                                                            Data_table_for_object_data.Rows[OD_id]["Y"] = Point1.Position.Y;
                                                            Data_table_for_object_data.Rows[OD_id]["XY"] = Point1.Position.X + "," + Point1.Position.Y;
                                                            Data_table_for_object_data.Rows[OD_id]["OBJECT_TYPE"] = "Point";
                                                            Data_table_for_object_data.Rows[OD_id]["CLOSED"] = "n/a";
                                                        }

                                                        foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                                        {
                                                            Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Record1.TableName];
                                                            Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;
                                                            for (int j = 0; j < Record1.Count; ++j)
                                                            {
                                                                Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[j];
                                                                string Nume_field = Field_def1.Name;
                                                                string Valoare_field = (string)Record1[j].StrValue;
                                                                if (Data_table_for_object_data.Columns.Contains("[" + Tabla1.Name + "]" + "-" + Nume_field) == false)
                                                                {
                                                                    Data_table_for_object_data.Columns.Add("[" + Tabla1.Name + "]" + "-" + Nume_field, typeof(String));
                                                                }
                                                                Data_table_for_object_data.Rows[OD_id]["[" + Tabla1.Name + "]" + "-" + Nume_field] = Valoare_field;
                                                            }
                                                        }
                                                        OD_id = OD_id + 1;
                                                    }
                                                    else
                                                    {
                                                        if (Color_index != -1)
                                                        {
                                                            Ent1.UpgradeOpen();
                                                            Ent1.ColorIndex = Color_index;
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    if (Color_index != -1)
                                                    {
                                                        Ent1.UpgradeOpen();
                                                        Ent1.ColorIndex = Color_index;
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
                    if (Data_table_for_object_data.Rows.Count > 0)
                    {
                        Worksheet W1 = Alignment_mdi.Functions.Get_NEW_worksheet_from_Excel();
                        for (int i = 0; i < Data_table_for_object_data.Columns.Count; i = i + 1)
                        {
                            W1.Cells[1, i + 1].Value2 = Data_table_for_object_data.Columns[i].ColumnName;

                        }
                        for (int i = 0; i < Data_table_for_object_data.Rows.Count; i = i + 1)
                        {

                            for (int j = 0; j < Data_table_for_object_data.Columns.Count; j = j + 1)
                            {
                                if (Data_table_for_object_data.Rows[i][j] != DBNull.Value)
                                {
                                    W1.Cells[i + 2, j + 1].Value2 = Data_table_for_object_data.Rows[i][j];
                                }
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                MessageBox.Show("Done");
            }
            else
            {
                return;
            }
        }

        [CommandMethod("PPL_MATCH_OD")]
        public void MATCH_OD_FROM_ONE_OBJ_TO_ANOTHER()
        {
            if (isSECURE() == true)
            {
                try
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

                            Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez1 = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rez1.MessageForAdding = "\nSelect the source_object:";
                            Prompt_rez1.SingleOnly = true;
                            Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez1);

                            if (Rezultat1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                            {

                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }
                            Entity Ent1 = (Entity)Trans1.GetObject(Rezultat1.Value[0].ObjectId, OpenMode.ForRead);

                            Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat2;
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez2 = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rez2.MessageForAdding = "\nSelect the destination object(s):";
                            Prompt_rez2.SingleOnly = false;
                            Rezultat2 = ThisDrawing.Editor.GetSelection(Prompt_rez2);

                            if (Rezultat2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                            {

                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }

                            for (int i = 0; i < Rezultat2.Value.Count; i = i + 1)
                            {


                                Entity Ent2 = (Entity)Trans1.GetObject(Rezultat2.Value[i].ObjectId, OpenMode.ForRead);



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
                                                else
                                                {
                                                    MessageBox.Show("NO OBJECT DATA ATTACHED TO THE SOURCE OBJECT");
                                                }
                                            }
                                            else
                                            {
                                                MessageBox.Show("NO OBJECT DATA ATTACHED TO THE SOURCE OBJECT");
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
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                return;
            }
        }


        [CommandMethod("PPL_CONC_OWNER1")]
        public void concatenate_2_fields_in_the_3_one()
        {
            if (isSECURE() == true)
            {
                try
                {

                    String OD_field1 = "OwnerL";
                    String OD_field2 = "OwnerF";
                    String OD_field3 = "Owner";


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



                            foreach (ObjectId Obj_ID1 in BTrecord)
                            {
                                Entity Ent1 = (Entity)Trans1.GetObject(Obj_ID1, OpenMode.ForWrite);
                                if (Ent1 != null)
                                {


                                    using (Records1 = Tables1.GetObjectRecords(Convert.ToUInt32(0), Ent1.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForWrite, false))
                                    {
                                        if (Records1 != null)
                                        {
                                            if (Records1.Count > 0)
                                            {

                                                String Valoare_de_adaugat = "";
                                                String Valoare_de_adaugat1 = "";
                                                String Valoare_de_adaugat2 = "";

                                                Boolean ADD = false;
                                                foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                                {
                                                    Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Record1.TableName];
                                                    Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;


                                                    for (int j = 0; j < Record1.Count; ++j)
                                                    {
                                                        Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[j];
                                                        string Nume_field = Field_def1.Name;
                                                        string Valoare_field = (string)Record1[j].StrValue;
                                                        if (Nume_field == OD_field1) Valoare_de_adaugat1 = Valoare_field;
                                                        if (Nume_field == OD_field2) Valoare_de_adaugat2 = Valoare_field;
                                                        if (Nume_field == OD_field3) ADD = true;
                                                    }

                                                }

                                                if (Valoare_de_adaugat1 != "")
                                                {
                                                    Valoare_de_adaugat = Valoare_de_adaugat1.ToUpper();
                                                }

                                                if (Valoare_de_adaugat2 != "")
                                                {
                                                    if (Valoare_de_adaugat1 != "")
                                                    {
                                                        Valoare_de_adaugat = Valoare_de_adaugat1.ToUpper() + ", " + Valoare_de_adaugat2.ToUpper();
                                                    }
                                                    else
                                                    {
                                                        Valoare_de_adaugat = Valoare_de_adaugat2.ToUpper();
                                                    }
                                                }

                                                if (ADD == true & Valoare_de_adaugat != "")
                                                {



                                                    foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                                    {
                                                        Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Record1.TableName];
                                                        Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;





                                                        Autodesk.Gis.Map.Utilities.MapValue MapVal;


                                                        for (int j = 0; j < Record1.Count; ++j)
                                                        {
                                                            Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[j];
                                                            string Nume_field = Field_def1.Name;
                                                            string Valoare_field = (string)Record1[j].StrValue;
                                                            MapVal = Record1[j];


                                                            if (Nume_field == OD_field3 & Valoare_field == "")
                                                            {
                                                                MapVal.Assign(Valoare_de_adaugat);
                                                            }
                                                            else
                                                            {
                                                                MapVal.Assign(Valoare_field);
                                                            }


                                                        }

                                                        Records1.UpdateRecord(Record1);

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

                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                MessageBox.Show("Done");
            }
            else
            {
                return;
            }
        }

        [CommandMethod("PPL_CONC_OWNER4")]
        public void concatenate_4_fields_in_the_3_one()
        {
            if (isSECURE() == true)
            {
                try
                {
                    string OWN1_LAST = "OWN1_LAST";
                    string OWN1_FRST = "OWN1_FRST";
                    string OWN2_LAST = "OWN2_LAST";
                    string OWN2_FRST = "OWN2_FRST";
                    string LEGAL3 = "LEGAL3";




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



                            foreach (ObjectId Obj_ID1 in BTrecord)
                            {
                                Entity Ent1 = (Entity)Trans1.GetObject(Obj_ID1, OpenMode.ForWrite);
                                if (Ent1 != null)
                                {


                                    using (Records1 = Tables1.GetObjectRecords(Convert.ToUInt32(0), Ent1.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForWrite, false))
                                    {
                                        if (Records1 != null)
                                        {
                                            if (Records1.Count > 0)
                                            {

                                                String Valoare_de_adaugat = "";
                                                String Valoare_de_adaugat1 = "";
                                                String Valoare_de_adaugat2 = "";
                                                String Valoare_de_adaugat3 = "";
                                                String Valoare_de_adaugat4 = "";

                                                Boolean ADD = false;
                                                foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                                {
                                                    Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Record1.TableName];
                                                    Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;


                                                    for (int j = 0; j < Record1.Count; ++j)
                                                    {
                                                        Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[j];
                                                        string Nume_field = Field_def1.Name;
                                                        string Valoare_field = (string)Record1[j].StrValue;
                                                        if (Nume_field == OWN1_LAST) Valoare_de_adaugat2 = Valoare_field;
                                                        if (Nume_field == OWN1_FRST) Valoare_de_adaugat1 = Valoare_field;
                                                        if (Nume_field == OWN2_LAST) Valoare_de_adaugat4 = Valoare_field;
                                                        if (Nume_field == OWN2_FRST) Valoare_de_adaugat3 = Valoare_field;
                                                        if (Nume_field == LEGAL3) ADD = true;
                                                    }

                                                }

                                                if (Valoare_de_adaugat1 != "")
                                                {
                                                    Valoare_de_adaugat = Valoare_de_adaugat1.ToUpper();
                                                }

                                                if (Valoare_de_adaugat2 != "")
                                                {
                                                    if (Valoare_de_adaugat1 != "")
                                                    {
                                                        Valoare_de_adaugat = Valoare_de_adaugat1.ToUpper() + " " + Valoare_de_adaugat2.ToUpper();
                                                    }
                                                    else
                                                    {
                                                        Valoare_de_adaugat = Valoare_de_adaugat2.ToUpper();
                                                    }
                                                }

                                                if (Valoare_de_adaugat3 != "")
                                                {
                                                    if (Valoare_de_adaugat != "")
                                                    {
                                                        Valoare_de_adaugat = Valoare_de_adaugat + ", " + Valoare_de_adaugat3.ToUpper();
                                                    }
                                                    else
                                                    {
                                                        Valoare_de_adaugat = Valoare_de_adaugat3.ToUpper();
                                                    }
                                                }

                                                if (Valoare_de_adaugat4 != "")
                                                {
                                                    if (Valoare_de_adaugat != "")
                                                    {
                                                        Valoare_de_adaugat = Valoare_de_adaugat + " " + Valoare_de_adaugat4.ToUpper();
                                                    }
                                                    else
                                                    {
                                                        Valoare_de_adaugat = Valoare_de_adaugat4.ToUpper();
                                                    }
                                                }

                                                if (ADD == true & Valoare_de_adaugat != "")
                                                {



                                                    foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                                    {
                                                        Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Record1.TableName];
                                                        Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;





                                                        Autodesk.Gis.Map.Utilities.MapValue MapVal;


                                                        for (int j = 0; j < Record1.Count; ++j)
                                                        {
                                                            Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[j];
                                                            string Nume_field = Field_def1.Name;
                                                            string Valoare_field = (string)Record1[j].StrValue;
                                                            MapVal = Record1[j];


                                                            if (Nume_field == LEGAL3 )
                                                            {
                                                                MapVal.Assign(Valoare_de_adaugat);
                                                            }
                                                            else
                                                            {
                                                                MapVal.Assign(Valoare_field);
                                                            }


                                                        }

                                                        Records1.UpdateRecord(Record1);

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

                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                MessageBox.Show("Done");
            }
            else
            {
                return;
            }
        }


        [CommandMethod("PPL_conc_owner2")]
        public void concatenate_2_fields_in_one()
        {
            if (isSECURE() == true)
            {
                try
                {

                    String OD_field1 = "LastName";
                    String OD_field2 = "FirstName";
                    String OD_field3 = "LastName";


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



                            foreach (ObjectId Obj_ID1 in BTrecord)
                            {
                                Entity Ent1 = (Entity)Trans1.GetObject(Obj_ID1, OpenMode.ForWrite);
                                if (Ent1 != null)
                                {


                                    using (Records1 = Tables1.GetObjectRecords(Convert.ToUInt32(0), Ent1.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForWrite, false))
                                    {
                                        if (Records1 != null)
                                        {
                                            if (Records1.Count > 0)
                                            {

                                                String Valoare_de_adaugat = "";
                                                String Valoare_de_adaugat1 = "";
                                                String Valoare_de_adaugat2 = "";

                                                Boolean ADD = false;
                                                foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                                {
                                                    Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Record1.TableName];
                                                    Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;


                                                    for (int j = 0; j < Record1.Count; ++j)
                                                    {
                                                        Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[j];
                                                        string Nume_field = Field_def1.Name;
                                                        string Valoare_field = (string)Record1[j].StrValue;
                                                        if (Nume_field == OD_field1) Valoare_de_adaugat1 = Valoare_field;
                                                        if (Nume_field == OD_field2) Valoare_de_adaugat2 = Valoare_field;
                                                        if (Nume_field == OD_field3) ADD = true;
                                                    }

                                                }

                                                if (Valoare_de_adaugat1 != "")
                                                {
                                                    Valoare_de_adaugat = Valoare_de_adaugat1.ToUpper();
                                                }

                                                if (Valoare_de_adaugat2 != "")
                                                {
                                                    if (Valoare_de_adaugat1 != "")
                                                    {
                                                        Valoare_de_adaugat = Valoare_de_adaugat1.ToUpper() + ", " + Valoare_de_adaugat2.ToUpper();
                                                    }
                                                    else
                                                    {
                                                        Valoare_de_adaugat = Valoare_de_adaugat2.ToUpper();
                                                    }
                                                }

                                                if (ADD == true & Valoare_de_adaugat != "")
                                                {



                                                    foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                                    {
                                                        Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Record1.TableName];
                                                        Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;





                                                        Autodesk.Gis.Map.Utilities.MapValue MapVal;


                                                        for (int j = 0; j < Record1.Count; ++j)
                                                        {
                                                            Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[j];
                                                            string Nume_field = Field_def1.Name;
                                                            string Valoare_field = (string)Record1[j].StrValue;
                                                            MapVal = Record1[j];


                                                            if (Nume_field == OD_field3)
                                                            {
                                                                MapVal.Assign(Valoare_de_adaugat);
                                                            }
                                                            else
                                                            {
                                                                MapVal.Assign(Valoare_field);
                                                            }
                                                        }

                                                        Records1.UpdateRecord(Record1);

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

                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                MessageBox.Show("Done");
            }
            else
            {
                return;
            }
        }

        [CommandMethod("PPL_ELEV_FROM_OD")]
        public void label_from_field()
        {
            if (isSECURE() == true)
            {
                try
                {



                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;


                    Autodesk.AutoCAD.EditorInput.PromptStringOptions Prompt_string = new Autodesk.AutoCAD.EditorInput.PromptStringOptions("\n" + "Specify Suffix:");

                    Prompt_string.AllowSpaces = true;

                    Autodesk.AutoCAD.EditorInput.PromptResult Rezultat_suffix = ThisDrawing.Editor.GetString(Prompt_string);
                    String Suffix1 = "";

                    if (Rezultat_suffix.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                    {
                        Suffix1 = Rezultat_suffix.StringResult;
                    }

                    int Round1 = 0;

                    Autodesk.AutoCAD.EditorInput.PromptIntegerOptions Prompt_int = new Autodesk.AutoCAD.EditorInput.PromptIntegerOptions("\n" + "Specify rounding:");
                    Prompt_int.AllowNegative = false;
                    Prompt_int.AllowZero = true;
                    Prompt_int.AllowNone = true;
                    Autodesk.AutoCAD.EditorInput.PromptIntegerResult Rezultat_int = ThisDrawing.Editor.GetInteger(Prompt_int);

                    if (Rezultat_int.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                    {
                        Round1 = Rezultat_int.Value;
                    }


                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                    object OLD_OSnap = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("OSMODE");
                    object NEW_OSnap = 512;

                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", NEW_OSnap);


            //End = 1,
                //Middle = 2,
                //Center = 4,
                //Node = 8,
                //Quadrant = 16,
                //Intersection = 32,
                //Insertion = 64,
                //Perpendicular = 128,
                //Tangent = 256,
                //Near = 512,
                // Quick = 1024,
                //ApparentIntersection = 2048,
                //Immediate = 65536,
                //AllowTangent = 131072,
                // DisablePerpendicular = 262144,
                //RelativeCartesian = 524288,
                //RelativePolar = 1048576,
                //NoneOverride = 2097152,    


                label1:

                    using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                            Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                            Autodesk.Gis.Map.ObjectData.Records Records1;


                            Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez1 = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rez1.MessageForAdding = "\nSelect the polyline:";
                            Prompt_rez1.SingleOnly = true;
                            Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez1);

                            if (Rezultat1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                            {
                                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }
                            Entity Ent1 = (Entity)Trans1.GetObject(Rezultat1.Value[0].ObjectId, OpenMode.ForRead);







                            if (Ent1 != null)
                            {
                                if (Ent1 is Curve)
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
                                                            string Valoare_field = (string)Record1[j].StrValue;
                                                            if (Nume_field.ToUpper() == "ELEV" & Functions.IsNumeric(Valoare_field) == true)
                                                            {

                                                                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res_position;
                                                                Autodesk.AutoCAD.EditorInput.PromptPointOptions PP_position;
                                                                PP_position = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the insertion point");
                                                                PP_position.AllowNone = false;
                                                                Point_res_position = Editor1.GetPoint(PP_position);

                                                                if (Point_res_position.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                                                {
                                                                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                                                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                                    return;
                                                                }




                                                                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                                                                Autodesk.AutoCAD.EditorInput.PromptPointOptions PP2;
                                                                PP2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the second point (rotation)");
                                                                PP2.AllowNone = false;
                                                                PP2.UseBasePoint = true;
                                                                PP2.BasePoint = Point_res_position.Value;

                                                                Point_res2 = Editor1.GetPoint(PP2);


                                                                if (Point_res2.Status != PromptStatus.OK)
                                                                {

                                                                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                                                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                                    return;
                                                                }

                                                                Point3d Punct1 = new Point3d();
                                                                Punct1 = Point_res_position.Value;//.TransformBy(CurentUCSmatrix);
                                                                Point3d Punct2 = new Point3d();
                                                                Punct2 = Point_res2.Value;//.TransformBy(CurentUCSmatrix);

                                                                double Bearing = Functions.GET_Bearing_rad(Punct1.X, Punct1.Y, Punct2.X, Punct2.Y);


                                                                MText Mtext1 = new MText();
                                                                Mtext1.Contents = Functions.Get_String_Rounded(Convert.ToDouble(Valoare_field), Round1) + Suffix1;
                                                                Mtext1.TextHeight = 6;
                                                                Mtext1.Location = Point_res_position.Value;
                                                                Mtext1.Rotation = Bearing;
                                                                Mtext1.Attachment = AttachmentPoint.MiddleCenter;
                                                                BTrecord.AppendEntity(Mtext1);
                                                                Trans1.AddNewlyCreatedDBObject(Mtext1, true);


                                                                j = Record1.Count;


                                                            }

                                                        }
                                                    }

                                                }
                                                else
                                                {
                                                    if (Ent1 is Polyline)
                                                    {
                                                        Polyline Poly1 = (Polyline)Ent1;

                                                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res_position;
                                                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP_position;
                                                        PP_position = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the insertion point");
                                                        PP_position.AllowNone = false;
                                                        Point_res_position = Editor1.GetPoint(PP_position);

                                                        if (Point_res_position.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                                        {
                                                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                                                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                            return;
                                                        }




                                                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                                                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP2;
                                                        PP2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the second point (rotation)");
                                                        PP2.AllowNone = false;
                                                        PP2.UseBasePoint = true;
                                                        PP2.BasePoint = Point_res_position.Value;

                                                        Point_res2 = Editor1.GetPoint(PP2);


                                                        if (Point_res2.Status != PromptStatus.OK)
                                                        {

                                                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                                                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                            return;
                                                        }

                                                        Point3d Punct1 = new Point3d();
                                                        Punct1 = Point_res_position.Value;//.TransformBy(CurentUCSmatrix);
                                                        Point3d Punct2 = new Point3d();
                                                        Punct2 = Point_res2.Value;//.TransformBy(CurentUCSmatrix);

                                                        double Bearing = Functions.GET_Bearing_rad(Punct1.X, Punct1.Y, Punct2.X, Punct2.Y);


                                                        MText Mtext1 = new MText();
                                                        Mtext1.Contents = "{\\Fromans|c0;" + Functions.Get_String_Rounded(Poly1.Elevation, Round1) + Suffix1 + "}";
                                                        Mtext1.TextHeight = 4;
                                                        Mtext1.Location = Point_res_position.Value;
                                                        Mtext1.Rotation = Bearing;
                                                        Mtext1.Attachment = AttachmentPoint.MiddleCenter;
                                                        BTrecord.AppendEntity(Mtext1);
                                                        Trans1.AddNewlyCreatedDBObject(Mtext1, true);





                                                    }
                                                }

                                            }
                                            else
                                            {
                                                if (Ent1 is Polyline)
                                                {
                                                    Polyline Poly1 = (Polyline)Ent1;

                                                    Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res_position;
                                                    Autodesk.AutoCAD.EditorInput.PromptPointOptions PP_position;
                                                    PP_position = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the insertion point");
                                                    PP_position.AllowNone = false;
                                                    Point_res_position = Editor1.GetPoint(PP_position);

                                                    if (Point_res_position.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                                    {
                                                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                        return;
                                                    }




                                                    Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                                                    Autodesk.AutoCAD.EditorInput.PromptPointOptions PP2;
                                                    PP2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the second point (rotation)");
                                                    PP2.AllowNone = false;
                                                    PP2.UseBasePoint = true;
                                                    PP2.BasePoint = Point_res_position.Value;

                                                    Point_res2 = Editor1.GetPoint(PP2);


                                                    if (Point_res2.Status != PromptStatus.OK)
                                                    {

                                                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                        return;
                                                    }

                                                    Point3d Punct1 = new Point3d();
                                                    Punct1 = Point_res_position.Value;//.TransformBy(CurentUCSmatrix);
                                                    Point3d Punct2 = new Point3d();
                                                    Punct2 = Point_res2.Value;//.TransformBy(CurentUCSmatrix);

                                                    double Bearing = Functions.GET_Bearing_rad(Punct1.X, Punct1.Y, Punct2.X, Punct2.Y);


                                                    MText Mtext1 = new MText();
                                                    Mtext1.Contents = "{\\Fromans|c0;" + Functions.Get_String_Rounded(Poly1.Elevation, Round1) + Suffix1 + "}";
                                                    Mtext1.TextHeight = 4;
                                                    Mtext1.Location = Point_res_position.Value;
                                                    Mtext1.Rotation = Bearing;
                                                    Mtext1.Attachment = AttachmentPoint.MiddleCenter;
                                                    BTrecord.AppendEntity(Mtext1);
                                                    Trans1.AddNewlyCreatedDBObject(Mtext1, true);





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

                        goto label1;
                    }

                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                return;
            }
        }


        [CommandMethod("PPL_ADD_OD")]
        public void Show_od_ATTACH_FORM()
        {
            if (isSECURE() == true)
            {
                foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
                {
                    if (Forma1 is OD_attach)
                    {
                        Forma1.Focus();
                        Forma1.WindowState = System.Windows.Forms.FormWindowState.Normal;
                        return;
                    }
                }

                try
                {
                    OD_attach forma2 = new OD_attach();
                    Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma2);
                }
                catch (System.Exception EX)
                {
                    MessageBox.Show(EX.Message);
                }


            }
            else
            {
                return;
            }
        }


        [CommandMethod("PPL_OD_TABLE")]
        public void Show_OD_TABLE_FORM()
        {
            if (isSECURE() == true)
            {
                foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
                {
                    if (Forma1 is OD_TABLE_form)
                    {
                        Forma1.Focus();
                        Forma1.WindowState = System.Windows.Forms.FormWindowState.Normal;
                        return;
                    }
                }

                try
                {
                    OD_TABLE_form forma2 = new OD_TABLE_form();
                    Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma2);
                }
                catch (System.Exception EX)
                {
                    MessageBox.Show(EX.Message);
                }


            }
            else
            {
                return;
            }
        }





    }
}
