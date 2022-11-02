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
    public partial class AttachOD_form : Form
    {
        private bool clickdragdown;
        private Point lastLocation;

        public AttachOD_form()
        {
            InitializeComponent();
        }

        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_attach_OD);
            lista_butoane.Add(textBox_default_value);
            lista_butoane.Add(textBox_description);
            lista_butoane.Add(textBox_field);
            lista_butoane.Add(textBox_row_start);
            lista_butoane.Add(textBox_table_name);
            lista_butoane.Add(textBox_type);
            lista_butoane.Add(textBox_value);
            lista_butoane.Add(button_refresh_ws);
            lista_butoane.Add(comboBox_ws);


            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = false;
            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_attach_OD);
            lista_butoane.Add(textBox_default_value);
            lista_butoane.Add(textBox_description);
            lista_butoane.Add(textBox_field);
            lista_butoane.Add(textBox_row_start);
            lista_butoane.Add(textBox_table_name);
            lista_butoane.Add(textBox_type);
            lista_butoane.Add(textBox_value);
            lista_butoane.Add(button_refresh_ws);
            lista_butoane.Add(comboBox_ws);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }


        private void clickmove_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            clickdragdown = true;
            lastLocation = e.Location;
        }

        private void clickmove_MouseMove(object sender, MouseEventArgs e)
        {
            if (clickdragdown)
            {
                this.Location = new Point(
                  (this.Location.X - lastLocation.X) + e.X, (this.Location.Y - lastLocation.Y) + e.Y);

                this.Update();
            }
        }

        private void clickmove_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            clickdragdown = false;
        }
        private void button_Exit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void button_minimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void button_refresh_ws_Click(object sender, EventArgs e)
        {
            Functions.Load_opened_worksheets_to_combobox(comboBox_ws);
            if (comboBox_ws.Items.Count > 0)
            {
                for (int i = 0; i < comboBox_ws.Items.Count; ++i)
                {

                    comboBox_ws.SelectedIndex = i;
                    i = comboBox_ws.Items.Count;

                }
            }
        }

        private void button_attach_OD_Click(object sender, EventArgs e)
        {
            set_enable_false();
            try
            {
                if (comboBox_ws.Text != "")
                {
                    string string1 = comboBox_ws.Text;
                    if (string1.Contains("[") == true && string1.Contains("]") == true)
                    {
                        string filename = string1.Substring(string1.IndexOf("]") + 4, string1.Length - (string1.IndexOf("]") + 4));
                        string sheet_name = string1.Substring(1, string1.IndexOf("]") - 1);
                        if (filename.Length > 0 && sheet_name.Length > 0)
                        {
                            Microsoft.Office.Interop.Excel.Worksheet W1 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, sheet_name);
                            if (W1 != null && textBox_table_name.Text != "")
                            {
                                string table_name = W1.Range[textBox_table_name.Text].Value2;

                                if (Functions.IsNumeric(textBox_row_start.Text) == true && table_name != null && table_name.Replace(" ", "") != "")
                                {
                                    table_name = table_name.Replace(" ", "");
                                    int start1 = Convert.ToInt32(textBox_row_start.Text);
                                    System.Data.DataTable dt1 = new System.Data.DataTable();
                                    dt1.Columns.Add("Field", typeof(string));
                                    dt1.Columns.Add("Type", typeof(string));
                                    dt1.Columns.Add("Description", typeof(string));
                                    dt1.Columns.Add("Default", typeof(string));
                                    dt1.Columns.Add("Value", typeof(string));
                                    dt1 = Functions.Populate_data_table_from_excel(dt1, W1, start1, textBox_field.Text, "", "", "", "", "", "", "", "", "", "", true);
                                    if (dt1.Rows.Count > 0)
                                    {

                                        List<string> Lista_field_name = new List<string>();
                                        List<Autodesk.Gis.Map.Constants.DataType> Lista_field_type = new List<Autodesk.Gis.Map.Constants.DataType>();
                                        List<string> Lista_field_description = new List<string>();
                                        List<object> Lista_val = new List<object>();
                                        List<Autodesk.Gis.Map.Constants.DataType> Lista_type = new List<Autodesk.Gis.Map.Constants.DataType>();




                                        for (int i = 0; i < dt1.Rows.Count; ++i)
                                        {
                                            if (dt1.Rows[i][0] != DBNull.Value)
                                            {
                                                string field1 = Convert.ToString(dt1.Rows[i][0]);
                                                if (Lista_field_name.Contains(field1) == false && field1.Replace(" ", "") != "")
                                                {
                                                    Lista_field_name.Add(field1.Replace(" ", ""));
                                                    Autodesk.Gis.Map.Constants.DataType type1 = Autodesk.Gis.Map.Constants.DataType.Character;


                                                    if (dt1.Rows[i][1] != DBNull.Value)
                                                    {
                                                        string type2 = Convert.ToString(dt1.Rows[i][1]);
                                                        if (type2.ToLower() == "character")
                                                        {
                                                            type1 = Autodesk.Gis.Map.Constants.DataType.Character;
                                                        }

                                                        if (type2.ToLower() == "integer")
                                                        {
                                                            type1 = Autodesk.Gis.Map.Constants.DataType.Integer;
                                                        }
                                                        if (type2.ToLower() == "real")
                                                        {
                                                            type1 = Autodesk.Gis.Map.Constants.DataType.Real;
                                                        }
                                                        if (type2.ToLower() == "point")
                                                        {
                                                            type1 = Autodesk.Gis.Map.Constants.DataType.Point;
                                                        }
                                                    }
                                                    Lista_field_type.Add(type1);
                                                    Lista_type.Add(type1);


                                                    string description1 = field1;
                                                    if (dt1.Rows[i][2] != DBNull.Value)
                                                    {
                                                        description1 = Convert.ToString(dt1.Rows[i][2]);
                                                    }

                                                    Lista_field_description.Add(description1);


                                                    if (dt1.Rows[i][4] != DBNull.Value)
                                                    {
                                                        string val1 = Convert.ToString(dt1.Rows[i][4]);
                                                        if (type1 == Autodesk.Gis.Map.Constants.DataType.Character)
                                                        {
                                                            Lista_val.Add(val1);
                                                        }
                                                        else if (type1 == Autodesk.Gis.Map.Constants.DataType.Real && Functions.IsNumeric(val1) == true)
                                                        {
                                                            Lista_val.Add(Convert.ToDouble(val1));
                                                        }
                                                        else if (type1 == Autodesk.Gis.Map.Constants.DataType.Integer && Functions.IsNumeric(val1) == true)
                                                        {
                                                            Lista_val.Add(Convert.ToInt32(val1));
                                                        }
                                                        else if (type1 == Autodesk.Gis.Map.Constants.DataType.Point)
                                                        {
                                                            Lista_val.Add(new Point3d(0, 0, 0));
                                                        }
                                                    }

                                                    else if (dt1.Rows[i][3] != DBNull.Value && dt1.Rows[i][4] == DBNull.Value)
                                                    {
                                                        string val1 = Convert.ToString(dt1.Rows[i][3]);
                                                        if (type1 == Autodesk.Gis.Map.Constants.DataType.Character)
                                                        {
                                                            Lista_val.Add(val1);
                                                        }
                                                        else if (type1 == Autodesk.Gis.Map.Constants.DataType.Real && Functions.IsNumeric(val1) == true)
                                                        {
                                                            Lista_val.Add(Convert.ToDouble(val1));
                                                        }
                                                        else if (type1 == Autodesk.Gis.Map.Constants.DataType.Integer && Functions.IsNumeric(val1) == true)
                                                        {
                                                            Lista_val.Add(Convert.ToInt32(val1));
                                                        }
                                                        else if (type1 == Autodesk.Gis.Map.Constants.DataType.Point)
                                                        {
                                                            Lista_val.Add(new Point3d(0, 0, 0));
                                                        }
                                                    }
                                                    else
                                                    {
                                                        if (type1 == Autodesk.Gis.Map.Constants.DataType.Character)
                                                        {
                                                            Lista_val.Add("");
                                                        }
                                                        else if (type1 == Autodesk.Gis.Map.Constants.DataType.Real)
                                                        {
                                                            Lista_val.Add(0);
                                                        }
                                                        else if (type1 == Autodesk.Gis.Map.Constants.DataType.Integer )
                                                        {
                                                            Lista_val.Add(0);
                                                        }
                                                        else if (type1 == Autodesk.Gis.Map.Constants.DataType.Point)
                                                        {
                                                            Lista_val.Add(new Point3d(0,0,0));
                                                        }
                                                    }

                                                }
                                                else
                                                {
                                                    MessageBox.Show("please verify your excel spreadsheet\r\noperation aborted");
                                                    set_enable_true();
                                                    return;
                                                }
                                            }

                                        }
                                        ObjectId[] Empty_array = null;
                                        Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

                                        using (DocumentLock lock1 = ThisDrawing.LockDocument())
                                        {
                                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                                            {
                                                BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                                                #region OBJECT DATA
                                                Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                                                Functions.Create_od_table(table_name, Lista_field_name, Lista_field_type, Lista_field_description);
                                                #endregion


                                                ThisDrawing.Editor.SetImpliedSelection(Empty_array);
                                                Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                                                Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                                                Prompt_rez.MessageForAdding = "\nSelect the objects:";
                                                Prompt_rez.SingleOnly = true;
                                                Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                                                if (Rezultat1.Status != PromptStatus.OK)
                                                {
                                                    Trans1.Commit();
                                                    set_enable_true();
                                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                    return;
                                                }


                                                for (int i = 0; i < Rezultat1.Value.Count; ++i)
                                                {
                                                    Entity Ent1 = Rezultat1.Value[i].ObjectId.GetObject(OpenMode.ForRead) as Entity;
                                                    if (Ent1 != null)
                                                    {
                                                        Functions.Populate_object_data_table_from_objectid(Tables1, Ent1.ObjectId, table_name, Lista_val, Lista_type);
                                                    }
                                                }
                                                ThisDrawing.Editor.SetImpliedSelection(Empty_array);
                                                Trans1.Commit();
                                            }
                                        }
                                    }


                                }
                                else
                                {
                                    MessageBox.Show("please verify your excel spreadsheet\r\noperation aborted");
                                }
                            }
                            else
                            {
                                MessageBox.Show("no excel specified\r\noperation aborted");
                            }
                        }
                        else
                        {
                            MessageBox.Show("no excel specified\r\noperation aborted");
                        }
                    }
                    else
                    {
                        MessageBox.Show("no excel specified\r\noperation aborted");
                    }
                }
                else
                {
                    MessageBox.Show("no excel specified\r\noperation aborted");
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            set_enable_true();
        }
    }
}
