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
    public partial class contours_form : Form
    {
        private bool clickdragdown;
        private Point lastLocation;




        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();

            lista_butoane.Add(comboBox_field);
            lista_butoane.Add(comboBox_scales);
            lista_butoane.Add(comboBox_text_styles);
            lista_butoane.Add(radioButton_use_elevation);





            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {

                bt1.Enabled = false;

            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();

            lista_butoane.Add(comboBox_field);
            lista_butoane.Add(comboBox_scales);
            lista_butoane.Add(comboBox_text_styles);
            lista_butoane.Add(radioButton_use_elevation);
            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }



        public contours_form()
        {
            InitializeComponent();

            comboBox_scales.SelectedIndex = 0;
            //if (Functions.is_dan_popescu() == true || Functions.is_eli_barboza() == true || Functions.is_hector_morales() == true
            //    || Functions.is_chad_mikkelsen() == true || Functions.is_richard_pangburn() == true)
            //{
            //    button_show_tools.Visible = true;
            //}

            _AGEN_mainform.config_path = "";
        }

        private void button_Exit_Click(object sender, EventArgs e)
        {
            try
            {

                int i = 0;

                do
                {
                    System.Windows.Forms.Form Forma1 = System.Windows.Forms.Application.OpenForms[i];
                    if (Forma1 is AGEN_custom_band_form)
                    {
                        Forma1.Close();
                    }

                    i = i + 1;

                } while (i < System.Windows.Forms.Application.OpenForms.Count);

            }
            catch (InvalidOperationException ex)
            {

            }

            this.Close();


        }
        private void button_minimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
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



        public void set_radioButton_canada(bool val)
        {
            radioButton_use_elevation.Checked = val;
        }
        public void set_radioButton_usa(bool val)
        {
            radioButton_use_OD.Checked = val;
        }

        private void button_refresh_text_styles_Click(object sender, EventArgs e)
        {
            try
            {
                set_enable_false();
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Functions.Incarca_existing_textstyles_to_combobox(comboBox_text_styles);

                        comboBox_text_styles.SelectedIndex = 0;


                        Trans1.Dispose();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            set_enable_true();

        }
        private void button_load_od_field_to_combobox_dropdown(object sender, EventArgs e)
        {

            set_enable_false();
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();





                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        comboBox_field.Items.Clear();

                        Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                        System.Collections.Specialized.StringCollection Nume_tables = new System.Collections.Specialized.StringCollection();
                        Nume_tables = Tables1.GetTableNames();
                        comboBox_field.Items.Clear();
                        for (int i = 0; i < Nume_tables.Count; ++i)
                        {
                            string Tabla1 = Nume_tables[i];

                            Functions.add_all_OD_fieds_to_combobox(Tabla1, comboBox_field);
                        }
                        this.Refresh();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            set_enable_true();

        }
        private void button_draw_Click(object sender, EventArgs e)
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
                    this.WindowState = FormWindowState.Minimized;

                    Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                    Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                    Prompt_rez.MessageForAdding = "\nSelect the contours:";
                    Prompt_rez.SingleOnly = false;
                    Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                    if (Rezultat1.Status != PromptStatus.OK)
                    {

                        this.WindowState = FormWindowState.Normal;
                        set_enable_true();
                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                        return;
                    }

                    double th = 1;
                    if (Functions.IsNumeric(textBox_text_height.Text) == true)
                    {
                        th = Convert.ToDouble(textBox_text_height.Text);
                    }

                    if (comboBox_scales.SelectedIndex > 0)
                    {
                        if (comboBox_scales.Text != "")
                        {
                            string txt = comboBox_scales.Text.Replace("1:", "");
                            if (Functions.IsNumeric(txt) == true)
                            {
                                th = th * Convert.ToDouble(txt);
                            }
                        }
                    }

                    int round1 = comboBox_precision.Text.Replace(".", "").Length - 1;

                    List<int> lista10 = new List<int>();
                    for (int i = -8000; i < 8001; i += 10)
                    {
                        lista10.Add(i);
                    }

                    bool run1 = true;

                    do
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                            TextStyleTable Text_style_table1 = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                            Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

                            ObjectId txtstyleid = ObjectId.Null;

                            foreach (ObjectId id1 in Text_style_table1)
                            {
                                TextStyleTableRecord st1 = Trans1.GetObject(id1, OpenMode.ForRead) as TextStyleTableRecord;
                                if (st1 != null)
                                {
                                    if (st1.Name == comboBox_text_styles.Text)
                                    {
                                        txtstyleid = id1;
                                    }
                                }
                            }


                            string txt_elev = "XX";


                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                            Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                            PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nStart:");
                            PP1.AllowNone = false;
                            Point_res1 = Editor1.GetPoint(PP1);

                            if (Point_res1.Status != PromptStatus.OK)
                            {
                                run1 = false;
                                this.WindowState = FormWindowState.Normal;
                                set_enable_true();
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }

                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                            Autodesk.AutoCAD.EditorInput.PromptPointOptions PP2;
                            PP2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nEnd:");
                            PP2.AllowNone = false;
                            PP2.UseBasePoint = true;
                            PP2.BasePoint = Point_res1.Value;


                            Point_res2 = Editor1.GetPoint(PP2);

                            if (Point_res2.Status != PromptStatus.OK)
                            {
                                run1 = false;
                                this.WindowState = FormWindowState.Normal;
                                set_enable_true();
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }

                            Point2d pt1 = new Point2d(Point_res1.Value.X, Point_res1.Value.Y);
                            Point2d pt2 = new Point2d(Point_res2.Value.X, Point_res2.Value.Y);
                            Polyline poly1 = new Polyline();
                            poly1.AddVertexAt(0, pt1, 0, 0, 0);
                            poly1.AddVertexAt(1, pt2, 0, 0, 0);

                            for (int i = 0; i < Rezultat1.Value.Count; i++)
                            {
                                Polyline poly2 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as Polyline;
                                Polyline3d poly3D2 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as Polyline3d;
                                Line line2 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as Line;



                                if (poly2 != null)
                                {
                                    poly1.Elevation = poly2.Elevation;
                                    Point3dCollection colint = Functions.Intersect_on_both_operands(poly1, poly2);
                                    if (colint.Count > 0)
                                    {
                                        for (int j = 0; j < colint.Count; j++)
                                        {
                                            Point3d ptins = colint[j];

                                            double index_m = poly2.GetParameterAtPoint(poly2.GetClosestPointTo(ptins, Vector3d.ZAxis, false));
                                            double index1 = Math.Floor(index_m);
                                            double index2 = Math.Ceiling(index_m);

                                            if (index1 == index2)
                                            {
                                                if (index1 == 0)
                                                {
                                                    index2 = 1;
                                                }
                                                else if (index1 == poly2.NumberOfVertices - 1)
                                                {
                                                    index2 = poly2.NumberOfVertices - 2;
                                                }
                                                else
                                                {
                                                    index2 = index1 + 1;
                                                }
                                            }

                                            Point3d p_poly1 = poly2.GetPointAtParameter(index1);
                                            Point3d p_poly2 = poly2.GetPointAtParameter(index2);

                                            double bearing2 = Functions.GET_Bearing_rad(p_poly1.X, p_poly1.Y, p_poly2.X, p_poly2.Y);

                                            if (bearing2 > Math.PI / 2 && bearing2 < Math.PI)
                                            {
                                                bearing2 = bearing2 + Math.PI;
                                            }
                                            else if (bearing2 >= Math.PI && bearing2 <= 1.5 * Math.PI)
                                            {
                                                bearing2 = bearing2 - Math.PI;
                                            }

                                            if (checkBox_rotate_180.Checked == true)
                                            {
                                                bearing2 = bearing2 + Math.PI;
                                            }

                                            #region object data
                                            if (radioButton_use_OD.Checked == true && comboBox_field.Text != "")
                                            {
                                                using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tables1.GetObjectRecords(Convert.ToUInt32(0), Rezultat1.Value[i].ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, false))
                                                {
                                                    if (Records1 != null)
                                                    {
                                                        if (Records1.Count > 0)
                                                        {

                                                            foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                                            {
                                                                Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Record1.TableName];
                                                                Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;
                                                                for (int k = 0; k < Record1.Count; ++k)
                                                                {
                                                                    Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[k];
                                                                    string Nume_field = Field_def1.Name;
                                                                    object valoare1 = Record1[k].StrValue;
                                                                    if (Nume_field == comboBox_field.Text)
                                                                    {
                                                                        txt_elev = Convert.ToString(valoare1);
                                                                        k = Record1.Count;
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            #endregion
                                            else
                                            {
                                                txt_elev = Functions.Get_String_Rounded(poly2.Elevation, round1);
                                            }

                                            if (textBox_find.Text != "")
                                            {
                                                txt_elev = txt_elev.Replace(textBox_find.Text, textBox_replace.Text);
                                            }


                                            if (Functions.IsNumeric(txt_elev) == true)
                                            {
                                                double val1 = Convert.ToDouble(txt_elev);

                                                txt_elev = Functions.Get_String_Rounded(val1, round1);
                                            }

                                            bool run = false;
                                            string last_letter = txt_elev.Substring(txt_elev.Length - 1);
                                            if (checkBox_label_only_10.Checked == true && last_letter == "0") run = true;
                                            if (checkBox_label_only_10.Checked == false) run = true;

                                            if (txt_elev != "XX" && run == true)
                                            {
                                                txt_elev = txt_elev + textBox_suffix.Text;

                                                MText mtxt2 = new MText();
                                                if (txtstyleid != ObjectId.Null)
                                                {
                                                    mtxt2.TextStyleId = txtstyleid;
                                                }

                                                mtxt2.TextHeight = th;
                                                mtxt2.Contents = txt_elev;
                                                mtxt2.UseBackgroundColor = true;
                                                mtxt2.BackgroundFill = true;
                                                mtxt2.BackgroundScaleFactor = 1.2;
                                                mtxt2.Location = ptins;
                                                mtxt2.Attachment = AttachmentPoint.MiddleCenter;
                                                mtxt2.Rotation = bearing2;
                                                BTrecord.AppendEntity(mtxt2);
                                                Trans1.AddNewlyCreatedDBObject(mtxt2, true);
                                                txt_elev = "XX";
                                            }

                                        }
                                    }

                                }
                                else if (poly3D2 != null)
                                {

                                    poly2 = Functions.Build_2dpoly_from_3d(poly3D2);
                                    poly2.Elevation = 0;
                                    poly1.Elevation = poly2.Elevation;
                                    Point3dCollection colint = Functions.Intersect_on_both_operands(poly1, poly2);
                                    if (colint.Count > 0)
                                    {
                                        for (int j = 0; j < colint.Count; j++)
                                        {
                                            Point3d ptins = colint[j];

                                            double param = poly2.GetParameterAtPoint(poly2.GetClosestPointTo(ptins, Vector3d.ZAxis, false));
                                            Point3d pt3d = poly3D2.GetPointAtParameter(param);

                                            double index1 = Math.Floor(param);
                                            double index2 = Math.Ceiling(param);

                                            if (index1 == index2)
                                            {
                                                if (index1 == 0)
                                                {
                                                    index2 = 1;
                                                }
                                                else if (index1 == poly2.NumberOfVertices - 1)
                                                {
                                                    index2 = poly2.NumberOfVertices - 2;
                                                }
                                                else
                                                {
                                                    index2 = index1 + 1;
                                                }
                                            }

                                            Point3d p_poly1 = poly2.GetPointAtParameter(index1);
                                            Point3d p_poly2 = poly2.GetPointAtParameter(index2);

                                            double bearing2 = Functions.GET_Bearing_rad(p_poly1.X, p_poly1.Y, p_poly2.X, p_poly2.Y);

                                            if (bearing2 > Math.PI / 2 && bearing2 < Math.PI)
                                            {
                                                bearing2 = bearing2 + Math.PI;
                                            }
                                            else if (bearing2 >= Math.PI && bearing2 <= 1.5 * Math.PI)
                                            {
                                                bearing2 = bearing2 - Math.PI;
                                            }

                                            if (checkBox_rotate_180.Checked == true)
                                            {
                                                bearing2 = bearing2 + Math.PI;
                                            }

                                            #region object data
                                            if (radioButton_use_OD.Checked == true && comboBox_field.Text != "")
                                            {
                                                using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tables1.GetObjectRecords(Convert.ToUInt32(0), Rezultat1.Value[i].ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, false))
                                                {
                                                    if (Records1 != null)
                                                    {
                                                        if (Records1.Count > 0)
                                                        {

                                                            foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                                            {
                                                                Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Record1.TableName];
                                                                Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;
                                                                for (int k = 0; k < Record1.Count; ++k)
                                                                {
                                                                    Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[k];
                                                                    string Nume_field = Field_def1.Name;
                                                                    object valoare1 = Record1[k].StrValue;
                                                                    if (Nume_field == comboBox_field.Text)
                                                                    {
                                                                        txt_elev = Convert.ToString(valoare1);
                                                                        k = Record1.Count;
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            #endregion
                                            else
                                            {

                                                txt_elev = Functions.Get_String_Rounded(pt3d.Z, round1);

                                            }


                                            if (textBox_find.Text != "")
                                            {
                                                txt_elev = txt_elev.Replace(textBox_find.Text, textBox_replace.Text);
                                            }

                                            if (Functions.IsNumeric(txt_elev) == true)
                                            {
                                                double val1 = Convert.ToDouble(txt_elev);

                                                txt_elev = Functions.Get_String_Rounded(val1, round1);
                                            }

                                            bool run = false;
                                            string last_letter = txt_elev.Substring(txt_elev.Length - 1);
                                            if (checkBox_label_only_10.Checked == true && last_letter == "0") run = true;
                                            if (checkBox_label_only_10.Checked == false) run = true;

                                            if (txt_elev != "XX" && run == true)
                                            {
                                                txt_elev = txt_elev + textBox_suffix.Text;
                                                MText mtxt2 = new MText();
                                                if (txtstyleid != ObjectId.Null)
                                                {
                                                    mtxt2.TextStyleId = txtstyleid;
                                                }

                                                mtxt2.TextHeight = th;
                                                mtxt2.Contents = txt_elev;
                                                mtxt2.UseBackgroundColor = true;
                                                mtxt2.BackgroundFill = true;
                                                mtxt2.BackgroundScaleFactor = 1.2;
                                                mtxt2.Location = ptins;
                                                mtxt2.Attachment = AttachmentPoint.MiddleCenter;
                                                mtxt2.Rotation = bearing2;
                                                BTrecord.AppendEntity(mtxt2);
                                                Trans1.AddNewlyCreatedDBObject(mtxt2, true);
                                                txt_elev = "XX";
                                            }

                                        }
                                    }

                                }
                                else if (line2 != null)
                                {

                                    poly2 = new Polyline();
                                    poly2.AddVertexAt(0, new Point2d(line2.StartPoint.X, line2.StartPoint.Y), 0, 0, 0);
                                    poly2.AddVertexAt(1, new Point2d(line2.EndPoint.X, line2.EndPoint.Y), 0, 0, 0);
                                    poly2.Elevation = 0;
                                    poly1.Elevation = poly2.Elevation;
                                    Point3dCollection colint = Functions.Intersect_on_both_operands(poly1, poly2);
                                    if (colint.Count > 0)
                                    {
                                        for (int j = 0; j < colint.Count; j++)
                                        {
                                            Point3d ptins = colint[j];
                                            double bearing2 = line2.Angle;

                                            if (bearing2 > Math.PI / 2 && bearing2 < Math.PI)
                                            {
                                                bearing2 = bearing2 + Math.PI;
                                            }
                                            else if (bearing2 >= Math.PI && bearing2 <= 1.5 * Math.PI)
                                            {
                                                bearing2 = bearing2 - Math.PI;
                                            }


                                            #region object data
                                            if (radioButton_use_OD.Checked == true && comboBox_field.Text != "")
                                            {
                                                using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tables1.GetObjectRecords(Convert.ToUInt32(0), Rezultat1.Value[i].ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, false))
                                                {
                                                    if (Records1 != null)
                                                    {
                                                        if (Records1.Count > 0)
                                                        {

                                                            foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                                            {
                                                                Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Record1.TableName];
                                                                Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;
                                                                for (int k = 0; k < Record1.Count; ++k)
                                                                {
                                                                    Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[k];
                                                                    string Nume_field = Field_def1.Name;
                                                                    object valoare1 = Record1[k].StrValue;
                                                                    if (Nume_field == comboBox_field.Text)
                                                                    {
                                                                        txt_elev = Convert.ToString(valoare1);
                                                                        k = Record1.Count;
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            #endregion
                                            else
                                            {
                                                Point3d p1 = line2.StartPoint;
                                                Point3d p2 = line2.EndPoint;

                                                double Z1 = p1.Z;
                                                double Z2 = p2.Z;
                                                double Z = Z1;


                                                double L = line2.Length;
                                                double L1 = Math.Pow(Math.Pow((p1.X - ptins.X), 2) + Math.Pow((p1.Y - ptins.Y), 2) + Math.Pow((p1.Z - ptins.Z), 2), 0.5);
                                                if (L >= L1)
                                                {
                                                    Z = Z1 + L1 * (Z2 - Z1) / L;
                                                }


                                                txt_elev = Functions.Get_String_Rounded(Z, 0) + "'";
                                            }
                                            bool run = false;
                                            string last_letter = txt_elev.Substring(txt_elev.Length - 1);
                                            if (checkBox_label_only_10.Checked == true && last_letter == "0") run = true;
                                            if (checkBox_label_only_10.Checked == false) run = true;

                                            if (txt_elev != "XX" && run == true)
                                            {
                                                MText mtxt2 = new MText();
                                                if (txtstyleid != ObjectId.Null)
                                                {
                                                    mtxt2.TextStyleId = txtstyleid;
                                                }

                                                mtxt2.TextHeight = th;
                                                mtxt2.Contents = txt_elev;
                                                mtxt2.UseBackgroundColor = true;
                                                mtxt2.BackgroundFill = true;
                                                mtxt2.BackgroundScaleFactor = 1.2;
                                                mtxt2.Location = ptins;
                                                mtxt2.Attachment = AttachmentPoint.MiddleCenter;
                                                mtxt2.Rotation = bearing2;
                                                BTrecord.AppendEntity(mtxt2);
                                                Trans1.AddNewlyCreatedDBObject(mtxt2, true);
                                                txt_elev = "XX";
                                            }

                                        }
                                    }

                                }


                            }






                            Trans1.Commit();
                        }
                    } while (run1 == true);

                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            set_enable_true();
            this.WindowState = FormWindowState.Normal;

        }
    }
}
