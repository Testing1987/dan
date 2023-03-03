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
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.Gis.Map;
using Autodesk.Gis.Map.ObjectData;

namespace Alignment_mdi
{
    public partial class scales_form : Form
    {


        string Col_MMid = "MMID";
        string Col_handle = "AcadHandle";
        string Col_dwg_name = "DwgNo";
        string Col_M1 = "StaBeg";
        string Col_M2 = "StaEnd";
        string Col_dispM1 = "Disp_StaBeg";
        string Col_dispM2 = "Disp_StaEnd";
        string Col_length = "Length";

        string Col_Width = "Width";
        string Col_Height = "Height";
        string Col_X1 = "X_Beg";
        string Col_Y1 = "Y_Beg";
        string Col_X2 = "X_End";
        string Col_Y2 = "Y_End";

        string col_scale = "Scale";
        string col_scaleName = "ScaleName";

        public SGEN_Sheet_Index forma3 = null;

        bool clickdragdown;
        Point lastLocation;


        public scales_form()
        {
            InitializeComponent();

        }



        private void clickmove_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            clickdragdown = true;
            lastLocation = e.Location;
        }
        private void clickmove_MouseMove(object sender, MouseEventArgs e)
        {
            if (clickdragdown == true)
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



        private void button_close_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button_minimize_Click_1(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        public double get_current_scale()
        {
            double scale1 = 1;
            if (radioButton1.Checked == true) scale1 = 1;
            if (radioButton10.Checked == true) scale1 = 10;
            if (radioButton20.Checked == true) scale1 = 20;
            if (radioButton30.Checked == true) scale1 = 30;
            if (radioButton40.Checked == true) scale1 = 40;
            if (radioButton50.Checked == true) scale1 = 50;
            if (radioButton60.Checked == true) scale1 = 60;
            if (radioButton100.Checked == true) scale1 = 100;
            if (radioButton200.Checked == true) scale1 = 200;
            if (radioButton300.Checked == true) scale1 = 300;
            if (radioButton400.Checked == true) scale1 = 400;
            if (radioButton500.Checked == true) scale1 = 500;
            if (radioButton600.Checked == true) scale1 = 600;
            if (radioButton1000.Checked == true) scale1 = 1000;
            if (radioButton2000.Checked == true) scale1 = 2000;
            if (radioButton3000.Checked == true) scale1 = 3000;
            if (radioButton4000.Checked == true) scale1 = 4000;
            if (radioButton5000.Checked == true) scale1 = 5000;
            if (radioButton6000.Checked == true) scale1 = 6000;

            return 1 / scale1;
        }
        public string get_current_scale_name()
        {
            string scale1 = "1:1";
            if (radioButton1.Checked == true) return "1:1";
            if (radioButton10.Checked == true) return "1:10";
            if (radioButton20.Checked == true) return "1:20";
            if (radioButton30.Checked == true) return "1:30";
            if (radioButton40.Checked == true) return "1:40";
            if (radioButton50.Checked == true) return "1:50";
            if (radioButton60.Checked == true) return "1:60";
            if (radioButton100.Checked == true) return "1:100";
            if (radioButton200.Checked == true) return "1:200";
            if (radioButton300.Checked == true) return "1:300";
            if (radioButton400.Checked == true) return "1:400";
            if (radioButton500.Checked == true) return "1:500";
            if (radioButton600.Checked == true) return "1:600";
            if (radioButton1000.Checked == true) return "1:1000";
            if (radioButton2000.Checked == true) return "1:2000";
            if (radioButton3000.Checked == true) return "1:3000";
            if (radioButton4000.Checked == true) return "1:4000";
            if (radioButton5000.Checked == true) return "1:5000";
            if (radioButton6000.Checked == true) return "1:6000";

            return scale1;
        }

        private void button_place_rectangles_Click(object sender, EventArgs e)
        {
            foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
            {
                if (Forma1 is SGEN_Sheet_Index)
                {

                    forma3 = Forma1 as SGEN_Sheet_Index;

                }
            }

            if (forma3 == null) return;
            if (forma3.get_checkBox_pick_name_from_OD() == false)
            {
                return;
            }

            string odfield1 = forma3.Get_comboBox_od_field();
            if (odfield1 == "")
            {
                return;
            }


            if (System.IO.Directory.Exists(_SGEN_mainform.project_main_folder) == false)
            {
                MessageBox.Show("no config file loaded\r\nOperation aborted");
                return;
            }




            if (_SGEN_mainform.Vw_height == 0 || _SGEN_mainform.Vw_width == 0)
            {
                MessageBox.Show("you do not have the dimensions for the matchline rectangles\r\nOperation aborted");
                return;
            }

            forma3.Create_ML_object_dataTABLE();

            ObjectId[] Empty_array = null;

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    Functions.Creaza_layer(_SGEN_mainform.Layer_name_ML_rectangle, 4, false);

                    if (_SGEN_mainform.dt_sheet_index == null)
                    {
                        _SGEN_mainform.dt_sheet_index = Creaza_sheet_index_datatable_structure();
                    }

                    string Scale1 = _SGEN_mainform.tpage_settings.Get_combobox_viewport_scale_text();


                    if (Functions.IsNumeric(Scale1) == true)
                    {
                        _SGEN_mainform.Vw_scale = Convert.ToDouble(Scale1);
                    }
                    else
                    {
                        if (Scale1.Contains(":") == true)
                        {
                            Scale1 = Scale1.Replace("1:", "");
                            if (Functions.IsNumeric(Scale1) == true)
                            {
                                _SGEN_mainform.Vw_scale = 1000 / Convert.ToDouble(Scale1);
                            }
                        }
                        else
                        {
                            string inch = "\u0022";

                            if (Scale1.Contains(inch + "=") == true && Scale1.Contains("'") == true)
                            {
                                Scale1 = Scale1.Replace("1" + inch + "=", "");
                                Scale1 = Scale1.Substring(0, Scale1.Length - 1);
                            }

                            inch = "\u0094";

                            if (Scale1.Contains(inch + "=") == true && Scale1.Contains("'") == true)
                            {
                                Scale1 = Scale1.Replace("1" + inch + "=", "");
                                Scale1 = Scale1.Substring(0, Scale1.Length - 1);

                            }

                            if (Functions.IsNumeric(Scale1) == true)
                            {
                                _SGEN_mainform.Vw_scale = 1 / Convert.ToDouble(Scale1);
                            }
                        }
                    }



                    double scale2 = _SGEN_mainform.Vw_scale;
                    string nume2 = "1:" + Convert.ToString(1 / scale2);

                    int Colorindex = 256;

                    string anchor = "TL";


                    Autodesk.AutoCAD.EditorInput.PromptKeywordOptions Prompt_string = new Autodesk.AutoCAD.EditorInput.PromptKeywordOptions("");
                    Prompt_string.Message = "\nSpecify ANCHOR:";


                    Prompt_string.Keywords.Add("CEN");

                    Prompt_string.Keywords.Add("TL");
                    Prompt_string.Keywords.Add("TR");
                    Prompt_string.Keywords.Add("BL");
                    Prompt_string.Keywords.Add("BR");


                    Prompt_string.Keywords.Default = "CEN";



                    Prompt_string.AllowNone = true;


                    anchor = "CEN";





                    bool run1 = true;


                    do
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {

                            BlockTable BlockTable1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                            BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                            Autodesk.AutoCAD.EditorInput.PromptPointResult Result_point1 = null;



                            //Alignment_mdi.Jig_show_rectangle_top_left Jig_top_left = new Alignment_mdi.Jig_show_rectangle_top_left();
                            //Result_point1 = Jig_top_left.StartJig(_SGEN_mainform.Vw_scale, _SGEN_mainform.Vw_width, _SGEN_mainform.Vw_height, anchor);


                            Jig_show_rectangle_with_JigPromptPointOptions jig1 = new Jig_show_rectangle_with_JigPromptPointOptions();


                            scale2 = get_current_scale();

                            Result_point1 = jig1.StartJig(scale2, _SGEN_mainform.Vw_width, _SGEN_mainform.Vw_height, anchor, true);



                            if (Result_point1 == null || Result_point1.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel)
                            {
                                run1 = false;
                                Editor1.WriteMessage("\nCommand:");


                            }

                            if (run1 == true)
                            {


                                scale2 = get_current_scale();
                                nume2 = get_current_scale_name();




                                Point3d Point1 = Result_point1.Value;
                                Polyline Rectangle2 = new Autodesk.AutoCAD.DatabaseServices.Polyline();

                                Rectangle2 = create_rectangle_VP(scale2, Point1, anchor, Colorindex);
                                Rectangle2.Layer = _SGEN_mainform.Layer_name_ML_rectangle;

                                BTrecord.AppendEntity(Rectangle2);
                                Trans1.AddNewlyCreatedDBObject(Rectangle2, true);

                                bool add_row = true;


                                if (MessageBox.Show("Is this Ok?", "Plats", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                                {
                                    Rectangle2.Erase();
                                    add_row = false;
                                }


                                if (add_row == true)
                                {

                                    _SGEN_mainform.dt_sheet_index.Rows.Add();

                                    Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_entity;
                                    Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_entity;
                                    Prompt_entity = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the entity containing object data:");
                                    Prompt_entity.SetRejectMessage("\nSelect an entity!");
                                    Prompt_entity.AllowNone = true;
                                    Prompt_entity.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Entity), false);
                                    Rezultat_entity = ThisDrawing.Editor.GetEntity(Prompt_entity);

                                    string Name_of_sheet = "XXX";

                                    if (Rezultat_entity.Status == PromptStatus.OK)
                                    {
                                        #region object data

                                        using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tables1.GetObjectRecords(Convert.ToUInt32(0), Rezultat_entity.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, false))
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
                                                            object valoare1 = Record1[j].StrValue;
                                                            if (Nume_field == odfield1)
                                                            {
                                                                Name_of_sheet = Convert.ToString(valoare1);
                                                                j = Record1.Count;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        #endregion

                                        _SGEN_mainform.dt_sheet_index.Rows[_SGEN_mainform.dt_sheet_index.Rows.Count - 1][_SGEN_mainform.Col_dwg_name] = Name_of_sheet;

                                        Entity ent1 = Trans1.GetObject(Rezultat_entity.ObjectId, OpenMode.ForWrite) as Entity;
                                        ent1.ColorIndex = 30;


                                    }

                                    _SGEN_mainform.dt_sheet_index.Rows[_SGEN_mainform.dt_sheet_index.Rows.Count - 1][Col_handle] = Rectangle2.ObjectId.Handle.Value.ToString();
                                    _SGEN_mainform.dt_sheet_index.Rows[_SGEN_mainform.dt_sheet_index.Rows.Count - 1][_SGEN_mainform.Col_x] = (Rectangle2.GetPoint3dAt(0).X + Rectangle2.GetPoint3dAt(2).X) / 2;
                                    _SGEN_mainform.dt_sheet_index.Rows[_SGEN_mainform.dt_sheet_index.Rows.Count - 1][_SGEN_mainform.Col_y] = (Rectangle2.GetPoint3dAt(0).Y + Rectangle2.GetPoint3dAt(2).Y) / 2;
                                    _SGEN_mainform.dt_sheet_index.Rows[_SGEN_mainform.dt_sheet_index.Rows.Count - 1][_SGEN_mainform.Col_rot] = Functions.GET_Bearing_rad(Rectangle2.GetPoint3dAt(1).X, Rectangle2.GetPoint3dAt(1).Y, Rectangle2.GetPoint3dAt(2).X, Rectangle2.GetPoint3dAt(2).Y) * 180 / Math.PI;
                                    _SGEN_mainform.dt_sheet_index.Rows[_SGEN_mainform.dt_sheet_index.Rows.Count - 1][Col_Width] = Rectangle2.GetPoint3dAt(1).DistanceTo(Rectangle2.GetPoint3dAt(2));
                                    _SGEN_mainform.dt_sheet_index.Rows[_SGEN_mainform.dt_sheet_index.Rows.Count - 1][Col_Height] = Rectangle2.GetPoint3dAt(0).DistanceTo(Rectangle2.GetPoint3dAt(1));
                                    _SGEN_mainform.dt_sheet_index.Rows[_SGEN_mainform.dt_sheet_index.Rows.Count - 1][col_scale] = scale2;
                                    _SGEN_mainform.dt_sheet_index.Rows[_SGEN_mainform.dt_sheet_index.Rows.Count - 1][col_scaleName] = nume2;

                                }

                                Trans1.Commit();

                            }
                        }
                    } while (run1 == true);




                    if (_SGEN_mainform.dt_sheet_index != null)
                    {
                        if (_SGEN_mainform.dt_sheet_index.Rows.Count > 0)
                        {


                            if (System.IO.Directory.Exists(_SGEN_mainform.project_main_folder) == true)
                            {
                                string ProjF = _SGEN_mainform.project_main_folder;
                                if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                                {
                                    ProjF = ProjF + "\\";
                                }

                                string fisier_si = ProjF + _SGEN_mainform.sheet_index_excel_name;

                                forma3.Append_ML_object_data();
                                forma3.populate_dataGridView_sheet_index();


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



        public System.Data.DataTable Creaza_sheet_index_datatable_structure()
        {


            System.Type type_string = typeof(string);
            System.Type type_double = typeof(double);

            List<string> Lista1 = new List<string>();
            List<System.Type> Lista2 = new List<System.Type>();

            Lista1.Add(Col_MMid);
            Lista1.Add(Col_handle);
            Lista1.Add(Col_dwg_name);
            Lista1.Add(Col_M1);
            Lista1.Add(Col_M2);
            Lista1.Add(Col_dispM1);
            Lista1.Add(Col_dispM2);
            Lista1.Add(Col_length);
            Lista1.Add(_SGEN_mainform.Col_x);
            Lista1.Add(_SGEN_mainform.Col_y);
            Lista1.Add(_SGEN_mainform.Col_rot);
            Lista1.Add(Col_Width);
            Lista1.Add(Col_Height);
            Lista1.Add(Col_X1);
            Lista1.Add(Col_Y1);
            Lista1.Add(Col_X2);
            Lista1.Add(Col_Y2);
            Lista1.Add(col_scale);
            Lista1.Add(col_scaleName);

            Lista2.Add(type_string);
            Lista2.Add(type_string);
            Lista2.Add(type_string);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_string);

            System.Data.DataTable dt1 = new System.Data.DataTable();

            for (int i = 0; i < Lista1.Count; ++i)
            {
                dt1.Columns.Add(Lista1[i], Lista2[i]);
            }
            return dt1;
        }

        private Polyline create_rectangle_VP(double scale1, Point3d Point1, string anchor, int cid)
        {



            Polyline poly1 = new Autodesk.AutoCAD.DatabaseServices.Polyline();

            if (anchor == "TL")
            {
                poly1.AddVertexAt(0, new Point2d(Point1.X, Point1.Y - _SGEN_mainform.Vw_height / scale1), 0, 0, 0);
                poly1.AddVertexAt(1, new Point2d(Point1.X, Point1.Y), 0, 0, 0);
                poly1.AddVertexAt(2, new Point2d(Point1.X + _SGEN_mainform.Vw_width / scale1, Point1.Y), 0, 0, 0);
                poly1.AddVertexAt(3, new Point2d(Point1.X + _SGEN_mainform.Vw_width / scale1, Point1.Y - _SGEN_mainform.Vw_height / scale1), 0, 0, 0);
            }

            if (anchor == "TR")
            {
                poly1.AddVertexAt(0, new Point2d(Point1.X - _SGEN_mainform.Vw_width / scale1, Point1.Y - _SGEN_mainform.Vw_height / scale1), 0, 0, 0);
                poly1.AddVertexAt(1, new Point2d(Point1.X - _SGEN_mainform.Vw_width / scale1, Point1.Y), 0, 0, 0);
                poly1.AddVertexAt(2, new Point2d(Point1.X, Point1.Y), 0, 0, 0);
                poly1.AddVertexAt(3, new Point2d(Point1.X, Point1.Y - _SGEN_mainform.Vw_height / scale1), 0, 0, 0);
            }

            if (anchor == "BR")
            {
                poly1.AddVertexAt(0, new Point2d(Point1.X - _SGEN_mainform.Vw_width / scale1, Point1.Y), 0, 0, 0);
                poly1.AddVertexAt(1, new Point2d(Point1.X - _SGEN_mainform.Vw_width / scale1, Point1.Y + _SGEN_mainform.Vw_height / scale1), 0, 0, 0);
                poly1.AddVertexAt(2, new Point2d(Point1.X, Point1.Y + _SGEN_mainform.Vw_height / scale1), 0, 0, 0);
                poly1.AddVertexAt(3, new Point2d(Point1.X, Point1.Y), 0, 0, 0);
            }

            if (anchor == "BL")
            {
                poly1.AddVertexAt(0, new Point2d(Point1.X, Point1.Y), 0, 0, 0);
                poly1.AddVertexAt(1, new Point2d(Point1.X, Point1.Y + _SGEN_mainform.Vw_height / scale1), 0, 0, 0);
                poly1.AddVertexAt(2, new Point2d(Point1.X + _SGEN_mainform.Vw_width / scale1, Point1.Y + _SGEN_mainform.Vw_height / scale1), 0, 0, 0);
                poly1.AddVertexAt(3, new Point2d(Point1.X + _SGEN_mainform.Vw_width / scale1, Point1.Y), 0, 0, 0);

            }

            if (anchor == "CEN")
            {
                poly1.AddVertexAt(0, new Point2d(Point1.X - _SGEN_mainform.Vw_width * 0.5 / scale1, Point1.Y - _SGEN_mainform.Vw_height * 0.5 / scale1), 0, 0, 0);
                poly1.AddVertexAt(1, new Point2d(Point1.X - _SGEN_mainform.Vw_width * 0.5 / scale1, Point1.Y + _SGEN_mainform.Vw_height * 0.5 / scale1), 0, 0, 0);
                poly1.AddVertexAt(2, new Point2d(Point1.X + _SGEN_mainform.Vw_width * 0.5 / scale1, Point1.Y + _SGEN_mainform.Vw_height * 0.5 / scale1), 0, 0, 0);
                poly1.AddVertexAt(3, new Point2d(Point1.X + _SGEN_mainform.Vw_width * 0.5 / scale1, Point1.Y - _SGEN_mainform.Vw_height * 0.5 / scale1), 0, 0, 0);

            }

            poly1.Closed = true;
            poly1.ColorIndex = cid;
            poly1.Elevation = 0;

            return poly1;
        }



        public void delete_mtext_with_OD(string layer_name, string od_table_name)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock lock1 = ThisDrawing.LockDocument())
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.Database.TransactionManager.StartTransaction())
                {
                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                    foreach (ObjectId id1 in BTrecord)
                    {
                        MText ent1 = Trans1.GetObject(id1, OpenMode.ForRead) as MText;
                        if (ent1 != null)
                        {
                            if (ent1.Layer == layer_name)
                            {
                                Autodesk.Gis.Map.ObjectData.Records Records1;
                                bool delete1 = false;
                                if (Tables1.IsTableDefined(od_table_name) == true)
                                {
                                    Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[od_table_name];
                                    using (Records1 = Tabla1.GetObjectTableRecords(Convert.ToUInt32(0), ent1.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, true))
                                    {
                                        if (Records1.Count > 0)
                                        {
                                            Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;
                                            foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                            {
                                                if (delete1 == false)
                                                {
                                                    for (int i = 0; i < Record1.Count; ++i)
                                                    {
                                                        Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[i];
                                                        string Nume_field = Field_def1.Name;
                                                        string Valoare1 = Record1[i].StrValue;
                                                        if (Nume_field == "SegmentName")
                                                        {
                                                            string segment1 = "";

                                                            if (Valoare1 == segment1)
                                                            {
                                                                delete1 = true;
                                                                i = Record1.Count;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    if (delete1 == true)
                                    {
                                        ent1.UpgradeOpen();
                                        ent1.Erase();
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
