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

namespace Alignment_mdi
{
    public partial class AGEN_Owner_Band_Scan : Form
    {

        bool Freeze_operations = false;
        public AGEN_Owner_Band_Scan()
        {
            InitializeComponent();
        }

        private void button_prop_Load_od_click(object sender, EventArgs e)
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
                            comboBox_prop_od_table.Items.Clear();
                            comboBox_prop_linelist_od.Items.Clear();
                            comboBox_prop_owner_od.Items.Clear();


                            Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                            System.Collections.Specialized.StringCollection Nume_tables = new System.Collections.Specialized.StringCollection();
                            Nume_tables = Tables1.GetTableNames();

                            for (int i = 0; i < Nume_tables.Count; i = i + 1)
                            {
                                string Tabla1 = Nume_tables[i];
                                if (comboBox_prop_od_table.Items.Contains(Tabla1) == false)
                                {
                                    comboBox_prop_od_table.Items.Add(Tabla1);
                                }
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


        private void comboBox_prop_od_table_SelectedIndexChanged(object sender, EventArgs e)
        {
            Functions.add_OD_fieds_to_combobox(comboBox_prop_od_table, comboBox_prop_linelist_od);
            Functions.add_OD_fieds_to_combobox(comboBox_prop_od_table, comboBox_prop_owner_od);
        }



        public string get_comboBox_prop_od_table()
        {
            return comboBox_prop_od_table.Text;
        }




        public string get_comboBox_prop_linelist_od()
        {
            return comboBox_prop_linelist_od.Text;
        }

        public string get_comboBox_prop_owner_od()
        {
            return comboBox_prop_owner_od.Text;
        }





        private void button_show_scan_and_draw_ownership_Click(object sender, EventArgs e)
        {
            _AGEN_mainform.tpage_processing.Hide();
            _AGEN_mainform.tpage_blank.Hide();
            _AGEN_mainform.tpage_setup.Hide();
            _AGEN_mainform.tpage_viewport_settings.Hide();
            _AGEN_mainform.tpage_tblk_attrib.Hide();
            _AGEN_mainform.tpage_sheetindex.Hide();
            _AGEN_mainform.tpage_layer_alias.Hide();
            _AGEN_mainform.tpage_crossing_scan.Hide();
            _AGEN_mainform.tpage_crossing_draw.Hide();
            _AGEN_mainform.tpage_profilescan.Hide();
            _AGEN_mainform.tpage_profdraw.Hide();
            _AGEN_mainform.tpage_owner_scan.Hide();

            _AGEN_mainform.tpage_mat.Hide();
            _AGEN_mainform.tpage_cust_scan.Hide();
            _AGEN_mainform.tpage_cust_draw.Hide();
            _AGEN_mainform.tpage_sheet_gen.Hide();


            _AGEN_mainform.tpage_owner_draw.Show();


            _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;



        }

        private void button_scan_properties_Click(object sender, EventArgs e)
        {


            Functions.Kill_excel();

            if (Functions.Get_if_workbook_is_open_in_Excel(_AGEN_mainform.property_excel_name) == true)
            {
                MessageBox.Show("Please close the " + _AGEN_mainform.property_excel_name + " file");
                return;
            }

            string cfg1 = System.IO.Path.GetFileName(_AGEN_mainform.config_path);
            if (Functions.Get_if_workbook_is_open_in_Excel(cfg1) == true)
            {
                MessageBox.Show("Please close the " + cfg1 + " file");
                return;
            }

            string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
            if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
            {
                ProjF = ProjF + "\\";
            }



            if (System.IO.Directory.Exists(ProjF) == true)
            {




                string fisier_cl = ProjF + _AGEN_mainform.cl_excel_name;

                if (System.IO.File.Exists(fisier_cl) == false)
                {
                    Freeze_operations = false;
                    MessageBox.Show("the centerline data file does not exist");
                    _AGEN_mainform.dt_station_equation = null;
                    return;
                }


                if (_AGEN_mainform.dt_centerline == null ||_AGEN_mainform.dt_centerline.Rows.Count == 0)
                {
                    _AGEN_mainform.tpage_setup.Load_centerline_and_station_equation(fisier_cl);
                }

                _AGEN_mainform.Poly3D = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);

                _AGEN_mainform.Poly2D = Functions.Build_2D_CL_from_dt_cl(_AGEN_mainform.dt_centerline);




            }
            else
            {
                Freeze_operations = false;
                MessageBox.Show("the project folder does not exist");
                return;
            }

            string fisier_prop = ProjF + _AGEN_mainform.property_excel_name;





            if (_AGEN_mainform.dt_centerline.Rows.Count == 0)
            {
                Freeze_operations = false;
                MessageBox.Show("the centerline file does not have any data");
                return;
            }

            double poly_length = 0;

            _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;


            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;

            ObjectId[] Empty_array = null;
            Editor1.SetImpliedSelection(Empty_array);
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            _AGEN_mainform.tpage_processing.Show();
            // Ag.WindowState = FormWindowState.Minimized;
            if (Freeze_operations == false)
            {
                try
                {
                    _AGEN_mainform.Data_Table_property = Functions.Creaza_property_datatable_structure();
                    System.Data.DataTable dt_intA = new System.Data.DataTable();

                    dt_intA.Columns.Add(_AGEN_mainform.Col_handle, typeof(string));
                    dt_intA.Columns.Add(_AGEN_mainform.Col_station, typeof(double));
                    dt_intA.Columns.Add("x", typeof(double));
                    dt_intA.Columns.Add("y", typeof(double));

                    System.Data.DataTable dt_intB = new System.Data.DataTable();

                    dt_intB.Columns.Add(_AGEN_mainform.Col_handle, typeof(string));
                    dt_intB.Columns.Add(_AGEN_mainform.Col_station, typeof(double));
                    dt_intB.Columns.Add("x", typeof(double));
                    dt_intB.Columns.Add("y", typeof(double));

                    string Linelist_field = _AGEN_mainform.tpage_owner_scan.get_comboBox_prop_linelist_od();
                    string Owner_field = _AGEN_mainform.tpage_owner_scan.get_comboBox_prop_owner_od();
                    string owner_table_name = _AGEN_mainform.tpage_owner_scan.get_comboBox_prop_od_table();
                    string prop_band_name = _AGEN_mainform.tpage_viewport_settings.get_comboBox_viewport_target_areas(3);

                    int index_property = -1;

                    if (prop_band_name != "")
                    {
                        if (_AGEN_mainform.Data_Table_regular_bands != null)
                        {
                            if (_AGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                            {

                                for (int i = 0; i < _AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                                {
                                    if (_AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] != DBNull.Value)
                                    {
                                        string bn = Convert.ToString(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"]);
                                        if (bn == prop_band_name)
                                        {
                                            index_property = i;
                                            i = _AGEN_mainform.Data_Table_regular_bands.Rows.Count;
                                        }


                                    }
                                }
                            }
                        }
                    }

                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {

                        Freeze_operations = true;

                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable1 = (Autodesk.AutoCAD.DatabaseServices.BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                            BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                            _AGEN_mainform.Poly3D = Trans1.GetObject(_AGEN_mainform.Poly3D.ObjectId, OpenMode.ForWrite) as Polyline3d;
                            poly_length = _AGEN_mainform.Poly3D.Length;

                            Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

                            LayerTable Layer_table = Trans1.GetObject(ThisDrawing.Database.LayerTableId, OpenMode.ForRead) as LayerTable;

                            Autodesk.Gis.Map.ObjectData.Table Tabla0;



                            if (Tables1.IsTableDefined(owner_table_name) == true)
                            {
                                Tabla0 = Tables1[owner_table_name];

                                if (index_property >= 0)
                                {
                                    _AGEN_mainform.Data_Table_regular_bands.Rows[index_property]["OD_table_name"] = owner_table_name;
                                }

                                bool field1_defined = false;
                                bool field2_defined = false;

                                Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla0.FieldDefinitions;
                                for (int i = 0; i < Field_defs1.Count; ++i)
                                {
                                    Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[i];
                                    string Nume_field = Field_def1.Name;

                                    if (Linelist_field != "")
                                    {
                                        if (Nume_field == Linelist_field)
                                        {
                                            if (dt_intA.Columns.Contains(Nume_field) == false)
                                            {
                                                dt_intA.Columns.Add(Nume_field, typeof(string));
                                                dt_intB.Columns.Add(Nume_field, typeof(string));
                                            }

                                            if (index_property >= 0)
                                            {
                                                field1_defined = true;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[index_property]["OD_field1"] = Nume_field;
                                            }

                                        }
                                    }

                                    if (Owner_field != "")
                                    {
                                        if (Nume_field == Owner_field)
                                        {
                                            if (dt_intA.Columns.Contains(Nume_field) == false)
                                            {
                                                dt_intA.Columns.Add(Nume_field, typeof(string));
                                                dt_intB.Columns.Add(Nume_field, typeof(string));
                                            }

                                            if (index_property >= 0)
                                            {
                                                field2_defined = true;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[index_property]["OD_field2"] = Nume_field;
                                            }
                                        }
                                    }

                                }

                                if (index_property >= 0)
                                {
                                    if (field1_defined == false) _AGEN_mainform.Data_Table_regular_bands.Rows[index_property]["OD_field1"] = DBNull.Value;
                                    if (field2_defined == false) _AGEN_mainform.Data_Table_regular_bands.Rows[index_property]["OD_field2"] = DBNull.Value;
                                }

                            }

                            else
                            {
                                MessageBox.Show("There is not such as object data table\r\nplease select the proper data table");

                                _AGEN_mainform.tpage_processing.Hide();
                                _AGEN_mainform.tpage_blank.Hide();
                                _AGEN_mainform.tpage_setup.Hide();
                                _AGEN_mainform.tpage_viewport_settings.Hide();
                                _AGEN_mainform.tpage_tblk_attrib.Hide();
                                _AGEN_mainform.tpage_sheetindex.Hide();
                                _AGEN_mainform.tpage_layer_alias.Hide();
                                _AGEN_mainform.tpage_crossing_scan.Hide();
                                _AGEN_mainform.tpage_crossing_draw.Hide();
                                _AGEN_mainform.tpage_profilescan.Hide();
                                _AGEN_mainform.tpage_profdraw.Hide();
                                _AGEN_mainform.tpage_owner_scan.Hide();

                                _AGEN_mainform.tpage_mat.Hide();
                                _AGEN_mainform.tpage_cust_scan.Hide();
                                _AGEN_mainform.tpage_cust_draw.Hide();
                                _AGEN_mainform.tpage_sheet_gen.Hide();


                                _AGEN_mainform.tpage_owner_draw.Show();

                                Ag.WindowState = FormWindowState.Normal;

                                Freeze_operations = false;
                                return;
                            }

                            // AGEN_mainform.tpage_crossing_scan_and_draw.Build_3d_2d_poly(Trans1, BTrecord, AGEN_mainform.Data_table_centerline, "0");

                            _AGEN_mainform.Poly2D = Functions.Build_2D_CL_from_dt_cl(_AGEN_mainform.dt_centerline);





                            string Layer_cu_property = "";


                            foreach (ObjectId ObjID in BTrecord)
                            {
                                Entity Ent_intersection = Trans1.GetObject(ObjID, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as Entity;

                                if (Ent_intersection != null)
                                {
                                    LayerTableRecord Layer_rec = Trans1.GetObject(Layer_table[Ent_intersection.Layer], OpenMode.ForRead) as LayerTableRecord;

                                    if (Ent_intersection is Polyline && Ent_intersection.ObjectId != _AGEN_mainform.Poly2D.ObjectId && Ent_intersection.ObjectId != _AGEN_mainform.Poly3D.ObjectId && Layer_rec.IsOff == false && Layer_rec.IsFrozen == false)
                                    {
                                        Polyline Poly_int = Ent_intersection as Polyline;

                                        using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tabla0.GetObjectTableRecords(Convert.ToUInt32(0), Ent_intersection.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, false))
                                        {
                                            if (Records1 != null)
                                            {
                                                Layer_cu_property = Ent_intersection.Layer;
                                                goto jump;
                                            }
                                        }
                                    }
                                }
                            }

                        jump:

                            if (Layer_cu_property == "")
                            {

                                MessageBox.Show("The object data table you specified is not attached to any parcel polyline in the drawing\r\nmaybe you shoud map import shp current....");
                                Ag.WindowState = FormWindowState.Normal;
                                _AGEN_mainform.tpage_processing.Hide();
                                Freeze_operations = false;
                                return;
                            }

                            foreach (ObjectId ObjID in BTrecord)
                            {
                                Polyline Poly_int = Trans1.GetObject(ObjID, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as Polyline;

                                if (Poly_int != null)
                                {
                                    if (Poly_int.Layer == Layer_cu_property)
                                    {
                                        Poly_int.UpgradeOpen();

                                        Poly_int.Elevation = _AGEN_mainform.Poly2D.Elevation;

                                        Point3dCollection Col_int = new Point3dCollection();
                                        Col_int = Functions.Intersect_on_both_operands(Poly_int, _AGEN_mainform.Poly2D);



                                        if (Col_int.Count > 0)
                                        {
                                            String Linelist1 = "";
                                            String Owner1 = "";


                                            using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tabla0.GetObjectTableRecords(Convert.ToUInt32(0), Poly_int.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, false))
                                            {
                                                if (Records1 != null)
                                                {
                                                    if (Records1.Count > 0)
                                                    {
                                                        Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla0.FieldDefinitions;

                                                        foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                                        {
                                                            for (int i = 0; i < Record1.Count; ++i)
                                                            {
                                                                Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[i];
                                                                string Nume_field = Field_def1.Name;

                                                                if (Linelist_field != "")
                                                                {
                                                                    if (Nume_field == Linelist_field)
                                                                    {
                                                                        Linelist1 = Record1[i].StrValue;
                                                                    }
                                                                }

                                                                if (Owner_field != "")
                                                                {

                                                                    if (Nume_field == Owner_field)
                                                                    {
                                                                        Owner1 = Record1[i].StrValue;
                                                                    }
                                                                }


                                                            }
                                                        }
                                                    }

                                                }
                                            }

                                            if (Col_int.Count > 0)
                                            {



                                                if (Poly_int.Area > 0.1)
                                                {
                                                    if (Poly_int.Closed == false)
                                                    {
                                                        Poly_int.UpgradeOpen();
                                                        Poly_int.Closed = true;

                                                    }


                                                    if (Poly_int.Closed == true)
                                                    {
                                                        try
                                                        {

                                                            Point3d startpt = new Point3d(_AGEN_mainform.Poly3D.StartPoint.X, _AGEN_mainform.Poly3D.StartPoint.Y, Poly_int.Elevation);
                                                            Point3d endpt = new Point3d(_AGEN_mainform.Poly3D.EndPoint.X, _AGEN_mainform.Poly3D.EndPoint.Y, Poly_int.Elevation);
                                                            double d1 = Poly_int.GetClosestPointTo(startpt, Vector3d.ZAxis, false).DistanceTo(startpt);


                                                            DBObjectCollection Poly_Colection = new DBObjectCollection();
                                                            Poly_Colection.Add(Poly_int);
                                                            DBObjectCollection Region_Colectionft = new DBObjectCollection();
                                                            Region_Colectionft = Autodesk.AutoCAD.DatabaseServices.Region.CreateFromCurves(Poly_Colection);

                                                            Autodesk.AutoCAD.DatabaseServices.Region reg1 = Region_Colectionft[0] as Autodesk.AutoCAD.DatabaseServices.Region;

                                                            Autodesk.AutoCAD.BoundaryRepresentation.PointContainment pc1 = Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Outside;
                                                            Autodesk.AutoCAD.BoundaryRepresentation.PointContainment pc2 = Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Outside;

                                                            using (Autodesk.AutoCAD.BoundaryRepresentation.Brep Brep_obj = new Autodesk.AutoCAD.BoundaryRepresentation.Brep(reg1))
                                                            {
                                                                if (Brep_obj != null)
                                                                {
                                                                    using (Autodesk.AutoCAD.BoundaryRepresentation.BrepEntity ent1 = Brep_obj.GetPointContainment(startpt, out pc1))
                                                                    {
                                                                        if (ent1 is Autodesk.AutoCAD.BoundaryRepresentation.Face)
                                                                        {
                                                                            pc1 = Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Inside;
                                                                        }
                                                                    }
                                                                    using (Autodesk.AutoCAD.BoundaryRepresentation.BrepEntity ent2 = Brep_obj.GetPointContainment(endpt, out pc2))
                                                                    {
                                                                        if (ent2 is Autodesk.AutoCAD.BoundaryRepresentation.Face)
                                                                        {
                                                                            pc2 = Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Inside;
                                                                        }
                                                                    }

                                                                }
                                                            }


                                                            if (pc1 == Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Inside)
                                                            {

                                                                dt_intA.Rows.Add();
                                                                dt_intA.Rows[dt_intA.Rows.Count - 1][_AGEN_mainform.Col_station] = 0;
                                                                dt_intA.Rows[dt_intA.Rows.Count - 1][_AGEN_mainform.Col_handle] = Poly_int.ObjectId.Handle.Value.ToString();
                                                                if (Linelist_field != "") dt_intA.Rows[dt_intA.Rows.Count - 1][Linelist_field] = Linelist1;
                                                                if (Owner_field != "") dt_intA.Rows[dt_intA.Rows.Count - 1][Owner_field] = Owner1;
                                                                dt_intA.Rows[dt_intA.Rows.Count - 1]["x"] = Math.Round(_AGEN_mainform.Poly3D.StartPoint.X, 3);
                                                                dt_intA.Rows[dt_intA.Rows.Count - 1]["y"] = Math.Round(_AGEN_mainform.Poly3D.StartPoint.Y, 3);

                                                                dt_intB.Rows.Add();
                                                                dt_intB.Rows[dt_intB.Rows.Count - 1][_AGEN_mainform.Col_station] = 0;
                                                                dt_intB.Rows[dt_intB.Rows.Count - 1][_AGEN_mainform.Col_handle] = Poly_int.ObjectId.Handle.Value.ToString();
                                                                if (Linelist_field != "") dt_intB.Rows[dt_intB.Rows.Count - 1][Linelist_field] = Linelist1;
                                                                if (Owner_field != "") dt_intB.Rows[dt_intB.Rows.Count - 1][Owner_field] = Owner1;
                                                                dt_intB.Rows[dt_intB.Rows.Count - 1]["x"] = Math.Round(_AGEN_mainform.Poly3D.StartPoint.X, 3);
                                                                dt_intB.Rows[dt_intB.Rows.Count - 1]["y"] = Math.Round(_AGEN_mainform.Poly3D.StartPoint.Y, 3);

                                                            }


                                                            if (pc2 == Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Inside)
                                                            {
                                                                double div1 = 10;

                                                                if (_AGEN_mainform.round1 == 1) div1 = 100;
                                                                if (_AGEN_mainform.round1 == 2) div1 = 1000;
                                                                if (_AGEN_mainform.round1 == 3) div1 = 10000;

                                                                dt_intA.Rows.Add();
                                                                dt_intA.Rows[dt_intA.Rows.Count - 1][_AGEN_mainform.Col_station] = Math.Floor(Math.Round(_AGEN_mainform.Poly3D.Length * div1, _AGEN_mainform.round1 + 1)) / div1;
                                                                dt_intA.Rows[dt_intA.Rows.Count - 1][_AGEN_mainform.Col_handle] = Poly_int.ObjectId.Handle.Value.ToString();
                                                                if (Linelist_field != "") dt_intA.Rows[dt_intA.Rows.Count - 1][Linelist_field] = Linelist1;
                                                                if (Owner_field != "") dt_intA.Rows[dt_intA.Rows.Count - 1][Owner_field] = Owner1;
                                                                dt_intA.Rows[dt_intA.Rows.Count - 1]["x"] = Math.Round(_AGEN_mainform.Poly3D.EndPoint.X, 3);
                                                                dt_intA.Rows[dt_intA.Rows.Count - 1]["y"] = Math.Round(_AGEN_mainform.Poly3D.EndPoint.Y, 3);

                                                                dt_intB.Rows.Add();
                                                                dt_intB.Rows[dt_intB.Rows.Count - 1][_AGEN_mainform.Col_station] = Math.Floor(Math.Round(_AGEN_mainform.Poly3D.Length * div1, _AGEN_mainform.round1 + 1)) / div1;
                                                                dt_intB.Rows[dt_intB.Rows.Count - 1][_AGEN_mainform.Col_handle] = Poly_int.ObjectId.Handle.Value.ToString();
                                                                if (Linelist_field != "") dt_intB.Rows[dt_intB.Rows.Count - 1][Linelist_field] = Linelist1;
                                                                if (Owner_field != "") dt_intB.Rows[dt_intB.Rows.Count - 1][Owner_field] = Owner1;
                                                                dt_intB.Rows[dt_intB.Rows.Count - 1]["x"] = Math.Round(_AGEN_mainform.Poly3D.EndPoint.X, 3);
                                                                dt_intB.Rows[dt_intB.Rows.Count - 1]["y"] = Math.Round(_AGEN_mainform.Poly3D.EndPoint.Y, 3);


                                                            }

                                                        }
                                                        catch (Autodesk.AutoCAD.Runtime.Exception ex)
                                                        {

                                                        }
                                                    }

                                                }
                                            }


                                            if (Col_int.Count > 0)
                                            {
                                                for (int index = 0; index < Col_int.Count; ++index)
                                                {



                                                    Point3d Point_on_poly2d = new Point3d();
                                                    Point3d Point_on_poly = new Point3d();
                                                    double Station_grid = 0;

                                                    Point_on_poly2d = _AGEN_mainform.Poly2D.GetClosestPointTo(Col_int[index], Vector3d.ZAxis, true);
                                                    double Param2d = _AGEN_mainform.Poly2D.GetParameterAtPoint(Point_on_poly2d);

                                                    if (Math.Round(Param2d, 3) < _AGEN_mainform.Poly2D.EndParam)
                                                    {
                                                        Station_grid = _AGEN_mainform.Poly3D.GetDistanceAtParameter(Param2d);
                                                        Point_on_poly = _AGEN_mainform.Poly3D.GetPointAtDist(Station_grid);
                                                    }
                                                    else
                                                    {

                                                        Station_grid = Math.Round(_AGEN_mainform.Poly3D.Length, 3) - 0.001;
                                                        Point_on_poly = _AGEN_mainform.Poly3D.EndPoint;
                                                    }


                                                    dt_intA.Rows.Add();
                                                    dt_intA.Rows[dt_intA.Rows.Count - 1][_AGEN_mainform.Col_station] = Station_grid;
                                                    dt_intA.Rows[dt_intA.Rows.Count - 1][_AGEN_mainform.Col_handle] = Poly_int.ObjectId.Handle.Value.ToString();
                                                    if (Linelist_field != "") dt_intA.Rows[dt_intA.Rows.Count - 1][Linelist_field] = Linelist1;
                                                    if (Owner_field != "") dt_intA.Rows[dt_intA.Rows.Count - 1][Owner_field] = Owner1;
                                                    dt_intA.Rows[dt_intA.Rows.Count - 1]["x"] = Math.Round(Point_on_poly.X, 3);
                                                    dt_intA.Rows[dt_intA.Rows.Count - 1]["y"] = Math.Round(Point_on_poly.Y, 3);

                                                    dt_intB.Rows.Add();
                                                    dt_intB.Rows[dt_intB.Rows.Count - 1][_AGEN_mainform.Col_station] = Station_grid;
                                                    dt_intB.Rows[dt_intB.Rows.Count - 1][_AGEN_mainform.Col_handle] = Poly_int.ObjectId.Handle.Value.ToString();
                                                    if (Linelist_field != "") dt_intB.Rows[dt_intB.Rows.Count - 1][Linelist_field] = Linelist1;
                                                    if (Owner_field != "") dt_intB.Rows[dt_intB.Rows.Count - 1][Owner_field] = Owner1;
                                                    dt_intB.Rows[dt_intB.Rows.Count - 1]["x"] = Math.Round(Point_on_poly.X, 3);
                                                    dt_intB.Rows[dt_intB.Rows.Count - 1]["y"] = Math.Round(Point_on_poly.Y, 3);

                                                }
                                            }
                                        }


                                    }
                                }
                            }

                            if (dt_intA.Rows.Count == 0)
                            {
                                MessageBox.Show("the data does not intersect the centerline");
                                Freeze_operations = false;
                                return;
                            }

                            dt_intA = Functions.Sort_data_table(dt_intA, _AGEN_mainform.Col_station);
                            dt_intA = Functions.Elimina_duplicates_from_data_table(dt_intA);

                            dt_intB = Functions.Sort_data_table(dt_intB, _AGEN_mainform.Col_station);
                            dt_intB = Functions.Elimina_duplicates_from_data_table(dt_intB);

                            int count1 = 0;

                            do
                            {
                                count1 = count1 + 1;
                                double Sta11 = Convert.ToDouble(dt_intA.Rows[0][_AGEN_mainform.Col_station]);

                                string Handle11 = dt_intA.Rows[0][_AGEN_mainform.Col_handle].ToString();
                                string Linelist11 = "";
                                if (Linelist_field != "") Linelist11 = dt_intA.Rows[0][Linelist_field].ToString();
                                string Owner11 = "";
                                if (Owner_field != "") Owner11 = dt_intA.Rows[0][Owner_field].ToString();
                                double x0 = Convert.ToDouble(dt_intA.Rows[0]["x"]);
                                double y0 = Convert.ToDouble(dt_intA.Rows[0]["y"]);

                                double Sta12 = Convert.ToDouble(dt_intA.Rows[1][_AGEN_mainform.Col_station]);
                                string Handle12 = dt_intA.Rows[1][_AGEN_mainform.Col_handle].ToString();

                                double Sta13 = Convert.ToDouble(dt_intA.Rows[1][_AGEN_mainform.Col_station]);
                                string Handle13 = dt_intA.Rows[1][_AGEN_mainform.Col_handle].ToString();
                                double x1 = Convert.ToDouble(dt_intA.Rows[1]["x"]);
                                double y1 = Convert.ToDouble(dt_intA.Rows[1]["y"]);

                                if (dt_intA.Rows.Count >= 3)
                                {
                                    if (dt_intA.Rows[2][_AGEN_mainform.Col_station] != DBNull.Value) Sta13 = Convert.ToDouble(dt_intA.Rows[2][_AGEN_mainform.Col_station]);

                                    if (dt_intA.Rows[2][_AGEN_mainform.Col_handle] != DBNull.Value) Handle13 = dt_intA.Rows[2][_AGEN_mainform.Col_handle].ToString();
                                }

                                bool linie_added = false;

                                if (Handle11 == Handle12)
                                {
                                    _AGEN_mainform.Data_Table_property.Rows.Add();
                                    linie_added = true;
                                }
                                else if (dt_intA.Rows.Count >= 3)
                                {
                                    if (Handle11 == Handle13)
                                    {
                                        _AGEN_mainform.Data_Table_property.Rows.Add();
                                        linie_added = true;
                                    }
                                }

                                if (linie_added == true)
                                {
                                    if (Linelist_field != "") _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1][_AGEN_mainform.Col_Linelist] = Linelist11;
                                    if (Owner_field != "") _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1][_AGEN_mainform.Col_Owner] = Owner11;

                                    if (_AGEN_mainform.tpage_sheetindex.get_radioButton_use3D_stations() == false)
                                    {
                                        _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1][_AGEN_mainform.Col_2DSta1] = Sta11;
                                    }
                                    else
                                    {
                                        _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1][_AGEN_mainform.Col_3DSta1] = Sta11;
                                    }
                                    _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1]["X_Beg"] = x0;
                                    _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1]["Y_Beg"] = y0;
                                }


                                if (linie_added == true)
                                {
                                    if (Math.Round(Sta11, 3) != Math.Round(Sta12, 3) || Math.Round(Sta11, 3) != Math.Round(Sta13, 3))
                                    {
                                        if (Handle11 == Handle12)
                                        {
                                            if (_AGEN_mainform.tpage_sheetindex.get_radioButton_use3D_stations() == false)
                                            {
                                                _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1][_AGEN_mainform.Col_2DSta2] = Sta12;
                                            }
                                            else
                                            {
                                                _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1][_AGEN_mainform.Col_3DSta2] = Sta12;
                                            }
                                            _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1]["X_End"] = x1;
                                            _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1]["Y_End"] = y1;

                                            dt_intA.Rows.RemoveAt(1);
                                            dt_intA.Rows.RemoveAt(0);
                                        }
                                        else if (Handle11 == Handle13)
                                        {
                                            if (_AGEN_mainform.tpage_sheetindex.get_radioButton_use3D_stations() == false)
                                            {
                                                _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1][_AGEN_mainform.Col_2DSta2] = Sta13;
                                            }
                                            else
                                            {
                                                _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1][_AGEN_mainform.Col_3DSta2] = Sta13;
                                            }
                                            _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1]["X_End"] = x1;
                                            _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1]["Y_End"] = y1;

                                            dt_intA.Rows.RemoveAt(2);
                                            dt_intA.Rows.RemoveAt(0);
                                        }

                                        else
                                        {
                                            MessageBox.Show("you have an error on parcel around the stations around \r\n" + Functions.Get_String_Rounded(Sta11, 0) + "\r\n" + Functions.Get_String_Rounded(Sta12, 0) + "\r\n" + Functions.Get_String_Rounded(Sta13, 0));
                                            ThisDrawing.Editor.WriteMessage(Functions.Get_String_Rounded(Sta11, 0) + "\r\n" + Functions.Get_String_Rounded(Sta12, 0) + "\r\n" + Functions.Get_String_Rounded(Sta13, 0));

                                            Ag.WindowState = FormWindowState.Normal;
                                            _AGEN_mainform.tpage_processing.Hide();
                                            Freeze_operations = false;
                                            return;
                                        }
                                    }
                                }
                                else
                                {
                                    dt_intA.Rows.RemoveAt(0);
                                }
                            } while (dt_intA.Rows.Count >= 2 && count1 < 50);



                            if (count1 == 50)
                            {
                                _AGEN_mainform.Data_Table_property = Functions.Creaza_property_datatable_structure();

                                for (int i = 0; i < dt_intB.Rows.Count; ++i)
                                {
                                    double Sta1 = Convert.ToDouble(dt_intB.Rows[i][_AGEN_mainform.Col_station]);
                                    string Handle1 = dt_intB.Rows[i][_AGEN_mainform.Col_handle].ToString();
                                    string Linelist1 = "";
                                    if (Linelist_field != "") Linelist1 = dt_intB.Rows[i][Linelist_field].ToString();
                                    string Owner1 = "";
                                    if (Owner_field != "") Owner1 = dt_intB.Rows[i][Owner_field].ToString();
                                    double x0 = Convert.ToDouble(dt_intB.Rows[i]["x"]);
                                    double y0 = Convert.ToDouble(dt_intB.Rows[i]["y"]);

                                    if (i < dt_intB.Rows.Count - 3)
                                    {
                                        for (int j = i + 1; j < i + 4; ++j)
                                        {
                                            double Sta2 = Convert.ToDouble(dt_intB.Rows[j][_AGEN_mainform.Col_station]);
                                            string Handle2 = dt_intB.Rows[j][_AGEN_mainform.Col_handle].ToString();
                                            string Linelist2 = "";
                                            if (Linelist_field != "") Linelist2 = dt_intB.Rows[j][Linelist_field].ToString();
                                            string Owner2 = "";
                                            if (Owner_field != "") Owner2 = dt_intB.Rows[j][Owner_field].ToString();
                                            double x1 = Convert.ToDouble(dt_intB.Rows[j]["x"]);
                                            double y1 = Convert.ToDouble(dt_intB.Rows[j]["y"]);
                                            if (Math.Round(Sta1, 0) != Math.Round(Sta2, 0) && Handle1 == Handle2)
                                            {
                                                _AGEN_mainform.Data_Table_property.Rows.Add();
                                                if (Linelist_field != "") _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1][_AGEN_mainform.Col_Linelist] = Linelist1;
                                                if (Owner_field != "") _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1][_AGEN_mainform.Col_Owner] = Owner1;

                                                if (_AGEN_mainform.tpage_sheetindex.get_radioButton_use3D_stations() == false)
                                                {
                                                    _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1][_AGEN_mainform.Col_2DSta1] = Sta1;
                                                    _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1][_AGEN_mainform.Col_2DSta2] = Sta2;
                                                }
                                                else
                                                {
                                                    _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1][_AGEN_mainform.Col_3DSta1] = Sta1;
                                                    _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1][_AGEN_mainform.Col_3DSta2] = Sta2;
                                                }
                                                _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1]["X_Beg"] = x0;
                                                _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1]["Y_Beg"] = y0;
                                                _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1]["X_End"] = x1;
                                                _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1]["Y_End"] = y1;
                                                j = i + 4;
                                            }
                                        }
                                    }
                                }

                                _AGEN_mainform.Data_Table_property.Rows.Add();
                                if (Linelist_field != "") _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1][_AGEN_mainform.Col_Linelist] = dt_intB.Rows[dt_intB.Rows.Count - 1][Linelist_field];
                                if (Owner_field != "") _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1][_AGEN_mainform.Col_Owner] = dt_intB.Rows[dt_intB.Rows.Count - 1][Owner_field];

                                if (_AGEN_mainform.tpage_sheetindex.get_radioButton_use3D_stations() == false)
                                {
                                    _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1][_AGEN_mainform.Col_2DSta1] = _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 2][_AGEN_mainform.Col_2DSta2];
                                    _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1][_AGEN_mainform.Col_2DSta2] = dt_intB.Rows[dt_intB.Rows.Count - 1][_AGEN_mainform.Col_station];
                                }
                                else
                                {
                                    _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1][_AGEN_mainform.Col_3DSta1] = _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 2][_AGEN_mainform.Col_2DSta2];
                                    _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1][_AGEN_mainform.Col_3DSta2] = dt_intB.Rows[dt_intB.Rows.Count - 1][_AGEN_mainform.Col_station];
                                }

                                _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1]["X_Beg"] = _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 2]["X_End"];
                                _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1]["Y_Beg"] = _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 2]["Y_End"];
                                _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1]["X_End"] = dt_intB.Rows[dt_intB.Rows.Count - 1]["x"];
                                _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1]["Y_End"] = dt_intB.Rows[dt_intB.Rows.Count - 1]["y"];


                            }


                            if (_AGEN_mainform.dt_station_equation != null && _AGEN_mainform.dt_station_equation.Rows.Count > 0)
                            {
                                if (_AGEN_mainform.dt_station_equation.Columns.Contains("measured") == false)
                                {
                                    _AGEN_mainform.dt_station_equation.Columns.Add("measured", typeof(double));
                                }
                                for (int i = 0; i < _AGEN_mainform.dt_station_equation.Rows.Count; ++i)
                                {
                                    if (_AGEN_mainform.dt_station_equation.Rows[i]["Reroute End X"] != DBNull.Value && _AGEN_mainform.dt_station_equation.Rows[i]["Reroute End Y"] != DBNull.Value)
                                    {
                                        double x = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i]["Reroute End X"]);
                                        double y = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i]["Reroute End Y"]);


                                        Point3d pt_on_2d = _AGEN_mainform.Poly2D.GetClosestPointTo(new Point3d(x, y, 0), Vector3d.ZAxis, false);
                                        double param1 = _AGEN_mainform.Poly2D.GetParameterAtPoint(pt_on_2d);
                                        double eq_meas = _AGEN_mainform.Poly3D.GetDistanceAtParameter(param1);
                                        _AGEN_mainform.dt_station_equation.Rows[i]["measured"] = eq_meas;

                                    }
                                }
                            }



                            if (_AGEN_mainform.Data_Table_property != null)
                            {
                                if (_AGEN_mainform.Data_Table_property.Rows.Count > 0)
                                {
                                    for (int i = 0; i < _AGEN_mainform.Data_Table_property.Rows.Count; ++i)
                                    {
                                        if (_AGEN_mainform.tpage_sheetindex.get_radioButton_use3D_stations() == false)
                                        {
                                            if (_AGEN_mainform.Data_Table_property.Rows[i][1] != DBNull.Value)
                                            {
                                                double St1 = Math.Round(Convert.ToDouble(_AGEN_mainform.Data_Table_property.Rows[i][1]), _AGEN_mainform.round1);
                                                double div1 = 10;
                                                if (_AGEN_mainform.round1 == 1) div1 = 100;
                                                if (_AGEN_mainform.round1 == 2) div1 = 1000;
                                                if (_AGEN_mainform.round1 == 3) div1 = 10000;

                                                if (St1 >= poly_length) St1 = Math.Floor(Math.Round(poly_length * div1, _AGEN_mainform.round1 + 1)) / div1;
                                                _AGEN_mainform.Data_Table_property.Rows[i][1] = St1;
                                                _AGEN_mainform.Data_Table_property.Rows[i][2] = DBNull.Value;
                                                if (_AGEN_mainform.dt_station_equation != null)
                                                {
                                                    if (_AGEN_mainform.dt_station_equation.Rows.Count > 0)
                                                    {
                                                        _AGEN_mainform.Data_Table_property.Rows[i][5] = Math.Round(Functions.Station_equation_ofV2(St1, _AGEN_mainform.dt_station_equation), _AGEN_mainform.round1);
                                                    }
                                                }
                                            }

                                            if (_AGEN_mainform.Data_Table_property.Rows[i][3] != DBNull.Value)
                                            {
                                                double St2 = Math.Round(Convert.ToDouble(_AGEN_mainform.Data_Table_property.Rows[i][3]), _AGEN_mainform.round1);
                                                double div1 = 10;
                                                if (_AGEN_mainform.round1 == 1) div1 = 100;
                                                if (_AGEN_mainform.round1 == 2) div1 = 1000;
                                                if (_AGEN_mainform.round1 == 3) div1 = 10000;
                                                if (St2 >= poly_length) St2 = Math.Floor(Math.Round(poly_length * div1, _AGEN_mainform.round1 + 1)) / div1;
                                                _AGEN_mainform.Data_Table_property.Rows[i][3] = St2;
                                                _AGEN_mainform.Data_Table_property.Rows[i][4] = DBNull.Value;
                                                if (_AGEN_mainform.dt_station_equation != null)
                                                {
                                                    if (_AGEN_mainform.dt_station_equation.Rows.Count > 0)
                                                    {
                                                        _AGEN_mainform.Data_Table_property.Rows[i][6] = Math.Round(Functions.Station_equation_ofV2(St2, _AGEN_mainform.dt_station_equation), _AGEN_mainform.round1);
                                                    }
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (_AGEN_mainform.Data_Table_property.Rows[i][2] != DBNull.Value)
                                            {
                                                double St1 = Math.Round(Convert.ToDouble(_AGEN_mainform.Data_Table_property.Rows[i][2]), _AGEN_mainform.round1);
                                                double div1 = 10;
                                                if (_AGEN_mainform.round1 == 1) div1 = 100;
                                                if (_AGEN_mainform.round1 == 2) div1 = 1000;
                                                if (_AGEN_mainform.round1 == 3) div1 = 10000;
                                                if (St1 >= poly_length) St1 = Math.Floor(Math.Round(poly_length * div1, _AGEN_mainform.round1 + 1)) / div1;
                                                _AGEN_mainform.Data_Table_property.Rows[i][2] = St1;
                                                _AGEN_mainform.Data_Table_property.Rows[i][1] = DBNull.Value;
                                                if (_AGEN_mainform.dt_station_equation != null)
                                                {
                                                    if (_AGEN_mainform.dt_station_equation.Rows.Count > 0)
                                                    {
                                                        _AGEN_mainform.Data_Table_property.Rows[i][5] = Math.Round(Functions.Station_equation_ofV2(St1, _AGEN_mainform.dt_station_equation), _AGEN_mainform.round1);
                                                    }
                                                }
                                            }
                                            if (_AGEN_mainform.Data_Table_property.Rows[i][4] != DBNull.Value)
                                            {
                                                double St2 = Math.Round(Convert.ToDouble(_AGEN_mainform.Data_Table_property.Rows[i][4]), _AGEN_mainform.round1);
                                                double div1 = 10;
                                                if (_AGEN_mainform.round1 == 1) div1 = 100;
                                                if (_AGEN_mainform.round1 == 2) div1 = 1000;
                                                if (_AGEN_mainform.round1 == 3) div1 = 10000;

                                                if (St2 >= poly_length) St2 = Math.Floor(Math.Round(poly_length * div1, _AGEN_mainform.round1 + 1)) / div1;
                                                _AGEN_mainform.Data_Table_property.Rows[i][4] = St2;
                                                _AGEN_mainform.Data_Table_property.Rows[i][3] = DBNull.Value;
                                                if (_AGEN_mainform.dt_station_equation != null)
                                                {
                                                    if (_AGEN_mainform.dt_station_equation.Rows.Count > 0)
                                                    {
                                                        _AGEN_mainform.Data_Table_property.Rows[i][6] = Math.Round(Functions.Station_equation_ofV2(St2, _AGEN_mainform.dt_station_equation), _AGEN_mainform.round1);
                                                    }
                                                }
                                            }
                                        }
                                    }

                                    for (int i = _AGEN_mainform.Data_Table_property.Rows.Count - 1; i >= 0; --i)
                                    {
                                        if (_AGEN_mainform.tpage_sheetindex.get_radioButton_use3D_stations() == false)
                                        {

                                            if (_AGEN_mainform.Data_Table_property.Rows[i][1] != DBNull.Value && _AGEN_mainform.Data_Table_property.Rows[i][3] != DBNull.Value)
                                            {
                                                double sta1 = Convert.ToDouble(_AGEN_mainform.Data_Table_property.Rows[i][1]);
                                                double sta2 = Convert.ToDouble(_AGEN_mainform.Data_Table_property.Rows[i][3]);
                                                string owner1 = "";
                                                string linelist1 = "";
                                                if (_AGEN_mainform.Data_Table_property.Rows[i][8] != DBNull.Value)
                                                {
                                                    owner1 = Convert.ToString(_AGEN_mainform.Data_Table_property.Rows[i][8]);
                                                }
                                                if (_AGEN_mainform.Data_Table_property.Rows[i][9] != DBNull.Value)
                                                {
                                                    linelist1 = Convert.ToString(_AGEN_mainform.Data_Table_property.Rows[i][9]);
                                                }
                                                if (i > 0)
                                                {
                                                    double sta11 = Convert.ToDouble(_AGEN_mainform.Data_Table_property.Rows[i - 1][1]);
                                                    double sta22 = Convert.ToDouble(_AGEN_mainform.Data_Table_property.Rows[i - 1][3]);
                                                    if (sta1 == sta11 && sta2 == sta22)
                                                    {
                                                        string owner11 = "";
                                                        string linelist11 = "";
                                                        string owner11_prev = "";
                                                        string linelist11_prev = "";
                                                        string owner11_next = "";
                                                        string linelist11_next = "";

                                                        if (_AGEN_mainform.Data_Table_property.Rows[i - 1][8] != DBNull.Value)
                                                        {
                                                            owner11 = Convert.ToString(_AGEN_mainform.Data_Table_property.Rows[i - 1][8]);
                                                        }
                                                        if (_AGEN_mainform.Data_Table_property.Rows[i - 1][9] != DBNull.Value)
                                                        {
                                                            linelist11 = Convert.ToString(_AGEN_mainform.Data_Table_property.Rows[i - 1][9]);
                                                        }
                                                        if (i < _AGEN_mainform.Data_Table_property.Rows.Count - 1)
                                                        {
                                                            if (_AGEN_mainform.Data_Table_property.Rows[i + 1][8] != DBNull.Value)
                                                            {
                                                                owner11_next = Convert.ToString(_AGEN_mainform.Data_Table_property.Rows[i + 1][8]);
                                                            }
                                                            if (_AGEN_mainform.Data_Table_property.Rows[i + 1][9] != DBNull.Value)
                                                            {
                                                                linelist11_next = Convert.ToString(_AGEN_mainform.Data_Table_property.Rows[i + 1][9]);
                                                            }
                                                        }

                                                        if (i > 1)
                                                        {
                                                            if (_AGEN_mainform.Data_Table_property.Rows[i - 2][8] != DBNull.Value)
                                                            {
                                                                owner11_prev = Convert.ToString(_AGEN_mainform.Data_Table_property.Rows[i - 2][8]);
                                                            }
                                                            if (_AGEN_mainform.Data_Table_property.Rows[i - 2][9] != DBNull.Value)
                                                            {
                                                                linelist11_prev = Convert.ToString(_AGEN_mainform.Data_Table_property.Rows[i - 2][9]);
                                                            }


                                                            if (linelist11_next == linelist1 && owner11_next == owner1)
                                                            {
                                                                _AGEN_mainform.Data_Table_property.Rows[i].Delete();
                                                            }
                                                            else if (linelist11_next == linelist11_prev && owner11_next == owner11_prev)
                                                            {
                                                                _AGEN_mainform.Data_Table_property.Rows[i - 1].Delete();

                                                            }
                                                            else if (linelist1 == linelist11_prev && owner1 == owner11_prev)
                                                            {
                                                                _AGEN_mainform.Data_Table_property.Rows[i - 1].Delete();

                                                            }
                                                        }
                                                    }

                                                }
                                            }

                                        }
                                        else
                                        {
                                            if (_AGEN_mainform.Data_Table_property.Rows[i][2] != DBNull.Value && _AGEN_mainform.Data_Table_property.Rows[i][4] != DBNull.Value)
                                            {
                                                double sta1 = Convert.ToDouble(_AGEN_mainform.Data_Table_property.Rows[i][2]);
                                                double sta2 = Convert.ToDouble(_AGEN_mainform.Data_Table_property.Rows[i][4]);
                                                string owner1 = "";
                                                string linelist1 = "";
                                                if (_AGEN_mainform.Data_Table_property.Rows[i][8] != DBNull.Value)
                                                {
                                                    owner1 = Convert.ToString(_AGEN_mainform.Data_Table_property.Rows[i][8]);
                                                }
                                                if (_AGEN_mainform.Data_Table_property.Rows[i][9] != DBNull.Value)
                                                {
                                                    linelist1 = Convert.ToString(_AGEN_mainform.Data_Table_property.Rows[i][9]);
                                                }
                                                if (i > 0)
                                                {
                                                    double sta11 = Convert.ToDouble(_AGEN_mainform.Data_Table_property.Rows[i - 1][2]);
                                                    double sta22 = Convert.ToDouble(_AGEN_mainform.Data_Table_property.Rows[i - 1][4]);
                                                    if (sta1 == sta11 && sta2 == sta22)
                                                    {
                                                        string owner11 = "";
                                                        string linelist11 = "";
                                                        string owner11_prev = "";
                                                        string linelist11_prev = "";
                                                        string owner11_next = "";
                                                        string linelist11_next = "";

                                                        if (_AGEN_mainform.Data_Table_property.Rows[i - 1][8] != DBNull.Value)
                                                        {
                                                            owner11 = Convert.ToString(_AGEN_mainform.Data_Table_property.Rows[i - 1][8]);
                                                        }
                                                        if (_AGEN_mainform.Data_Table_property.Rows[i - 1][9] != DBNull.Value)
                                                        {
                                                            linelist11 = Convert.ToString(_AGEN_mainform.Data_Table_property.Rows[i - 1][9]);
                                                        }
                                                        if (i < _AGEN_mainform.Data_Table_property.Rows.Count - 1)
                                                        {
                                                            if (_AGEN_mainform.Data_Table_property.Rows[i + 1][8] != DBNull.Value)
                                                            {
                                                                owner11_next = Convert.ToString(_AGEN_mainform.Data_Table_property.Rows[i + 1][8]);
                                                            }
                                                            if (_AGEN_mainform.Data_Table_property.Rows[i + 1][9] != DBNull.Value)
                                                            {
                                                                linelist11_next = Convert.ToString(_AGEN_mainform.Data_Table_property.Rows[i + 1][9]);
                                                            }
                                                        }

                                                        if (i > 1)
                                                        {
                                                            if (_AGEN_mainform.Data_Table_property.Rows[i - 2][8] != DBNull.Value)
                                                            {
                                                                owner11_prev = Convert.ToString(_AGEN_mainform.Data_Table_property.Rows[i - 2][8]);
                                                            }
                                                            if (_AGEN_mainform.Data_Table_property.Rows[i - 2][9] != DBNull.Value)
                                                            {
                                                                linelist11_prev = Convert.ToString(_AGEN_mainform.Data_Table_property.Rows[i - 2][9]);
                                                            }


                                                            if (linelist11_next == linelist1 && owner11_next == owner1)
                                                            {
                                                                _AGEN_mainform.Data_Table_property.Rows[i].Delete();
                                                            }
                                                            else if (linelist11_next == linelist11_prev && owner11_next == owner11_prev)
                                                            {
                                                                _AGEN_mainform.Data_Table_property.Rows[i - 1].Delete();

                                                            }
                                                            else if (linelist1 == linelist11_prev && owner1 == owner11_prev)
                                                            {
                                                                _AGEN_mainform.Data_Table_property.Rows[i - 1].Delete();

                                                            }
                                                        }
                                                    }

                                                }
                                            }
                                        }
                                    }

                                    bool arata_mesaj = false;

                                    for (int i = 0; i < _AGEN_mainform.Data_Table_property.Rows.Count - 1; ++i)
                                    {
                                        if (_AGEN_mainform.tpage_sheetindex.get_radioButton_use3D_stations() == false)
                                        {
                                            if (_AGEN_mainform.Data_Table_property.Rows[i][1] != DBNull.Value && _AGEN_mainform.Data_Table_property.Rows[i][3] != DBNull.Value)
                                            {
                                                double sta1 = Convert.ToDouble(_AGEN_mainform.Data_Table_property.Rows[i][1]);
                                                double sta2 = Convert.ToDouble(_AGEN_mainform.Data_Table_property.Rows[i][3]);
                                                _AGEN_mainform.Data_Table_property.Rows[i][9] = sta2 - sta1;

                                                if (_AGEN_mainform.Data_Table_property.Rows[i + 1][1] != DBNull.Value)
                                                {
                                                    double sta11 = Convert.ToDouble(_AGEN_mainform.Data_Table_property.Rows[i + 1][1]);
                                                    if (sta2 != sta11)
                                                    {
                                                        arata_mesaj = true;
                                                    }
                                                }
                                                if (sta1 >= sta2)
                                                {
                                                    arata_mesaj = true;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (_AGEN_mainform.Data_Table_property.Rows[i][2] != DBNull.Value && _AGEN_mainform.Data_Table_property.Rows[i][4] != DBNull.Value)
                                            {
                                                double sta1 = Convert.ToDouble(_AGEN_mainform.Data_Table_property.Rows[i][2]);
                                                double sta2 = Convert.ToDouble(_AGEN_mainform.Data_Table_property.Rows[i][4]);
                                                _AGEN_mainform.Data_Table_property.Rows[i][9] = sta2 - sta1;
                                                if (_AGEN_mainform.Data_Table_property.Rows[i + 1][2] != DBNull.Value)
                                                {
                                                    double sta11 = Convert.ToDouble(_AGEN_mainform.Data_Table_property.Rows[i + 1][2]);
                                                    if (sta2 != sta11)
                                                    {
                                                        arata_mesaj = true;
                                                    }
                                                }
                                                if (sta1 >= sta2)
                                                {
                                                    arata_mesaj = true;
                                                }
                                            }
                                        }
                                    }

                                    Populate_property_file_with_settings(fisier_prop, _AGEN_mainform.config_path);

                                    if (arata_mesaj == true)
                                    {
                                        MessageBox.Show("there are gaps in data.... \r\n please review the measured stations columns");
                                    }
                                }
                            }

                            _AGEN_mainform.Poly3D.Erase();
                            Trans1.Commit();

                        }
                    }




                    ThisDrawing.Editor.WriteMessage("\nCommand:");
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                _AGEN_mainform.tpage_processing.Hide();
                Freeze_operations = false;



            }

            Ag.WindowState = FormWindowState.Normal;

        }


        public void Populate_property_file(string file_prop)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application Excel1 = null;
               
                try
                {
                    Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                }
                catch (System.Exception ex)
                {
                    Excel1 = new Microsoft.Office.Interop.Excel.Application();
                   
                }

                if (Excel1.Workbooks.Count==0) Excel1.Visible = _AGEN_mainform.ExcelVisible;

                Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;

                if (System.IO.File.Exists(file_prop) == false)
                {
                    Workbook1 = Excel1.Workbooks.Add();
                }
                else
                {
                    Workbook1 = Excel1.Workbooks.Open(file_prop);
                }
                Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];

                try
                {
                    string segment1 = _AGEN_mainform.tpage_setup.Get_segment_name1();
                    if (segment1 == "not defined") segment1 = "";
                    Functions.Transfer_to_worksheet_Data_table(W1, _AGEN_mainform.Data_Table_property, _AGEN_mainform.Start_row_property, "General");
                    Functions.Create_header_property_file(W1, _AGEN_mainform.tpage_setup.Get_client_name(), _AGEN_mainform.tpage_setup.Get_project_name(), segment1);

                    if (System.IO.File.Exists(file_prop) == false)
                    {
                        Workbook1.SaveAs(file_prop);
                    }
                    else
                    {
                        Workbook1.Save();
                    }
                    Workbook1.Close();

                    if (Excel1.Workbooks.Count==0)
                    {
                        Excel1.Quit();
                    }
                    else
                    {
                        Excel1.Visible = true;
                    }
                }
                catch (System.Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);

                }
                finally
                {
                    if (W1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W1);
                    if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                    if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }

        }



        public void Populate_property_file_with_settings(string file_prop, string cfg1)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application Excel1 = null;
               
                try
                {
                    Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                }
                catch (System.Exception ex)
                {
                    Excel1 = new Microsoft.Office.Interop.Excel.Application();
                   
                }

                if (Excel1.Workbooks.Count==0) Excel1.Visible = _AGEN_mainform.ExcelVisible;

                Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;

                if (System.IO.File.Exists(file_prop) == false)
                {
                    Workbook1 = Excel1.Workbooks.Add();
                }
                else
                {
                    Workbook1 = Excel1.Workbooks.Open(file_prop);
                }
                Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];
                Microsoft.Office.Interop.Excel.Workbook Workbook2 = Excel1.Workbooks.Open(cfg1);
                Microsoft.Office.Interop.Excel.Worksheet W2 = null;

                string segment1 = _AGEN_mainform.current_segment;
                if (segment1 == "not defined") segment1 = "";

                foreach (Microsoft.Office.Interop.Excel.Worksheet wsh2 in Workbook2.Worksheets)
                {
                    if (wsh2.Name == "Ownership_data_config_" + segment1)
                    {
                        W2 = wsh2;
                    }
                }

                if (W2 == null)
                {
                    W2 = Workbook2.Worksheets.Add(System.Reflection.Missing.Value, Workbook2.Worksheets[Workbook2.Worksheets.Count], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                    W2.Name = "Ownership_data_config_" + segment1;
                }

                try
                {

                    Functions.Transfer_to_worksheet_Data_table(W1, _AGEN_mainform.Data_Table_property, _AGEN_mainform.Start_row_property, "General");
                    Functions.Create_header_property_file(W1, _AGEN_mainform.tpage_setup.Get_client_name(), _AGEN_mainform.tpage_setup.Get_project_name(), segment1);

                    if (System.IO.File.Exists(file_prop) == false)
                    {
                        Workbook1.SaveAs(file_prop);
                    }
                    else
                    {
                        Workbook1.Save();
                    }
                    Workbook1.Close();


                    int NrR = 9;
                    int NrC = 2;



                    Object[,] values = new object[NrR, NrC];
                    values[0, 0] = "Ownership Block Name";
                    values[1, 0] = "Station start Attribute";
                    values[2, 0] = "Station end Attribute";
                    values[3, 0] = "Length Attribute";
                    values[4, 0] = "Linelist Attribute";
                    values[5, 0] = "Ownership Attribute";


                    values[6, 0] = "Ownership Data Table";
                    values[6, 1] = get_comboBox_prop_od_table();
                    values[7, 0] = "Owner field";
                    values[7, 1] = get_comboBox_prop_owner_od();
                    values[8, 0] = "Linelist(Tract) Field";
                    values[8, 1] = get_comboBox_prop_linelist_od();


                    Microsoft.Office.Interop.Excel.Range range2 = W2.Range["A1:B9"];
                    range2.Cells.NumberFormat = "General";
                    range2.Value2 = values;
                    Functions.Color_border_range_inside(range2, 0);

                    Workbook2.Save();
                    Workbook2.Close();




                    if (Excel1.Workbooks.Count==0)
                    {
                        Excel1.Quit();
                    }
                    else
                    {
                        Excel1.Visible = true;
                    }
                }
                catch (System.Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);

                }
                finally
                {
                    if (W1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W1);
                    if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                    if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }

        }

        

        public void set_comboBox_prop_od_table(string string1)
        {
            if (comboBox_prop_od_table.Items.Contains(string1) == false)
            {
                comboBox_prop_od_table.Items.Add(string1);
            }
            comboBox_prop_od_table.SelectedIndex = comboBox_prop_od_table.Items.IndexOf(string1);
        }

        public void set_comboBox_prop_owner_od(string string1)
        {
            if (comboBox_prop_owner_od.Items.Contains(string1) == false)
            {
                comboBox_prop_owner_od.Items.Add(string1);
            }
            comboBox_prop_owner_od.SelectedIndex = comboBox_prop_owner_od.Items.IndexOf(string1);
        }

        public void set_comboBox_prop_linelist_od(string string1)
        {
            if (comboBox_prop_linelist_od.Items.Contains(string1) == false)
            {
                comboBox_prop_linelist_od.Items.Add(string1);
            }
            comboBox_prop_linelist_od.SelectedIndex = comboBox_prop_linelist_od.Items.IndexOf(string1);
        }

        public void clear_combobox()
        {
            comboBox_prop_od_table.Items.Clear();
            comboBox_prop_linelist_od.Items.Clear();
            comboBox_prop_owner_od.Items.Clear();
        }





        private void button_open_ownership_xlxs_Click(object sender, EventArgs e)
        {
            if (Freeze_operations == false)
            {
                Freeze_operations = true;
                try
                {
                    string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                    if (System.IO.Directory.Exists(ProjF) == true)
                    {

                        if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                        {
                            ProjF = ProjF + "\\";
                        }

                        string fisier_prop = ProjF + _AGEN_mainform.property_excel_name;

                        if (System.IO.File.Exists(fisier_prop) == false)
                        {
                            Freeze_operations = false;
                            MessageBox.Show("the ownership data file does not exist");
                            return;
                        }

                        Microsoft.Office.Interop.Excel.Application Excel1 = null;

                        try
                        {
                            Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                        }
                        catch (System.Exception ex)
                        {
                            Excel1 = new Microsoft.Office.Interop.Excel.Application();

                        }
                        Excel1.Visible = true;
                        Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(fisier_prop);
                    }
                    else
                    {
                        _AGEN_mainform.tpage_processing.Hide();

                        MessageBox.Show("the project folder does not exist");
                    }



                }
                catch (System.Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);

                }
                Freeze_operations = false;

            }
        }

        public void Fill_combobox_segments()
        {
            comboBox_segment_name.Items.Clear();
            if (_AGEN_mainform.lista_segments != null && _AGEN_mainform.lista_segments.Count > 0)
            {
                try
                {
                    for (int i = 0; i < _AGEN_mainform.lista_segments.Count; ++i)
                    {
                        comboBox_segment_name.Items.Add(_AGEN_mainform.lista_segments[i]);
                    }
                    comboBox_segment_name.SelectedIndex = 0;
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        public void set_combobox_segment_name()
        {
            comboBox_segment_name.SelectedIndex = comboBox_segment_name.Items.IndexOf(_AGEN_mainform.current_segment);
        }

        private void ComboBox_segment_name_SelectedIndexChanged(object sender, EventArgs e)
        {
            _AGEN_mainform.current_segment = comboBox_segment_name.Text;
            _AGEN_mainform.tpage_setup.set_combobox_segment_name();


        }
    }
}
