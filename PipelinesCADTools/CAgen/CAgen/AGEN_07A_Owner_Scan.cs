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


        public AGEN_Owner_Band_Scan()
        {
            InitializeComponent();
        }

        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_open_ownership_xlxs);
            lista_butoane.Add(button_scan_properties);
            lista_butoane.Add(button_show_scan_and_draw_ownership);
            lista_butoane.Add(comboBox_segment_name);
            lista_butoane.Add(comboBox_prop_od_table);
            lista_butoane.Add(comboBox_prop_linelist_od);
            lista_butoane.Add(comboBox_prop_owner_od);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {

                bt1.Enabled = false;

            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_open_ownership_xlxs);
            lista_butoane.Add(button_scan_properties);
            lista_butoane.Add(button_show_scan_and_draw_ownership);
            lista_butoane.Add(comboBox_segment_name);
            lista_butoane.Add(comboBox_prop_od_table);
            lista_butoane.Add(comboBox_prop_linelist_od);
            lista_butoane.Add(comboBox_prop_owner_od);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }

        private void button_prop_Load_od_click(object sender, EventArgs e)
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
            set_enable_true();

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

        private void button_scan_properties_old()
        {


            

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

            set_enable_false();

            if (System.IO.Directory.Exists(ProjF) == true)
            {




                string fisier_cl = ProjF + _AGEN_mainform.cl_excel_name;

                if (System.IO.File.Exists(fisier_cl) == false)
                {
                    set_enable_true();
                    MessageBox.Show("the centerline data file does not exist");
                    _AGEN_mainform.dt_station_equation = null;
                    return;
                }


                if (_AGEN_mainform.dt_centerline == null || _AGEN_mainform.dt_centerline.Rows.Count == 0)
                {
                    _AGEN_mainform.tpage_setup.Load_centerline_and_station_equation(fisier_cl);
                }






            }
            else
            {
                set_enable_true();
                MessageBox.Show("the project folder does not exist");
                return;
            }

            string fisier_prop = ProjF + _AGEN_mainform.property_excel_name;


            if (_AGEN_mainform.dt_centerline.Rows.Count == 0)
            {
                set_enable_true();
                MessageBox.Show("the centerline file does not have any data");
                return;
            }

            Functions.create_backup(fisier_prop);
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

            Polyline poly2d = null;
            Polyline3d poly3d = null;
            set_enable_false();

            System.Data.DataTable dt1 = null;


            string Col_2DSta1 = "2DStaBeg";
            string Col_3DSta1 = "3DStaBeg";
            string Col_2DSta2 = "2DStaEnd";
            string Col_3DSta2 = "3DStaEnd";
            string Col_EqSta1 = "EqStaBeg";
            string Col_EqSta2 = "EqStaEnd";
            string Col_Owner = "Owner";
            string Col_Linelist = "ParcelId";

            string Col_X1 = "X_Beg";
            string Col_Y1 = "Y_Beg";
            string Col_X2 = "X_End";
            string Col_Y2 = "Y_End";
            try
            {
                dt1 = Functions.Creaza_property_datatable_structure();

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
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable1 = (Autodesk.AutoCAD.DatabaseServices.BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                        poly2d = Functions.Build_2D_CL_from_dt_cl(_AGEN_mainform.dt_centerline);
                        poly_length = poly2d.Length;
                        if (_AGEN_mainform.Project_type == "3D")
                        {
                            poly3d = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);
                            poly_length = poly3d.Length;
                        }
                        string Layer_cu_property = "";

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
                            set_enable_true();
                            return;
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


                                    Point3d pt_on_2d = poly2d.GetClosestPointTo(new Point3d(x, y, 0), Vector3d.ZAxis, false);
                                    double eq_meas = poly2d.GetDistAtPoint(pt_on_2d);

                                    if (_AGEN_mainform.Project_type == "3D")
                                    {
                                        double param1 = poly2d.GetParameterAtPoint(pt_on_2d);
                                        eq_meas = poly3d.GetDistanceAtParameter(param1);

                                    }
                                    _AGEN_mainform.dt_station_equation.Rows[i]["measured"] = eq_meas;

                                }
                            }
                        }

                        foreach (ObjectId ObjID in BTrecord)
                        {
                            Entity Ent_intersection = Trans1.GetObject(ObjID, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as Entity;
                            if (Ent_intersection != null)
                            {
                                LayerTableRecord Layer_rec = Trans1.GetObject(Layer_table[Ent_intersection.Layer], OpenMode.ForRead) as LayerTableRecord;
                                if (Ent_intersection is Polyline && Layer_rec.IsOff == false && Layer_rec.IsFrozen == false)
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
                            set_enable_true();
                            return;
                        }

                        foreach (ObjectId ObjID in BTrecord)
                        {
                            Polyline Poly_int = Trans1.GetObject(ObjID, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as Polyline;

                            if (Poly_int != null)
                            {
                                if (Poly_int.Layer == Layer_cu_property)
                                {

                                    poly2d.Elevation = Poly_int.Elevation;

                                    Point3d pt_start = poly2d.StartPoint;
                                    Point3d pt_end = poly2d.EndPoint;

                                    Point3dCollection Col_int = Functions.Intersect_on_both_operands(Poly_int, poly2d);

                                    if (Col_int.Count > 0)
                                    {
                                        string Linelist1 = "";
                                        string Owner1 = "";

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
                                            if (Col_int.Count == 1)
                                            {
                                                dt1.Rows.Add();
                                                if (Owner1 != "") dt1.Rows[dt1.Rows.Count - 1][Col_Owner] = Owner1;
                                                if (Linelist1 != "") dt1.Rows[dt1.Rows.Count - 1][Col_Linelist] = Linelist1;

                                                double Station_grid = -1;
                                                if (_AGEN_mainform.Project_type == "2D")
                                                {
                                                    Point3d Point_on_poly = poly2d.GetClosestPointTo(Col_int[0], Vector3d.ZAxis, true);
                                                    Station_grid = poly2d.GetDistAtPoint(Point_on_poly);

                                                    if (Station_grid < poly_length - Station_grid)
                                                    {
                                                        dt1.Rows[dt1.Rows.Count - 1][Col_2DSta2] = Math.Round(Station_grid, _AGEN_mainform.round1);
                                                    }
                                                    else
                                                    {
                                                        dt1.Rows[dt1.Rows.Count - 1][Col_2DSta1] = Math.Round(Station_grid, _AGEN_mainform.round1);

                                                    }

                                                }
                                                else
                                                {
                                                    Point3d Point_on_poly = poly2d.GetClosestPointTo(Col_int[0], Vector3d.ZAxis, true);
                                                    double Param2 = poly2d.GetParameterAtPoint(Point_on_poly);
                                                    Station_grid = poly3d.GetDistanceAtParameter(Param2);
                                                    if (Station_grid < poly_length - Station_grid)
                                                    {
                                                        dt1.Rows[dt1.Rows.Count - 1][Col_3DSta2] = Math.Round(Station_grid, _AGEN_mainform.round1);
                                                    }
                                                    else
                                                    {
                                                        dt1.Rows[dt1.Rows.Count - 1][Col_3DSta1] = Math.Round(Station_grid, _AGEN_mainform.round1);
                                                    }

                                                }

                                                if (Station_grid < poly_length - Station_grid)
                                                {
                                                    if (Station_grid != -1 && _AGEN_mainform.dt_station_equation != null && _AGEN_mainform.dt_station_equation.Rows.Count > 0) dt1.Rows[dt1.Rows.Count - 1][Col_EqSta2] = Math.Round(Functions.Station_equation_ofV2(Station_grid, _AGEN_mainform.dt_station_equation), _AGEN_mainform.round1);
                                                    dt1.Rows[dt1.Rows.Count - 1][Col_X2] = Col_int[0].X;
                                                    dt1.Rows[dt1.Rows.Count - 1][Col_Y2] = Col_int[0].Y;
                                                }
                                                else
                                                {
                                                    if (Station_grid != -1 && _AGEN_mainform.dt_station_equation != null && _AGEN_mainform.dt_station_equation.Rows.Count > 0) dt1.Rows[dt1.Rows.Count - 1][Col_EqSta1] = Math.Round(Functions.Station_equation_ofV2(Station_grid, _AGEN_mainform.dt_station_equation), _AGEN_mainform.round1);
                                                    dt1.Rows[dt1.Rows.Count - 1][Col_X1] = Col_int[0].X;
                                                    dt1.Rows[dt1.Rows.Count - 1][Col_Y1] = Col_int[0].Y;
                                                }

                                                #region brep    
                                                if (Poly_int.Closed == true && Math.Abs(Station_grid - poly_length) > 0.4999)
                                                {
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
                                                            using (Autodesk.AutoCAD.BoundaryRepresentation.BrepEntity ent1 = Brep_obj.GetPointContainment(pt_start, out pc1))
                                                            {
                                                                if (ent1 is Autodesk.AutoCAD.BoundaryRepresentation.Face)
                                                                {
                                                                    pc1 = Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Inside;
                                                                }
                                                            }
                                                            using (Autodesk.AutoCAD.BoundaryRepresentation.BrepEntity ent2 = Brep_obj.GetPointContainment(pt_end, out pc2))
                                                            {
                                                                if (ent2 is Autodesk.AutoCAD.BoundaryRepresentation.Face)
                                                                {
                                                                    pc2 = Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Inside;
                                                                }
                                                            }

                                                            if (pc1 == Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Inside)
                                                            {

                                                                if (_AGEN_mainform.dt_station_equation != null && _AGEN_mainform.dt_station_equation.Rows.Count > 0) dt1.Rows[dt1.Rows.Count - 1][Col_EqSta1] = Math.Round(Functions.Station_equation_ofV2(0, _AGEN_mainform.dt_station_equation), _AGEN_mainform.round1);
                                                                dt1.Rows[dt1.Rows.Count - 1][Col_X1] = pt_start.X;
                                                                dt1.Rows[dt1.Rows.Count - 1][Col_Y1] = pt_start.Y;

                                                                if (_AGEN_mainform.Project_type == "2D")
                                                                {
                                                                    dt1.Rows[dt1.Rows.Count - 1][Col_2DSta1] = 0;
                                                                }
                                                                else
                                                                {
                                                                    dt1.Rows[dt1.Rows.Count - 1][Col_3DSta1] = 0;
                                                                }
                                                            }

                                                            if (pc2 == Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Inside)
                                                            {

                                                                if (_AGEN_mainform.dt_station_equation != null && _AGEN_mainform.dt_station_equation.Rows.Count > 0) dt1.Rows[dt1.Rows.Count - 1][Col_EqSta2] = Math.Round(Functions.Station_equation_ofV2(poly_length, _AGEN_mainform.dt_station_equation), _AGEN_mainform.round1);
                                                                dt1.Rows[dt1.Rows.Count - 1][Col_X2] = pt_end.X;
                                                                dt1.Rows[dt1.Rows.Count - 1][Col_Y2] = pt_end.Y;

                                                                if (_AGEN_mainform.Project_type == "2D")
                                                                {
                                                                    dt1.Rows[dt1.Rows.Count - 1][Col_2DSta2] = Math.Round(poly_length, _AGEN_mainform.round1);
                                                                }
                                                                else
                                                                {
                                                                    dt1.Rows[dt1.Rows.Count - 1][Col_3DSta2] = Math.Round(poly_length, _AGEN_mainform.round1);
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                #endregion


                                            }

                                            else if (Col_int.Count == 2)
                                            {
                                                dt1.Rows.Add();
                                                if (Owner1 != "") dt1.Rows[dt1.Rows.Count - 1][Col_Owner] = Owner1;
                                                if (Linelist1 != "") dt1.Rows[dt1.Rows.Count - 1][Col_Linelist] = Linelist1;

                                                double x1, y1, x2, y2, sta1, sta2;
                                                x1 = Col_int[0].X;
                                                y1 = Col_int[0].Y;
                                                x2 = Col_int[1].X;
                                                y2 = Col_int[1].Y;

                                                if (_AGEN_mainform.Project_type == "2D")
                                                {
                                                    Point3d Point_on_poly1 = poly2d.GetClosestPointTo(new Point3d(x1, y1, poly2d.Elevation), Vector3d.ZAxis, true);
                                                    sta1 = poly2d.GetDistAtPoint(Point_on_poly1);
                                                    Point3d Point_on_poly2 = poly2d.GetClosestPointTo(new Point3d(x2, y2, poly2d.Elevation), Vector3d.ZAxis, true);
                                                    sta2 = poly2d.GetDistAtPoint(Point_on_poly2);
                                                }
                                                else
                                                {
                                                    Point3d Point_on_poly1 = poly2d.GetClosestPointTo(new Point3d(x1, y1, poly2d.Elevation), Vector3d.ZAxis, true);
                                                    double param1 = poly2d.GetParameterAtPoint(Point_on_poly1);
                                                    sta1 = poly3d.GetDistanceAtParameter(param1);

                                                    Point3d Point_on_poly2 = poly2d.GetClosestPointTo(new Point3d(x2, y2, poly2d.Elevation), Vector3d.ZAxis, true);
                                                    double param2 = poly2d.GetParameterAtPoint(Point_on_poly2);
                                                    sta2 = poly3d.GetDistanceAtParameter(param2);
                                                }

                                                if (sta1 > sta2)
                                                {
                                                    double t = sta2;
                                                    sta2 = sta1;
                                                    sta1 = t;
                                                    t = x2;
                                                    x2 = x1;
                                                    x1 = t;
                                                    t = y2;
                                                    y2 = y1;
                                                    y1 = t;
                                                }

                                                if (_AGEN_mainform.Project_type == "2D")
                                                {
                                                    dt1.Rows[dt1.Rows.Count - 1][Col_2DSta1] = Math.Round(sta1, _AGEN_mainform.round1);
                                                    dt1.Rows[dt1.Rows.Count - 1][Col_2DSta2] = Math.Round(sta2, _AGEN_mainform.round1);
                                                }
                                                else
                                                {
                                                    dt1.Rows[dt1.Rows.Count - 1][Col_3DSta1] = Math.Round(sta1, _AGEN_mainform.round1);
                                                    dt1.Rows[dt1.Rows.Count - 1][Col_3DSta2] = Math.Round(sta2, _AGEN_mainform.round1);
                                                }

                                                dt1.Rows[dt1.Rows.Count - 1][Col_X1] = x1;
                                                dt1.Rows[dt1.Rows.Count - 1][Col_Y1] = y1;
                                                dt1.Rows[dt1.Rows.Count - 1][Col_X2] = x2;
                                                dt1.Rows[dt1.Rows.Count - 1][Col_Y2] = y2;

                                                if (_AGEN_mainform.dt_station_equation != null && _AGEN_mainform.dt_station_equation.Rows.Count > 0) dt1.Rows[dt1.Rows.Count - 1][Col_EqSta1] = Math.Round(Functions.Station_equation_ofV2(sta1, _AGEN_mainform.dt_station_equation), _AGEN_mainform.round1);
                                                if (_AGEN_mainform.dt_station_equation != null && _AGEN_mainform.dt_station_equation.Rows.Count > 0) dt1.Rows[dt1.Rows.Count - 1][Col_EqSta2] = Math.Round(Functions.Station_equation_ofV2(sta2, _AGEN_mainform.dt_station_equation), _AGEN_mainform.round1);


                                            }
                                            else
                                            {
                                                for (int k = 0; k < Col_int.Count; k += 2)
                                                {
                                                    dt1.Rows.Add();
                                                    if (Owner1 != "") dt1.Rows[dt1.Rows.Count - 1][Col_Owner] = Owner1;
                                                    if (Linelist1 != "") dt1.Rows[dt1.Rows.Count - 1][Col_Linelist] = Linelist1;

                                                    double x1, y1, x2, y2, sta1, sta2;
                                                    x1 = Col_int[k].X;
                                                    y1 = Col_int[k].Y;

                                                    if (k + 1 < Col_int.Count)
                                                    {
                                                        x2 = Col_int[k + 1].X;
                                                        y2 = Col_int[k + 1].Y;
                                                    }
                                                    else
                                                    {
                                                        x2 = -1;
                                                        y2 = -1;
                                                    }


                                                    if (_AGEN_mainform.Project_type == "2D")
                                                    {
                                                        Point3d Point_on_poly1 = poly2d.GetClosestPointTo(new Point3d(x1, y1, poly2d.Elevation), Vector3d.ZAxis, true);
                                                        sta1 = poly2d.GetDistAtPoint(Point_on_poly1);
                                                        if (x2 != -1)
                                                        {
                                                            Point3d Point_on_poly2 = poly2d.GetClosestPointTo(new Point3d(x2, y2, poly2d.Elevation), Vector3d.ZAxis, true);
                                                            sta2 = poly2d.GetDistAtPoint(Point_on_poly2);
                                                        }
                                                        else
                                                        {
                                                            sta2 = -1;
                                                        }

                                                    }
                                                    else
                                                    {
                                                        Point3d Point_on_poly1 = poly2d.GetClosestPointTo(new Point3d(x1, y1, poly2d.Elevation), Vector3d.ZAxis, true);
                                                        double param1 = poly2d.GetParameterAtPoint(Point_on_poly1);
                                                        sta1 = poly3d.GetDistanceAtParameter(param1);

                                                        if (x2 != -1)
                                                        {
                                                            Point3d Point_on_poly2 = poly2d.GetClosestPointTo(new Point3d(x2, y2, poly2d.Elevation), Vector3d.ZAxis, true);
                                                            double param2 = poly2d.GetParameterAtPoint(Point_on_poly2);
                                                            sta2 = poly3d.GetDistanceAtParameter(param2);
                                                        }
                                                        else
                                                        {
                                                            sta2 = -1;
                                                        }
                                                    }

                                                    if (x2 != -1)
                                                    {
                                                        if (sta1 > sta2)
                                                        {
                                                            double t = sta2;
                                                            sta2 = sta1;
                                                            sta1 = t;
                                                            t = x2;
                                                            x2 = x1;
                                                            x1 = t;
                                                            t = y2;
                                                            y2 = y1;
                                                            y1 = t;
                                                        }

                                                        if (_AGEN_mainform.Project_type == "2D")
                                                        {
                                                            dt1.Rows[dt1.Rows.Count - 1][Col_2DSta1] = Math.Round(sta1, _AGEN_mainform.round1);
                                                            dt1.Rows[dt1.Rows.Count - 1][Col_2DSta2] = Math.Round(sta2, _AGEN_mainform.round1);
                                                        }
                                                        else
                                                        {
                                                            dt1.Rows[dt1.Rows.Count - 1][Col_3DSta1] = Math.Round(sta1, _AGEN_mainform.round1);
                                                            dt1.Rows[dt1.Rows.Count - 1][Col_3DSta2] = Math.Round(sta2, _AGEN_mainform.round1);
                                                        }


                                                        dt1.Rows[dt1.Rows.Count - 1][Col_X1] = x1;
                                                        dt1.Rows[dt1.Rows.Count - 1][Col_Y1] = y1;
                                                        dt1.Rows[dt1.Rows.Count - 1][Col_X2] = x2;
                                                        dt1.Rows[dt1.Rows.Count - 1][Col_Y2] = y2;

                                                        if (_AGEN_mainform.dt_station_equation != null && _AGEN_mainform.dt_station_equation.Rows.Count > 0) dt1.Rows[dt1.Rows.Count - 1][Col_EqSta1] = Math.Round(Functions.Station_equation_ofV2(sta1, _AGEN_mainform.dt_station_equation), _AGEN_mainform.round1);
                                                        if (_AGEN_mainform.dt_station_equation != null && _AGEN_mainform.dt_station_equation.Rows.Count > 0) dt1.Rows[dt1.Rows.Count - 1][Col_EqSta2] = Math.Round(Functions.Station_equation_ofV2(sta2, _AGEN_mainform.dt_station_equation), _AGEN_mainform.round1);
                                                    }
                                                    else if(x2==-1)
                                                    {
                                                        if (sta1 < poly_length - sta1)
                                                        {
                                                            dt1.Rows[dt1.Rows.Count - 1][Col_X2] = x1;
                                                            dt1.Rows[dt1.Rows.Count - 1][Col_Y2] = y1;
                                                            if (_AGEN_mainform.Project_type == "2D")
                                                            {
                                                                dt1.Rows[dt1.Rows.Count - 1][Col_2DSta2] = Math.Round(sta1, _AGEN_mainform.round1);
                                                            }
                                                            else
                                                            {
                                                                dt1.Rows[dt1.Rows.Count - 1][Col_3DSta2] = Math.Round(sta1, _AGEN_mainform.round1);
                                                            }
                                                            if (_AGEN_mainform.dt_station_equation != null && _AGEN_mainform.dt_station_equation.Rows.Count > 0) dt1.Rows[dt1.Rows.Count - 1][Col_EqSta2] = Math.Round(Functions.Station_equation_ofV2(sta1, _AGEN_mainform.dt_station_equation), _AGEN_mainform.round1);

                                                        }
                                                        else
                                                        {
                                                            dt1.Rows[dt1.Rows.Count - 1][Col_X1] = x1;
                                                            dt1.Rows[dt1.Rows.Count - 1][Col_Y1] = y1;
                                                            if (_AGEN_mainform.Project_type == "2D")
                                                            {
                                                                dt1.Rows[dt1.Rows.Count - 1][Col_2DSta1] = Math.Round(sta1, _AGEN_mainform.round1);
                                                            }
                                                            else
                                                            {
                                                                dt1.Rows[dt1.Rows.Count - 1][Col_3DSta1] = Math.Round(sta1, _AGEN_mainform.round1);
                                                            }
                                                            if (_AGEN_mainform.dt_station_equation != null && _AGEN_mainform.dt_station_equation.Rows.Count > 0) dt1.Rows[dt1.Rows.Count - 1][Col_EqSta1] = Math.Round(Functions.Station_equation_ofV2(sta1, _AGEN_mainform.dt_station_equation), _AGEN_mainform.round1);

                                                        }


                                                        #region brep    
                                                        if (Poly_int.Closed == true && Math.Abs(sta1 - poly_length) > 0.4999)
                                                        {
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
                                                                    using (Autodesk.AutoCAD.BoundaryRepresentation.BrepEntity ent1 = Brep_obj.GetPointContainment(pt_start, out pc1))
                                                                    {
                                                                        if (ent1 is Autodesk.AutoCAD.BoundaryRepresentation.Face)
                                                                        {
                                                                            pc1 = Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Inside;
                                                                        }
                                                                    }
                                                                    using (Autodesk.AutoCAD.BoundaryRepresentation.BrepEntity ent2 = Brep_obj.GetPointContainment(pt_end, out pc2))
                                                                    {
                                                                        if (ent2 is Autodesk.AutoCAD.BoundaryRepresentation.Face)
                                                                        {
                                                                            pc2 = Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Inside;
                                                                        }
                                                                    }

                                                                    if (pc1 == Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Inside)
                                                                    {

                                                                        if (_AGEN_mainform.dt_station_equation != null && _AGEN_mainform.dt_station_equation.Rows.Count > 0) dt1.Rows[dt1.Rows.Count - 1][Col_EqSta1] = Math.Round(Functions.Station_equation_ofV2(0, _AGEN_mainform.dt_station_equation), _AGEN_mainform.round1);
                                                                        dt1.Rows[dt1.Rows.Count - 1][Col_X1] = pt_start.X;
                                                                        dt1.Rows[dt1.Rows.Count - 1][Col_Y1] = pt_start.Y;

                                                                        if (_AGEN_mainform.Project_type == "2D")
                                                                        {
                                                                            dt1.Rows[dt1.Rows.Count - 1][Col_2DSta1] = 0;
                                                                        }
                                                                        else
                                                                        {
                                                                            dt1.Rows[dt1.Rows.Count - 1][Col_3DSta1] = 0;
                                                                        }
                                                                    }

                                                                    if (pc2 == Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Inside)
                                                                    {

                                                                        if (_AGEN_mainform.dt_station_equation != null && _AGEN_mainform.dt_station_equation.Rows.Count > 0) dt1.Rows[dt1.Rows.Count - 1][Col_EqSta2] = Math.Round(Functions.Station_equation_ofV2(poly_length, _AGEN_mainform.dt_station_equation), _AGEN_mainform.round1);
                                                                        dt1.Rows[dt1.Rows.Count - 1][Col_X2] = pt_end.X;
                                                                        dt1.Rows[dt1.Rows.Count - 1][Col_Y2] = pt_end.Y;

                                                                        if (_AGEN_mainform.Project_type == "2D")
                                                                        {
                                                                            dt1.Rows[dt1.Rows.Count - 1][Col_2DSta2] = Math.Round(poly_length, _AGEN_mainform.round1);
                                                                        }
                                                                        else
                                                                        {
                                                                            dt1.Rows[dt1.Rows.Count - 1][Col_3DSta2] = Math.Round(poly_length, _AGEN_mainform.round1);
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                        #endregion
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }





                        string col_sta = Col_2DSta2;

                        if (_AGEN_mainform.Project_type == "2D")
                        {
                            _AGEN_mainform.Data_Table_property = Functions.Sort_data_table(dt1, Col_2DSta1);

                      }
                        else
                        {
                            _AGEN_mainform.Data_Table_property = Functions.Sort_data_table(dt1, Col_3DSta1);
                            col_sta = Col_3DSta2;
                        }

                        if (_AGEN_mainform.Data_Table_property != null && _AGEN_mainform.Data_Table_property.Rows.Count > 0 && _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1][col_sta] != DBNull.Value)
                        {
                            double sta_end = Convert.ToDouble(_AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1][col_sta]);
                            if (sta_end > poly_length)
                            {
                                if (poly_length < Math.Round(poly_length, _AGEN_mainform.round1))
                                {
                                    double new_end = Math.Round(poly_length, _AGEN_mainform.round1 + 1);
                                    if (new_end > poly_length)
                                    {
                                        if (_AGEN_mainform.round1 == 0)
                                        {
                                            new_end = new_end - 0.1;
                                        }
                                        else if (_AGEN_mainform.round1 == 1)
                                        {
                                            new_end = new_end - 0.01;
                                        }
                                        else if (_AGEN_mainform.round1 == 2)
                                        {
                                            new_end = new_end - 0.02;
                                        }
                                        else
                                        {
                                            new_end = new_end - 1;
                                        }
                                    }
                                    _AGEN_mainform.Data_Table_property.Rows[_AGEN_mainform.Data_Table_property.Rows.Count - 1][col_sta] = new_end;
                                }

                            }
                        }


                        if (_AGEN_mainform.Data_Table_property != null)
                        {
                            if (_AGEN_mainform.Data_Table_property.Rows.Count > 0)
                            {
                                Populate_property_file(fisier_prop);

                                Populate_settings_with_control_data(_AGEN_mainform.config_path);
                            }
                        }

                        if (_AGEN_mainform.Project_type == "3D")
                        {
                            if (poly3d != null && poly3d.IsErased == false)
                            {
                                poly3d.UpgradeOpen();
                                poly3d.Erase();
                            }
                        }

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
            set_enable_true();





            Ag.WindowState = FormWindowState.Normal;

        }

        private void button_scan_properties_Click(object sender, EventArgs e)
        {


            

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

            set_enable_false();

            if (System.IO.Directory.Exists(ProjF) == true)
            {




                string fisier_cl = ProjF + _AGEN_mainform.cl_excel_name;

                if (System.IO.File.Exists(fisier_cl) == false)
                {
                    set_enable_true();
                    MessageBox.Show("the centerline data file does not exist");
                    _AGEN_mainform.dt_station_equation = null;
                    return;
                }


                if (_AGEN_mainform.dt_centerline == null || _AGEN_mainform.dt_centerline.Rows.Count == 0)
                {
                    _AGEN_mainform.tpage_setup.Load_centerline_and_station_equation(fisier_cl);
                }
            }
            else
            {
                set_enable_true();
                MessageBox.Show("the project folder does not exist");
                return;
            }

            string fisier_prop = ProjF + _AGEN_mainform.property_excel_name;

            if (_AGEN_mainform.dt_centerline.Rows.Count == 0)
            {
                set_enable_true();
                MessageBox.Show("the centerline file does not have any data");
                return;
            }

            Functions.create_backup(fisier_prop);
          

            _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;

            ObjectId[] Empty_array = null;
            Editor1.SetImpliedSelection(Empty_array);
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            _AGEN_mainform.tpage_processing.Show();
            // Ag.WindowState = FormWindowState.Minimized;

            set_enable_false();

            try
            {

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
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable1 = (Autodesk.AutoCAD.DatabaseServices.BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                      

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
                            set_enable_true();
                            return;
                        }

                        string Layer_cu_property = "";

                        foreach (ObjectId ObjID in BTrecord)
                        {
                            Entity Ent_intersection = Trans1.GetObject(ObjID, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as Entity;
                            if (Ent_intersection != null && Ent_intersection.IsErased==false)
                            {
                                LayerTableRecord Layer_rec = Trans1.GetObject(Layer_table[Ent_intersection.Layer], OpenMode.ForRead) as LayerTableRecord;
                                if (Ent_intersection is Polyline && Layer_rec.IsOff == false && Layer_rec.IsFrozen == false)
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
                            set_enable_true();
                            return;
                        }
                        
                        System.Data.DataTable dt1 = Functions.Creaza_property_datatable_structure();
                        string Col_Owner = "Owner";
                        string Col_Linelist = "ParcelId";
                        _AGEN_mainform.Data_Table_property = Functions.Scan_parcels(_AGEN_mainform.dt_centerline, _AGEN_mainform.dt_station_equation,dt1, Layer_cu_property, _AGEN_mainform.Project_type, owner_table_name, Linelist_field, Owner_field, Col_Linelist,Col_Owner);

                        if (_AGEN_mainform.Data_Table_property != null)
                        {
                            if (_AGEN_mainform.Data_Table_property.Rows.Count > 0)
                            {
                                Populate_property_file(fisier_prop);

                                Populate_settings_with_control_data( _AGEN_mainform.config_path);
                            }
                        }

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
            set_enable_true();





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

                if (Excel1.Workbooks.Count == 0) Excel1.Visible = _AGEN_mainform.ExcelVisible;

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

                    if (Excel1.Workbooks.Count == 0)
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



        public void Populate_settings_with_control_data(string file_cfg1)
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

                if (Excel1.Workbooks.Count == 0) Excel1.Visible = _AGEN_mainform.ExcelVisible;


                Microsoft.Office.Interop.Excel.Workbook Workbook2 = Excel1.Workbooks.Open(file_cfg1);
                Microsoft.Office.Interop.Excel.Worksheet W2 = null;

                string segment1 = _AGEN_mainform.current_segment;
                if (segment1 == "not defined") segment1 = "";

                foreach (Microsoft.Office.Interop.Excel.Worksheet wsh2 in Workbook2.Worksheets)
                {
                    if (wsh2.Name == "O_dc_" + segment1)
                    {
                        W2 = wsh2;
                    }
                }

                if (W2 == null)
                {
                    W2 = Workbook2.Worksheets.Add(System.Reflection.Missing.Value, Workbook2.Worksheets[Workbook2.Worksheets.Count], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                    W2.Name = "O_dc_" + segment1;
                }

                try
                {


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




                    if (Excel1.Workbooks.Count == 0)
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
                    if (W2 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W2);
                    if (Workbook2 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook2);
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

            set_enable_false();
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
                        set_enable_true();
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
            set_enable_true();


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
