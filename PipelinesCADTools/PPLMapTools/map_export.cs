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
using System.Management;
using Autodesk.AutoCAD.EditorInput;

namespace Alignment_mdi
{
    public partial class map_export : Form
    {
        System.Data.DataTable dtp = null;
        System.Data.DataTable dt_cont = null;
        List<string> lista_layere = null;
        int no_parcels_crossed_by_cl = 0;
        double cl_len = 0;
        int no_parcels_approved = 0;
        double total_length_parcel_approved = 0;
        int no_parcels_denied = 0;
        double total_length_parcel_denied = 0;
        int no_parcels_pending = 0;
        double total_length_parcel_pending = 0;

        public map_export()
        {
            InitializeComponent();
        }

        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_scan);


            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {

                bt1.Enabled = false;

            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();

            lista_butoane.Add(button_scan);
            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }







        private List<string> get_od_fields_from_dwg()
        {
            List<string> lista_fields = new List<string>();
            List<string> lista_OD = get_od_tables_from_dwg();

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                    for (int i = 0; i < lista_OD.Count; ++i)
                    {
                        if (Tables1.IsTableDefined(lista_OD[i]) == true)
                        {
                            Autodesk.Gis.Map.ObjectData.Table tabla1 = Tables1[lista_OD[i]];
                            Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = tabla1.FieldDefinitions;
                            for (int j = 0; j < Field_defs1.Count; ++j)
                            {
                                Autodesk.Gis.Map.ObjectData.FieldDefinition fielddef1 = Field_defs1[j];
                                lista_fields.Add(fielddef1.Name);
                            }
                        }
                    }



                    Trans1.Commit();
                }
            }

            return lista_fields;
        }

        private List<string> get_od_tables_from_dwg()
        {
            List<string> lista_OD = new List<string>();


            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();


                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {

                        Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                        System.Collections.Specialized.StringCollection Nume_tables = new System.Collections.Specialized.StringCollection();
                        Nume_tables = Tables1.GetTableNames();

                        for (int i = 0; i < Nume_tables.Count; i = i + 1)
                        {
                            lista_OD.Add(Nume_tables[i]);

                        }
                    }
                }

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return lista_OD;

        }

        private List<string> get_layers_from_dwg()
        {
            List<string> lista_layers = new List<string>();

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument();
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                Autodesk.AutoCAD.DatabaseServices.LayerTable layer_table = (Autodesk.AutoCAD.DatabaseServices.LayerTable)Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);


                System.Data.DataTable dt1 = new System.Data.DataTable();
                dt1.Columns.Add("ln", typeof(string));


                foreach (ObjectId Layer_id in layer_table)
                {
                    LayerTableRecord Layer1 = (LayerTableRecord)Trans1.GetObject(Layer_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                    string Name_of_layer = Layer1.Name;
                    if (Name_of_layer.Contains("|") == false & Name_of_layer.Contains("$") == false)
                    {
                        dt1.Rows.Add();
                        dt1.Rows[dt1.Rows.Count - 1][0] = Name_of_layer;


                    }
                }

                System.Data.DataTable dt2 = Functions.Sort_data_table(dt1, "ln");
                for (int i = 0; i < dt2.Rows.Count; ++i)
                {
                    lista_layers.Add(dt2.Rows[i][0].ToString());
                }

                Trans1.Commit();
            }
            return lista_layers;
        }



        private void button_scan_properties_Click(object sender, EventArgs e)
        {
            lista_layere = new List<string>();

            Functions.Kill_excel();

            if (Functions.Get_no_of_workbooks_from_Excel() > 0)
            {

            }

            double poly_length = 0;

            scan_mainform Ag = this.MdiParent as scan_mainform;

            comboBox_approved.Items.Clear();
            comboBox_denied.Items.Clear();
            comboBox_pending.Items.Clear();

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;

            ObjectId[] Empty_array = null;
            Editor1.SetImpliedSelection(Empty_array);
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

            Ag.WindowState = FormWindowState.Minimized;

            try
            {
                dtp = new System.Data.DataTable();
                dtp.Columns.Add("sta1", typeof(double));
                dtp.Columns.Add("sta2", typeof(double));
                dtp.Columns.Add("layer", typeof(string));
                dtp.Columns.Add("x1", typeof(double));
                dtp.Columns.Add("y1", typeof(double));
                dtp.Columns.Add("x2", typeof(double));
                dtp.Columns.Add("y2", typeof(double));




                System.Data.DataTable dt_intA = new System.Data.DataTable();

                dt_intA.Columns.Add(scan_mainform.Col_handle, typeof(string));
                dt_intA.Columns.Add(scan_mainform.Col_station, typeof(double));
                dt_intA.Columns.Add("x", typeof(double));
                dt_intA.Columns.Add("y", typeof(double));

                System.Data.DataTable dt_intB = new System.Data.DataTable();

                dt_intB.Columns.Add(scan_mainform.Col_handle, typeof(string));
                dt_intB.Columns.Add(scan_mainform.Col_station, typeof(double));
                dt_intB.Columns.Add("x", typeof(double));
                dt_intB.Columns.Add("y", typeof(double));




                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {

                    set_enable_false();

                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable1 = (Autodesk.AutoCAD.DatabaseServices.BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                        ObjectId selid = ObjectId.Null;

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                        Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the centerline:");
                        Prompt_centerline.SetRejectMessage("\nSelect a polyline!");
                        Prompt_centerline.AllowNone = true;
                        Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline3d), false);
                        Rezultat_centerline = ThisDrawing.Editor.GetEntity(Prompt_centerline);

                        if (Rezultat_centerline.Status != PromptStatus.OK)
                        {
                            set_enable_true();
                            Ag.WindowState = FormWindowState.Normal;
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        bool deleteCL = false;

                        Polyline3d Poly3D = null;

                        Polyline Poly2D = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForRead) as Polyline;
                        if (Poly2D == null)
                        {
                            Poly3D = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForRead) as Polyline3d;
                            selid = Poly3D.ObjectId;
                        }
                        else
                        {
                            selid = Poly2D.ObjectId;
                            Poly3D = Functions.Build_3d_poly_from2D_poly(Poly2D);
                            deleteCL = true;
                        }

                        if (Poly3D != null)
                        {

                            poly_length = Poly3D.Length;

                            Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

                            LayerTable Layer_table = Trans1.GetObject(ThisDrawing.Database.LayerTableId, OpenMode.ForRead) as LayerTable;
                            Poly2D = Functions.Build_2dpoly_from_3d(Poly3D);
                            cl_len = Poly3D.Length;
                            Poly2D.Elevation = 0;

                            foreach (ObjectId ObjID in BTrecord)
                            {
                                Polyline Poly_int = Trans1.GetObject(ObjID, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as Polyline;

                                if (Poly_int != null)
                                {
                                    LayerTableRecord Layer_rec = Trans1.GetObject(Layer_table[Poly_int.Layer], OpenMode.ForRead) as LayerTableRecord;
                                    if (Poly_int.ObjectId != Poly2D.ObjectId && Poly_int.ObjectId != Poly3D.ObjectId && Poly_int.ObjectId != selid && Layer_rec.IsOff == false && Layer_rec.IsFrozen == false)
                                    {

                                        Poly_int.UpgradeOpen();

                                        Poly_int.Elevation = 0;



                                        Point3dCollection Col_int1 = new Point3dCollection();
                                        Col_int1 = Functions.Intersect_on_both_operands(Poly_int, Poly2D);

                                        bool proceseaza = true;

                                        if (Col_int1.Count > 0)
                                        {
                                            if (Col_int1.Count == 1)
                                            {
                                                Point3d pt1 = Col_int1[0];
                                                double sta1 = Math.Round(Poly3D.GetDistanceAtParameter(Poly2D.GetParameterAtPoint(pt1)), 0);
                                                if (Math.Round(sta1, 2) < 0.49 || Math.Round(sta1, 2) > Poly3D.Length - 0.49)
                                                {
                                                    proceseaza = false;
                                                }
                                                else
                                                {
                                                    proceseaza = true;
                                                }
                                            }
                                            else
                                            {
                                                proceseaza = true;
                                            }

                                            if (proceseaza == true)
                                            {
                                                ++no_parcels_crossed_by_cl;

                                                dtp.Rows.Add();
                                                dtp.Rows[dtp.Rows.Count - 1]["layer"] = Poly_int.Layer;
                                                if (lista_layere.Contains(Poly_int.Layer) == false) lista_layere.Add(Poly_int.Layer);
                                                #region object data
                                                using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tables1.GetObjectRecords(Convert.ToUInt32(0), Poly_int.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, false))
                                                {
                                                    if (Records1 != null)
                                                    {
                                                        if (Records1.Count > 0)
                                                        {

                                                            foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                                            {
                                                                Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Record1.TableName];
                                                                Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;


                                                                for (int i = 0; i < Record1.Count; ++i)
                                                                {
                                                                    Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[i];
                                                                    string Nume_field = Field_def1.Name;
                                                                    object valoare1 = Record1[i].StrValue;
                                                                    if (Record1[i].StrValue == "") valoare1 = DBNull.Value;
                                                                    if (valoare1 == null) valoare1 = DBNull.Value;
                                                                    if (dtp.Columns.Contains(Nume_field) == false) dtp.Columns.Add(Nume_field, typeof(string));

                                                                    dtp.Rows[dtp.Rows.Count - 1][Nume_field] = valoare1;

                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                #endregion


                                                Polyline poly1 = new Polyline();
                                                poly1 = Poly_int.Clone() as Polyline;

                                                if (poly1.Closed == false)
                                                {
                                                    poly1.Closed = true;
                                                }

                                                Point3dCollection Col_int = new Point3dCollection();
                                                Col_int = Functions.Intersect_on_both_operands(poly1, Poly2D);

                                                DBObjectCollection Poly_Colection = new DBObjectCollection();
                                                Poly_Colection.Add(poly1);
                                                DBObjectCollection Region_Colectionft = new DBObjectCollection();
                                                Region_Colectionft = Autodesk.AutoCAD.DatabaseServices.Region.CreateFromCurves(Poly_Colection);
                                                Autodesk.AutoCAD.DatabaseServices.Region reg1 = Region_Colectionft[0] as Autodesk.AutoCAD.DatabaseServices.Region;

                                                if (Col_int.Count == 2)
                                                {

                                                    for (int i = 0; i < Col_int.Count; ++i)
                                                    {
                                                        Point3d pt1 = Col_int[i];
                                                        double sta = Poly3D.GetDistanceAtParameter(Poly2D.GetParameterAtPoint(pt1));
                                                        if (dtp.Rows[dtp.Rows.Count - 1]["sta1"] == DBNull.Value)
                                                        {
                                                            dtp.Rows[dtp.Rows.Count - 1]["sta1"] = sta;
                                                            dtp.Rows[dtp.Rows.Count - 1]["x1"] = pt1.X;
                                                            dtp.Rows[dtp.Rows.Count - 1]["y1"] = pt1.Y;
                                                        }
                                                        else
                                                        {

                                                            double sta0 = Convert.ToDouble(dtp.Rows[dtp.Rows.Count - 1]["sta1"]);
                                                            double x0 = Convert.ToDouble(dtp.Rows[dtp.Rows.Count - 1]["x1"]);
                                                            double y0 = Convert.ToDouble(dtp.Rows[dtp.Rows.Count - 1]["y1"]);
                                                            if (sta0 < sta)
                                                            {
                                                                dtp.Rows[dtp.Rows.Count - 1]["sta2"] = sta;
                                                                dtp.Rows[dtp.Rows.Count - 1]["x2"] = pt1.X;
                                                                dtp.Rows[dtp.Rows.Count - 1]["y2"] = pt1.Y;
                                                            }
                                                            else
                                                            {
                                                                dtp.Rows[dtp.Rows.Count - 1]["sta1"] = sta;
                                                                dtp.Rows[dtp.Rows.Count - 1]["x1"] = pt1.X;
                                                                dtp.Rows[dtp.Rows.Count - 1]["y1"] = pt1.Y;
                                                                dtp.Rows[dtp.Rows.Count - 1]["sta2"] = sta0;
                                                                dtp.Rows[dtp.Rows.Count - 1]["x2"] = x0;
                                                                dtp.Rows[dtp.Rows.Count - 1]["y2"] = y0;
                                                            }

                                                        }
                                                    }

                                                }
                                                else if (Col_int.Count > 2)
                                                {
                                                    for (int i = 0; i < Col_int.Count - 1; i = i + 2)
                                                    {
                                                        Point3d pt1 = Col_int[i];
                                                        double sta1 = Poly3D.GetDistanceAtParameter(Poly2D.GetParameterAtPoint(pt1));
                                                        double sta1_2d = Poly2D.GetDistAtPoint(pt1);

                                                        Point3d pt2 = Col_int[i + 1];
                                                        double sta2 = Poly3D.GetDistanceAtParameter(Poly2D.GetParameterAtPoint(pt2));
                                                        double sta2_2d = Poly2D.GetDistAtPoint(pt2);

                                                        double stam = (sta1_2d + sta2_2d) / 2;

                                                        Point3d ptm = Poly2D.GetPointAtDist(stam);

                                                        if (sta2 < sta1)
                                                        {
                                                            double t = sta1;
                                                            sta1 = sta2;
                                                            sta2 = t;
                                                            Point3d ptt = pt1;
                                                            pt1 = pt2;
                                                            pt2 = ptt;

                                                        }

                                                        Autodesk.AutoCAD.BoundaryRepresentation.PointContainment pcm = Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Outside;


                                                        using (Autodesk.AutoCAD.BoundaryRepresentation.Brep Brep_obj = new Autodesk.AutoCAD.BoundaryRepresentation.Brep(reg1))
                                                        {
                                                            if (Brep_obj != null)
                                                            {
                                                                using (Autodesk.AutoCAD.BoundaryRepresentation.BrepEntity ent1 = Brep_obj.GetPointContainment(ptm, out pcm))
                                                                {
                                                                    if (ent1 is Autodesk.AutoCAD.BoundaryRepresentation.Face)
                                                                    {
                                                                        pcm = Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Inside;
                                                                    }
                                                                }
                                                            }
                                                        }

                                                        if (pcm == Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Inside)
                                                        {

                                                            if (i > 0)
                                                            {
                                                                DataRow row1 = dtp.NewRow();
                                                                for (int j = 0; j < dtp.Columns.Count; ++j)
                                                                {
                                                                    row1[j] = dtp.Rows[dtp.Rows.Count - 1][j];
                                                                }
                                                                dtp.Rows.InsertAt(row1, dtp.Rows.Count);

                                                            }



                                                            dtp.Rows[dtp.Rows.Count - 1]["sta1"] = sta1;
                                                            dtp.Rows[dtp.Rows.Count - 1]["sta2"] = sta2;
                                                            dtp.Rows[dtp.Rows.Count - 1]["x1"] = pt1.X;
                                                            dtp.Rows[dtp.Rows.Count - 1]["y1"] = pt1.Y;
                                                            dtp.Rows[dtp.Rows.Count - 1]["x2"] = pt2.X;
                                                            dtp.Rows[dtp.Rows.Count - 1]["y2"] = pt2.Y;

                                                        }
                                                        else
                                                        {

                                                            Point3d startpt = new Point3d(Poly3D.StartPoint.X, Poly3D.StartPoint.Y, 0);
                                                            Point3d endpt = new Point3d(Poly3D.EndPoint.X, Poly3D.EndPoint.Y, 0);
                                                            Autodesk.AutoCAD.BoundaryRepresentation.PointContainment pc_start = Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Outside;
                                                            Autodesk.AutoCAD.BoundaryRepresentation.PointContainment pc_end = Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Outside;


                                                            using (Autodesk.AutoCAD.BoundaryRepresentation.Brep Brep_obj = new Autodesk.AutoCAD.BoundaryRepresentation.Brep(reg1))
                                                            {
                                                                if (Brep_obj != null)
                                                                {
                                                                    using (Autodesk.AutoCAD.BoundaryRepresentation.BrepEntity ent1 = Brep_obj.GetPointContainment(startpt, out pc_start))
                                                                    {
                                                                        if (ent1 is Autodesk.AutoCAD.BoundaryRepresentation.Face)
                                                                        {
                                                                            pc_start = Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Inside;
                                                                        }
                                                                    }

                                                                    using (Autodesk.AutoCAD.BoundaryRepresentation.BrepEntity ent1 = Brep_obj.GetPointContainment(endpt, out pc_end))
                                                                    {
                                                                        if (ent1 is Autodesk.AutoCAD.BoundaryRepresentation.Face)
                                                                        {
                                                                            pc_end = Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Inside;
                                                                        }
                                                                    }
                                                                }
                                                            }

                                                            if (pc_start == Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Inside)
                                                            {

                                                                dtp.Rows[dtp.Rows.Count - 1]["sta1"] = 0;
                                                                dtp.Rows[dtp.Rows.Count - 1]["sta2"] = sta1;
                                                                dtp.Rows[dtp.Rows.Count - 1]["x1"] = startpt.X;
                                                                dtp.Rows[dtp.Rows.Count - 1]["y1"] = startpt.Y;
                                                                dtp.Rows[dtp.Rows.Count - 1]["x2"] = pt1.X;
                                                                dtp.Rows[dtp.Rows.Count - 1]["y2"] = pt1.Y;
                                                                i = i - 1;

                                                            }
                                                            else if (pc_end == Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Inside)
                                                            {
                                                                dtp.Rows[dtp.Rows.Count - 1]["sta1"] = sta2;
                                                                dtp.Rows[dtp.Rows.Count - 1]["sta2"] = Poly3D.Length;
                                                                dtp.Rows[dtp.Rows.Count - 1]["x1"] = pt2.X;
                                                                dtp.Rows[dtp.Rows.Count - 1]["y1"] = pt2.Y;
                                                                dtp.Rows[dtp.Rows.Count - 1]["x2"] = endpt.X;
                                                                dtp.Rows[dtp.Rows.Count - 1]["y2"] = endpt.Y;
                                                            }


                                                        }


                                                    }

                                                }
                                                else if (Col_int.Count == 1)
                                                {
                                                    Point3d pt1 = Col_int[0];
                                                    double sta1 = Poly3D.GetDistanceAtParameter(Poly2D.GetParameterAtPoint(pt1));



                                                    if (Math.Round(sta1, 2) > 0.5 && Math.Round(sta1, 2) < Poly3D.Length - 0.5)
                                                    {

                                                        Point3d startpt = new Point3d(Poly3D.StartPoint.X, Poly3D.StartPoint.Y, 0);
                                                        Point3d endpt = new Point3d(Poly3D.EndPoint.X, Poly3D.EndPoint.Y, 0);
                                                        Autodesk.AutoCAD.BoundaryRepresentation.PointContainment pc_start = Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Outside;
                                                        Autodesk.AutoCAD.BoundaryRepresentation.PointContainment pc_end = Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Outside;


                                                        using (Autodesk.AutoCAD.BoundaryRepresentation.Brep Brep_obj = new Autodesk.AutoCAD.BoundaryRepresentation.Brep(reg1))
                                                        {
                                                            if (Brep_obj != null)
                                                            {
                                                                using (Autodesk.AutoCAD.BoundaryRepresentation.BrepEntity ent1 = Brep_obj.GetPointContainment(startpt, out pc_start))
                                                                {
                                                                    if (ent1 is Autodesk.AutoCAD.BoundaryRepresentation.Face)
                                                                    {
                                                                        pc_start = Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Inside;
                                                                    }
                                                                }

                                                                using (Autodesk.AutoCAD.BoundaryRepresentation.BrepEntity ent1 = Brep_obj.GetPointContainment(endpt, out pc_end))
                                                                {
                                                                    if (ent1 is Autodesk.AutoCAD.BoundaryRepresentation.Face)
                                                                    {
                                                                        pc_end = Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Inside;
                                                                    }
                                                                }
                                                            }
                                                        }


                                                        if (pc_start == Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Inside)
                                                        {

                                                            dtp.Rows[dtp.Rows.Count - 1]["sta1"] = 0;
                                                            dtp.Rows[dtp.Rows.Count - 1]["sta2"] = sta1;
                                                            dtp.Rows[dtp.Rows.Count - 1]["x1"] = startpt.X;
                                                            dtp.Rows[dtp.Rows.Count - 1]["y1"] = startpt.Y;
                                                            dtp.Rows[dtp.Rows.Count - 1]["x2"] = pt1.X;
                                                            dtp.Rows[dtp.Rows.Count - 1]["y2"] = pt1.Y;


                                                        }
                                                        else if (pc_end == Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Inside)
                                                        {
                                                            dtp.Rows[dtp.Rows.Count - 1]["sta1"] = sta1;
                                                            dtp.Rows[dtp.Rows.Count - 1]["sta2"] = Poly3D.Length;

                                                            dtp.Rows[dtp.Rows.Count - 1]["x1"] = pt1.X;
                                                            dtp.Rows[dtp.Rows.Count - 1]["y1"] = pt1.Y;
                                                            dtp.Rows[dtp.Rows.Count - 1]["x2"] = endpt.X;
                                                            dtp.Rows[dtp.Rows.Count - 1]["y2"] = endpt.Y;

                                                        }

                                                    }


                                                }
                                            }
                                        }
                                    }
                                }
                            }



                            dtp = Functions.Sort_data_table(dtp, "sta1");
                            //Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dtp);

                            for (int i = 0; i < dtp.Rows.Count - 1; ++i)
                            {
                                double sta1 = Convert.ToDouble(dtp.Rows[i]["sta1"]);
                                double sta2 = Convert.ToDouble(dtp.Rows[i]["sta2"]);

                                for (int j = i + 1; j < dtp.Rows.Count; ++j)
                                {
                                    double sta11 = Convert.ToDouble(dtp.Rows[j]["sta1"]);
                                    double sta22 = Convert.ToDouble(dtp.Rows[j]["sta2"]);

                                    double x11 = Convert.ToDouble(dtp.Rows[j]["x1"]);
                                    double x22 = Convert.ToDouble(dtp.Rows[j]["x2"]);

                                    double y11 = Convert.ToDouble(dtp.Rows[j]["y1"]);
                                    double y22 = Convert.ToDouble(dtp.Rows[j]["y2"]);

                                    if (sta1 < sta11 && sta2 > sta22)
                                    {

                                        DataRow row1 = dtp.NewRow();
                                        for (int k = 0; k < dtp.Columns.Count; ++k)
                                        {
                                            row1[k] = dtp.Rows[i][k];
                                        }
                                        row1["sta1"] = sta22;
                                        row1["x1"] = x22;
                                        row1["y1"] = y22;
                                        dtp.Rows.InsertAt(row1, dtp.Rows.Count);
                                        dtp.Rows[i]["sta2"] = sta11;
                                        dtp.Rows[i]["x2"] = x11;
                                        dtp.Rows[i]["y2"] = y11;
                                        sta2 = sta11;

                                    }

                                }

                            }

                            dtp = Functions.Sort_data_table(dtp, "sta1");




                            System.Data.DataTable dt_display = new System.Data.DataTable();
                            dt_display = dtp.Copy();
                            dt_display.Columns.Remove("x1");
                            dt_display.Columns.Remove("y1");
                            dt_display.Columns.Remove("x2");
                            dt_display.Columns.Remove("y2");

                            for (int i = 0; i < dt_display.Rows.Count; ++i)
                            {
                                double sta1 = Convert.ToDouble(dt_display.Rows[i]["sta1"]);
                                double sta2 = Convert.ToDouble(dt_display.Rows[i]["sta2"]);
                                dt_display.Rows[i]["sta1"] = Math.Round(sta1, 0);
                                dt_display.Rows[i]["sta2"] = Math.Round(sta2, 0);


                            }


                            if (dtp != null)
                            {
                                dataGridView_prop.DataSource = dt_display;
                                dataGridView_prop.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                                dataGridView_prop.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                                dataGridView_prop.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                                dataGridView_prop.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                                dataGridView_prop.DefaultCellStyle.ForeColor = Color.White;
                                dataGridView_prop.EnableHeadersVisualStyles = false;
                            }

                            if (deleteCL == true) Poly3D.Erase();

                            dtp.Columns["x1"].SetOrdinal(dtp.Columns.Count - 1);
                            dtp.Columns["y1"].SetOrdinal(dtp.Columns.Count - 1);
                            dtp.Columns["x2"].SetOrdinal(dtp.Columns.Count - 1);
                            dtp.Columns["y2"].SetOrdinal(dtp.Columns.Count - 1);

                            fill_comboboxes();
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
;
            set_enable_true();





            Ag.WindowState = FormWindowState.Normal;

        }

        private void button_generate_report_Click(object sender, EventArgs e)
        {

            if (dtp != null && dtp.Rows.Count > 0)
            {
                dt_cont = new System.Data.DataTable();
                dt_cont.Columns.Add()

                if (radioButton_layers.Checked == true)
                {

                    if (comboBox_approved.Text == comboBox_denied.Text || comboBox_pending.Text == comboBox_denied.Text || comboBox_approved.Text == comboBox_pending.Text)
                    {
                        MessageBox.Show("Please check your comboboxes because they have identical values!");
                    }
                    else
                    {
                        for (int i = dtp.Rows.Count - 1; i >= 0; --i)
                        {
                            if (dtp.Rows[i]["sta1"] != DBNull.Value && dtp.Rows[i]["sta2"] != DBNull.Value)
                            {
                                double sta1 = Convert.ToDouble(dtp.Rows[i]["sta1"]);
                                double sta2 = Convert.ToDouble(dtp.Rows[i]["sta2"]);

                                if (dtp.Rows[i]["layer"] != DBNull.Value)
                                {
                                    string layer1 = Convert.ToString(dtp.Rows[i]["layer"]);

                                    if (layer1 == comboBox_approved.Text)
                                    {
                                        ++no_parcels_approved;
                                        total_length_parcel_approved = total_length_parcel_approved + sta2 - sta1;

                                    }
                                    if (layer1 == comboBox_denied.Text)
                                    {
                                        ++no_parcels_denied;
                                        total_length_parcel_denied = total_length_parcel_denied + sta2 - sta1;

                                    }
                                    if (layer1 == comboBox_pending.Text)
                                    {
                                        ++no_parcels_pending;
                                        total_length_parcel_pending = total_length_parcel_pending + sta2 - sta1;

                                    }
                                    if (Functions.is_dan_popescu() == false)
                                    {
                                        if (layer1 == comboBox_pending.Text || layer1 == comboBox_denied.Text || layer1 == comboBox_approved.Text)
                                        {
                                            dtp.Rows[i].Delete();

                                        }
                                    }
                                }

                            }
                        }
                    }
                }

                else if (radioButton_object_data.Checked == true)
                {
                    for (int i = dtp.Rows.Count - 1; i >= 0; --i)
                    {
                        if (dtp.Rows[i]["sta1"] != DBNull.Value && dtp.Rows[i]["sta2"] != DBNull.Value)
                        {
                            double sta1 = Convert.ToDouble(dtp.Rows[i]["sta1"]);
                            double sta2 = Convert.ToDouble(dtp.Rows[i]["sta2"]);

                            if (dtp.Rows[i]["layer"] != DBNull.Value)
                            {
                                string layer1 = Convert.ToString(dtp.Rows[i]["layer"]);

                                if (layer1 == comboBox_approved.Text)
                                {
                                    if (dtp.Rows[i][comboBox_denied.Text] != DBNull.Value)
                                    {
                                        string valoare1 = Convert.ToString(dtp.Rows[i][comboBox_denied.Text]);

                                        if (valoare1.ToLower() == "approved")
                                        {
                                            ++no_parcels_approved;
                                            total_length_parcel_approved = total_length_parcel_approved + sta2 - sta1;

                                        }
                                        else if (valoare1.ToLower() == "denied")
                                        {
                                            ++no_parcels_denied;
                                            total_length_parcel_denied = total_length_parcel_denied + sta2 - sta1;

                                        }
                                        else if (valoare1.ToLower() == "pending")
                                        {
                                            ++no_parcels_pending;
                                            total_length_parcel_pending = total_length_parcel_pending + sta2 - sta1;

                                        }
                                        else
                                        {
                                            MessageBox.Show("You have a parcel with non standard data between stations " + Functions.Get_chainage_from_double(sta1, "f", 0) + " to " + Functions.Get_chainage_from_double(sta2, "f", 0));
                                            return;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

            }

            create_scanning_report(dtp);

        }


        private void create_scanning_report(System.Data.DataTable dt1)
        {
            if (dt1 != null)
            {
                if (dt1.Rows.Count > 0)
                {
                    Microsoft.Office.Interop.Excel.Workbook Workbook1 = Functions.Get_NEW_workbook_from_Excel();
                    Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.ActiveSheet;


                    W1.Name = "CAD Input";
                    int maxRows = dt1.Rows.Count;
                    int maxCols = dt1.Columns.Count;

                    W1.Range[W1.Columns[3], W1.Columns[maxCols - 5]].NumberFormat = "@";
                    W1.Range[W1.Columns[maxCols - 4], W1.Columns[maxCols]].NumberFormat = "0.00";
                    W1.Columns["A:B"].NumberFormat = "0+00";
                    W1.Columns["A:G"].ColumnWidth = 15.11;
                    Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[3, 1], W1.Cells[maxRows + 2, maxCols]];
                    object[,] values1 = new object[maxRows, maxCols];

                    for (int i = 0; i < maxRows; ++i)
                    {
                        for (int j = 0; j < maxCols; ++j)
                        {
                            if (dt1.Rows[i][j] != DBNull.Value)
                            {
                                values1[i, j] = Convert.ToString(dt1.Rows[i][j]);
                            }
                        }
                    }

                    for (int i = 0; i < dt1.Columns.Count; ++i)
                    {
                        W1.Cells[2, i + 1].value2 = dt1.Columns[i].ColumnName;
                    }


                    range1.Value2 = values1;
                    range1 = W1.Range[W1.Cells[1, 1], W1.Cells[1, maxCols]];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Value2 = "CAD Input";
                    range1.Font.Name = "Arial Black";
                    range1.Font.Size = 20;
                    range1.Font.Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleSingle;

                    if (Workbook1.Worksheets.Count == 1)
                    {
                        Microsoft.Office.Interop.Excel.Worksheet W3 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[1], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(W3);
                    }

                    Microsoft.Office.Interop.Excel.Worksheet W2 = Workbook1.Worksheets[2];
                    W2.Name = "Survey Permission";
                    W2.Columns["A:E"].ColumnWidth = 11;
                    W2.Rows[1].RowHeight = 40;
                    W2.Rows["2:27"].RowHeight = 15.75;
                    W2.Columns["H:I"].ColumnWidth = 8.43;
                    W2.Columns["J"].ColumnWidth = 11.43;
                    W2.Columns["K"].ColumnWidth = 13;
                    W2.Columns["L"].ColumnWidth = 23.86;
                    W2.Columns["H:L"].NumberFormat = "General";




                    range1 = W2.Range["A1:F1"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Value2 = "Survey Permission Summary By CL";
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 20;
                    range1.Font.Bold = true;
                    range1.Interior.ColorIndex = 40;
                    range1.Font.ColorIndex = 3;

                    range1 = W2.Range["A2:D2"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 11;
                    range1.Interior.ColorIndex = 15;
                    range1.Font.ColorIndex = 1;
                    range1.Font.Bold = true;

                    range1 = W2.Range["E2"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 11;
                    range1.Interior.ColorIndex = 15;
                    range1.Font.ColorIndex = 1;
                    range1.Font.Bold = true;
                    range1.Value2 = "Miles";


                    range1 = W2.Range["F2"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 11;
                    range1.Interior.ColorIndex = 15;
                    range1.Font.ColorIndex = 1;
                    range1.Font.Bold = true;
                    range1.Value2 = "Feet";


                    range1 = W2.Range["A3:D3"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 11;
                    range1.Font.ColorIndex = 1;
                    range1.Font.Bold = true;
                    range1.Value2 = "Total Parcels Crossed by Centerline";

                    range1 = W2.Range["E3"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 11;
                    range1.Font.ColorIndex = 1;
                    range1.Font.Bold = true;
                    range1.Value2 = no_parcels_crossed_by_cl;

                    range1 = W2.Range["A4:D4"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 11;
                    range1.Font.ColorIndex = 1;
                    range1.Font.Bold = true;
                    range1.Value2 = "Total Length of Centerline";

                    range1 = W2.Range["E4"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 11;

                    range1.Font.ColorIndex = 1;
                    range1.Font.Bold = true;
                    range1.Value2 = Math.Round(cl_len / 5280, 1);

                    range1 = W2.Range["F4"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 11;
                    range1.Font.ColorIndex = 1;
                    range1.Font.Bold = true;
                    range1.Value2 = Math.Round(cl_len, 0);

                    #region approved
                    range1 = W2.Range["A5:D5"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 11;
                    range1.Font.ColorIndex = 1;
                    range1.Font.Bold = true;
                    range1.Interior.ColorIndex = 43;
                    range1.Value2 = "Total Parcels Approved";

                    range1 = W2.Range["A6:D6"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 11;
                    range1.Font.ColorIndex = 1;
                    range1.Font.Bold = true;
                    range1.Interior.ColorIndex = 43;
                    range1.Value2 = "Total Length Across Approved Parcels";

                    range1 = W2.Range["A7:D7"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 11;
                    range1.Font.ColorIndex = 1;
                    range1.Font.Bold = true;
                    range1.Interior.ColorIndex = 43;
                    range1.Value2 = "Percentage of Centerline Approved";

                    range1 = W2.Range["E5"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 11;
                    range1.Font.ColorIndex = 1;
                    range1.Font.Bold = true;
                    range1.Value2 = no_parcels_approved;

                    range1 = W2.Range["E6"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 11;
                    range1.Font.ColorIndex = 1;
                    range1.Font.Bold = true;
                    range1.Value2 = Math.Round(total_length_parcel_approved / 5280, 1);

                    range1 = W2.Range["F6"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 11;
                    range1.Font.ColorIndex = 1;
                    range1.Font.Bold = true;
                    range1.Value2 = Math.Round(total_length_parcel_approved, 0);

                    range1 = W2.Range["E7"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 11;
                    range1.Font.ColorIndex = 1;
                    range1.Font.Bold = true;
                    range1.Value2 = total_length_parcel_approved / cl_len;
                    range1.NumberFormat = "0.0%";
                    #endregion

                    #region denied
                    range1 = W2.Range["A8:D8"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 11;
                    range1.Font.ColorIndex = 2;
                    range1.Font.Bold = true;
                    range1.Interior.ColorIndex = 3;
                    range1.Value2 = "Total Parcels Denied";

                    range1 = W2.Range["A9:D9"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 11;
                    range1.Font.ColorIndex = 2;
                    range1.Font.Bold = true;
                    range1.Interior.ColorIndex = 3;
                    range1.Value2 = "Total Length Across Denied Parcels";

                    range1 = W2.Range["A10:D10"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 11;
                    range1.Font.ColorIndex = 2;
                    range1.Font.Bold = true;
                    range1.Interior.ColorIndex = 3;
                    range1.Value2 = "Percentage of Centerline Denied";


                    range1 = W2.Range["E8"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 11;
                    range1.Font.ColorIndex = 1;
                    range1.Font.Bold = true;
                    range1.Value2 = no_parcels_denied;

                    range1 = W2.Range["E9"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 11;
                    range1.Font.ColorIndex = 1;
                    range1.Font.Bold = true;
                    range1.Value2 = Math.Round(total_length_parcel_denied / 5280, 1);

                    range1 = W2.Range["F9"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 11;
                    range1.Font.ColorIndex = 1;
                    range1.Font.Bold = true;
                    range1.Value2 = Math.Round(total_length_parcel_denied, 0);

                    range1 = W2.Range["E10"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 11;
                    range1.Font.ColorIndex = 1;
                    range1.Font.Bold = true;
                    range1.Value2 = total_length_parcel_denied / cl_len;
                    range1.NumberFormat = "0.0%";
                    #endregion

                    #region pending
                    range1 = W2.Range["A11:D11"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 11;
                    range1.Font.ColorIndex = 1;
                    range1.Font.Bold = true;
                    range1.Interior.ColorIndex = 6;
                    range1.Value2 = "Total Parcels Pending";

                    range1 = W2.Range["A12:D12"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 11;
                    range1.Font.ColorIndex = 1;
                    range1.Font.Bold = true;
                    range1.Interior.ColorIndex = 6;
                    range1.Value2 = "Total Length Across Pending Parcels";

                    range1 = W2.Range["A13:D13"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 11;
                    range1.Font.ColorIndex = 1;
                    range1.Font.Bold = true;
                    range1.Interior.ColorIndex = 6;
                    range1.Value2 = "Percentage of Centerline Pending";


                    range1 = W2.Range["E11"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 11;
                    range1.Font.ColorIndex = 1;
                    range1.Font.Bold = true;
                    range1.Value2 = no_parcels_pending;

                    range1 = W2.Range["E12"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 11;
                    range1.Font.ColorIndex = 1;
                    range1.Font.Bold = true;
                    range1.Value2 = Math.Round(total_length_parcel_pending / 5280, 1);

                    range1 = W2.Range["F12"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 11;
                    range1.Font.ColorIndex = 1;
                    range1.Font.Bold = true;
                    range1.Value2 = Math.Round(total_length_parcel_pending, 0);

                    range1 = W2.Range["E13"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 11;
                    range1.Font.ColorIndex = 1;
                    range1.Font.Bold = true;
                    range1.Value2 = total_length_parcel_pending / cl_len;
                    range1.NumberFormat = "0.0%";
                    #endregion

                    #region Contiguous
                    range1 = W2.Range["H1:L1"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Value2 = "Summary of Contiguous Approved Parcels";
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 20;
                    range1.Font.Bold = true;
                    range1.Interior.ColorIndex = 43;
                    range1.Font.ColorIndex = 3;


                    range1 = W2.Range["H2"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 11;
                    range1.Interior.ColorIndex = 15;
                    range1.Font.ColorIndex = 1;
                    range1.Font.Bold = true;
                    range1.Value2 = "Count";

                    range1 = W2.Range["I2"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 11;
                    range1.Interior.ColorIndex = 15;
                    range1.Font.ColorIndex = 1;
                    range1.Font.Bold = true;
                    range1.Value2 = "Start MP";

                    range1 = W2.Range["J2"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 11;
                    range1.Interior.ColorIndex = 15;
                    range1.Font.ColorIndex = 1;
                    range1.Font.Bold = true;
                    range1.Value2 = "End MP";

                    range1 = W2.Range["K2"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 11;
                    range1.Interior.ColorIndex = 15;
                    range1.Font.ColorIndex = 1;
                    range1.Font.Bold = true;
                    range1.Value2 = "Parcel Count";

                    range1 = W2.Range["L2"];
                    range1.MergeCells = true;
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range1.Font.Name = "Calibri";
                    range1.Font.Size = 11;
                    range1.Interior.ColorIndex = 15;
                    range1.Font.ColorIndex = 1;
                    range1.Font.Bold = true;
                    range1.Value2 = "Length in Feet";

                    #endregion

                }
            }
        }


        private void fill_comboboxes()
        {
            comboBox_approved.Items.Clear();
            comboBox_denied.Items.Clear();
            comboBox_pending.Items.Clear();



            List<string> lista1 = get_layers_from_dwg();
            List<string> lista2 = get_od_fields_from_dwg();

            if (lista_layere != null && lista_layere.Count > 0)
            {
                for (int i = lista1.Count - 1; i >= 0; --i)
                {
                    if (lista_layere.Contains(lista1[i]) == false)
                    {
                        lista1.RemoveAt(i);
                    }
                }
            }

            if (radioButton_layers.Checked == true)
            {
                comboBox_approved.Visible = true;
                comboBox_denied.Visible = true;
                comboBox_pending.Visible = true;
                label_approved.Text = "Approved";
                label_denied.Text = "Denied";
                label_pending.Text = "Pending";
                label_approved.Visible = true;
                label_denied.Visible = true;
                label_pending.Visible = true;
                label_type.Visible = true;

                if (lista1.Count > 0)
                    {
                        for (int i = 0; i < lista1.Count; ++i)
                        {
                            comboBox_approved.Items.Add(lista1[i]);
                            comboBox_denied.Items.Add(lista1[i]);
                            comboBox_pending.Items.Add(lista1[i]);
                        }
           
                    label_type.Text = "Layers:";
                }
            }

            if (radioButton_object_data.Checked == true)
            {
                comboBox_approved.Visible = true;
                comboBox_denied.Visible = true;
                comboBox_pending.Visible = false;
                label_approved.Text = "Layer";
                label_denied.Text = "OD Field";

                label_approved.Visible = true;
                label_denied.Visible = true;
                label_pending.Visible = false;
                label_type.Visible = false;

                if (lista1.Count > 0)
                {
                    for (int i = 0; i < lista1.Count; ++i)
                    {
                        comboBox_approved.Items.Add(lista1[i]);
                    }
                }

                if (lista2.Count > 0)
                {
                    for (int i = 0; i < lista2.Count; ++i)
                    {
                        comboBox_denied.Items.Add(lista2[i]);
                    }

                }


               

            }
        }


        private void label_for_dan_Click(object sender, EventArgs e)
        {
            if (Functions.is_dan_popescu() == true)
            {

            }
        }

        private void radioButton_layers_CheckedChanged(object sender, EventArgs e)
        {
            fill_comboboxes();
        }


    }

    public class Command_class
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

                case "120E4B54":
                    return true;
                case "F6633173":
                    return true;
                case "40D6BDCB":
                    return true;
                case "18399D24":
                    return true;


                case "B63AD3F6":
                    return true;
                default:
                    try
                    {
                        string UserDNS = Environment.GetEnvironmentVariable("USERDNSDOMAIN");
                        if (UserDNS.ToUpper() == "HMMG.CC" | UserDNS.ToLower() == "mottmac.group.int")
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


        [CommandMethod("_scan")]
        public void Show_HDD_mainForm()
        {
            if (isSECURE() == true)
            {



                foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
                {
                    if (Forma1 is Alignment_mdi.scan_mainform)
                    {
                        Forma1.Focus();
                        Forma1.WindowState = System.Windows.Forms.FormWindowState.Normal;
                        Forma1.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - Forma1.Width) / 2,
                          (Screen.PrimaryScreen.WorkingArea.Height - Forma1.Height) / 2);
                        return;
                    }
                }

                try
                {
                    Alignment_mdi.scan_mainform forma2 = new Alignment_mdi.scan_mainform();
                    Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma2);
                    forma2.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - forma2.Width) / 2,
                         (Screen.PrimaryScreen.WorkingArea.Height - forma2.Height) / 2);
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
