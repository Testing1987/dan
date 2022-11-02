using Autodesk.AutoCAD.EditorInput;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.Settings;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace p_network
{
    public partial class form1 : Form
    {
        private bool clickdragdown;
        private System.Drawing.Point lastLocation;






        private void make_variables_null()
        {

        }

        public form1()
        {
            InitializeComponent();




        }

        [CommandMethod("pn_intersector")]
        public void ShowForm()
        {
            if (Functions.isSECURE() == true)
            {
                foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
                {
                    if (Forma1 is form1)
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
                    form1 forma2 = new form1();
                    Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma2);
                    forma2.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - forma2.Width) / 2,
                         (Screen.PrimaryScreen.WorkingArea.Height - forma2.Height) / 2);
                }
                catch (System.Exception EX)
                {
                    MessageBox.Show(EX.Message);
                }
            }
        }


        protected override void OnLoad(EventArgs e)
        {
            // Hides the ugly border around the mdi container (main form)
            var mdiclient = this.Controls.OfType<MdiClient>().Single();
            this.SuspendLayout();
            mdiclient.SuspendLayout();
            var hdiff = mdiclient.Size.Width - mdiclient.ClientSize.Width;
            var vdiff = mdiclient.Size.Height - mdiclient.ClientSize.Height;
            var size = new System.Drawing.Size(mdiclient.Width + hdiff, mdiclient.Height + vdiff);
            var location = new System.Drawing.Point(mdiclient.Left - (hdiff / 2), mdiclient.Top - (vdiff / 2));
            mdiclient.Dock = DockStyle.None;
            mdiclient.Size = size;
            mdiclient.Location = location;
            mdiclient.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom;
            mdiclient.ResumeLayout(true);
            this.ResumeLayout(true);
            base.OnLoad(e);
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
                this.Location = new System.Drawing.Point(
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


            make_variables_null();


            this.Close();
        }
        private void button_minimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void button_scan_Click(object sender, EventArgs e)
        {
            CivilDocument doc1 = CivilApplication.ActiveDocument;
            Editor editor1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            SettingsPipeNetwork set1 = doc1.Settings.GetSettings<SettingsPipeNetwork>() as SettingsPipeNetwork;



            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {

                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                        Autodesk.AutoCAD.DatabaseServices.LayerTable layer_table = (Autodesk.AutoCAD.DatabaseServices.LayerTable)Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

                        //Vous devez ajouter une référence à AecBaseMgd.dll(dans le répertoire d'installation).
                        ObjectId dsvid = Autodesk.Aec.ApplicationServices.DrawingSetupVariables.GetInstance(ThisDrawing.Database, false);
                        Autodesk.Aec.ApplicationServices.DrawingSetupVariables dsv = Trans1.GetObject(dsvid, OpenMode.ForRead) as Autodesk.Aec.ApplicationServices.DrawingSetupVariables;
                        Matrix3d CurentUCSmatrix = Editor1.CurrentUserCoordinateSystem;

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                        Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the centerline:");
                        Prompt_centerline.SetRejectMessage("\nSelect a polyline or a 3D polyline!");
                        Prompt_centerline.AllowNone = true;
                        Prompt_centerline.AddAllowedClass(typeof(Polyline), false);
                        Prompt_centerline.AddAllowedClass(typeof(Polyline3d), false);
                        Rezultat_centerline = ThisDrawing.Editor.GetEntity(Prompt_centerline);

                        if (Rezultat_centerline.Status != PromptStatus.OK)
                        {
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        

                        Polyline poly2d = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForWrite) as Polyline;
                        Polyline3d poly3d = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForRead) as Polyline3d;

                        if (poly2d == null)
                        {
                            poly2d = Functions.Build_2dpoly_from_3d(poly3d);
                        }


                        bool delete_poly3d = false;

                        if (poly3d == null && poly2d != null)
                        {
                            System.Data.DataTable dt_cl = new System.Data.DataTable();
                            dt_cl.Columns.Add("X", typeof(double));
                            dt_cl.Columns.Add("Y", typeof(double));
                            dt_cl.Columns.Add("Z", typeof(double));

                            for (int i = 0; i < poly2d.NumberOfVertices; ++i)
                            {
                                Point2d pt1 = poly2d.GetPoint2dAt(i);
                                dt_cl.Rows.Add();
                                dt_cl.Rows[i]["X"] = pt1.X;
                                dt_cl.Rows[i]["Y"] = pt1.Y;
                                dt_cl.Rows[i]["Z"] = 0;
                            }



                            poly3d = Functions.Build_3d_poly_for_scanning(dt_cl);
                            delete_poly3d = true;
                        }

                        System.Data.DataTable dt1 = new System.Data.DataTable();
                        dt1.Columns.Add("Type", typeof(string));
                        dt1.Columns.Add("Layer", typeof(string));
                        dt1.Columns.Add("Sta", typeof(double));
                        dt1.Columns.Add("X", typeof(double));
                        dt1.Columns.Add("Y", typeof(double));
                        dt1.Columns.Add("Z on CL", typeof(double));
                        dt1.Columns.Add("Z on object", typeof(double));

                        ObjectIdCollection col1 = new ObjectIdCollection();
                        col1.Add(poly2d.ObjectId);
                        col1.Add(poly3d.ObjectId);

                        List<string> lista_layere = new List<string>();
                        foreach (ObjectId Layer_id in layer_table)
                        {
                            LayerTableRecord ltr = (LayerTableRecord)Trans1.GetObject(Layer_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);

                            if (ltr.Name.Contains("|") == false && ltr.Name.Contains("$") == false && ltr.IsFrozen == false && ltr.IsOff == false)
                            {

                                lista_layere.Add(ltr.Name);

                            }
                        }




                        foreach (ObjectId id1 in BTrecord)
                        {
                            if (col1.Contains(id1) == false)
                            {
                                Polyline poly_int_2d = Trans1.GetObject(id1, OpenMode.ForRead) as Polyline;
                                Polyline3d poly_int_3d = Trans1.GetObject(id1, OpenMode.ForRead) as Polyline3d;
                                Line line_int = Trans1.GetObject(id1, OpenMode.ForRead) as Line;


                                Entity ent1 = Trans1.GetObject(id1, OpenMode.ForRead) as Entity;

                                if (ent1 != null && lista_layere.Contains(ent1.Layer) == true)
                                {
                                    if (poly_int_2d != null)
                                    {
                                        poly2d.Elevation = poly_int_2d.Elevation;
                                        Point3dCollection col_int = Functions.Intersect_on_both_operands(poly2d, poly_int_2d);
                                        if (col_int.Count > 0)
                                        {


                                            for (int i = 0; i < col_int.Count; ++i)
                                            {
                                                dt1.Rows.Add();
                                                dt1.Rows[dt1.Rows.Count - 1]["Type"] = "Polyline";
                                                dt1.Rows[dt1.Rows.Count - 1]["Layer"] = poly_int_2d.Layer;
                                                dt1.Rows[dt1.Rows.Count - 1]["X"] = col_int[i].X;
                                                dt1.Rows[dt1.Rows.Count - 1]["Y"] = col_int[i].Y;

                                                Point3d pt1 = poly2d.GetClosestPointTo(col_int[i], Vector3d.ZAxis, false);
                                                double param1 = poly2d.GetParameterAtPoint(pt1);
                                                if (param1 > poly3d.EndParam) param1 = poly3d.EndParam;
                                                dt1.Rows[dt1.Rows.Count - 1]["Z on CL"] = poly3d.GetPointAtParameter(param1).Z;
                                                dt1.Rows[dt1.Rows.Count - 1]["Sta"] = poly3d.GetDistanceAtParameter(param1);
                                                dt1.Rows[dt1.Rows.Count - 1]["Z on object"] = col_int[i].Z;

                                                add_object_data_to_datatable(dt1, Tables1, poly_int_2d.ObjectId);

                                            }


                                        }

                                    }

                                    if (poly_int_3d != null)
                                    {
                                        Polyline poly1 = Functions.Build_2dpoly_from_3d(poly_int_3d);
                                        poly1.Elevation = poly2d.Elevation;



                                        Point3dCollection col_int = Functions.Intersect_on_both_operands(poly2d, poly1);
                                        if (col_int.Count > 0)
                                        {
                                            for (int i = 0; i < col_int.Count; ++i)
                                            {
                                                dt1.Rows.Add();
                                                dt1.Rows[dt1.Rows.Count - 1]["Type"] = "Polyline3d";
                                                dt1.Rows[dt1.Rows.Count - 1]["Layer"] = poly_int_3d.Layer;
                                                dt1.Rows[dt1.Rows.Count - 1]["X"] = col_int[i].X;
                                                dt1.Rows[dt1.Rows.Count - 1]["Y"] = col_int[i].Y;

                                                Point3d pt1 = poly2d.GetClosestPointTo(col_int[i], Vector3d.ZAxis, false);
                                                double param1 = poly2d.GetParameterAtPoint(pt1);
                                                if (param1 > poly3d.EndParam) param1 = poly3d.EndParam;
                                                dt1.Rows[dt1.Rows.Count - 1]["Z on CL"] = poly3d.GetPointAtParameter(param1).Z;
                                                dt1.Rows[dt1.Rows.Count - 1]["Sta"] = poly3d.GetDistanceAtParameter(param1);

                                                Point3d p2 = poly1.GetClosestPointTo(col_int[i], Vector3d.ZAxis, false);
                                                double param2 = poly1.GetParameterAtPoint(p2);
                                                if (param2 > poly_int_3d.EndParam) param2 = poly_int_3d.EndParam;

                                                double z_obj = poly_int_3d.GetPointAtParameter(param2).Z;

                                                dt1.Rows[dt1.Rows.Count - 1]["Z on object"] = z_obj;
                                                add_object_data_to_datatable(dt1, Tables1, poly_int_3d.ObjectId);
                                            }
                                        }
                                    }

                                    if (line_int != null)
                                    {
                                        Polyline poly1 = new Polyline();
                                        poly1.AddVertexAt(0, new Point2d(line_int.StartPoint.X, line_int.StartPoint.Y), 0, 0, 0);
                                        poly1.AddVertexAt(1, new Point2d(line_int.EndPoint.X, line_int.EndPoint.Y), 0, 0, 0);

                                        poly1.Elevation = poly2d.Elevation;



                                        Point3dCollection col_int = Functions.Intersect_on_both_operands(poly2d, poly1);
                                        if (col_int.Count > 0)
                                        {
                                            for (int i = 0; i < col_int.Count; ++i)
                                            {
                                                dt1.Rows.Add();
                                                dt1.Rows[dt1.Rows.Count - 1]["Type"] = "Line";
                                                dt1.Rows[dt1.Rows.Count - 1]["Layer"] = line_int.Layer;
                                                dt1.Rows[dt1.Rows.Count - 1]["X"] = col_int[i].X;
                                                dt1.Rows[dt1.Rows.Count - 1]["Y"] = col_int[i].Y;

                                                Point3d pt1 = poly2d.GetClosestPointTo(col_int[i], Vector3d.ZAxis, false);
                                                double param1 = poly2d.GetParameterAtPoint(pt1);
                                                if (param1 > poly3d.EndParam) param1 = poly3d.EndParam;
                                                dt1.Rows[dt1.Rows.Count - 1]["Z on CL"] = poly3d.GetPointAtParameter(param1).Z;
                                                dt1.Rows[dt1.Rows.Count - 1]["Sta"] = poly3d.GetDistanceAtParameter(param1);


                                                System.Data.DataTable dt3 = new System.Data.DataTable();
                                                dt3.Columns.Add("X", typeof(double));
                                                dt3.Columns.Add("Y", typeof(double));
                                                dt3.Columns.Add("Z", typeof(double));


                                                dt3.Rows.Add();
                                                dt3.Rows[0]["X"] = line_int.StartPoint.X;
                                                dt3.Rows[0]["Y"] = line_int.StartPoint.Y;
                                                dt3.Rows[0]["Z"] = line_int.StartPoint.Z;
                                                dt3.Rows.Add();
                                                dt3.Rows[1]["X"] = line_int.EndPoint.X;
                                                dt3.Rows[1]["Y"] = line_int.EndPoint.Y;
                                                dt3.Rows[1]["Z"] = line_int.EndPoint.Z;



                                                Polyline3d poly_line_3d = Functions.Build_3d_poly_for_scanning(dt3);
                                                Point3d p2 = poly1.GetClosestPointTo(col_int[i], Vector3d.ZAxis, false);
                                                double param2 = poly1.GetParameterAtPoint(p2);
                                                if (param2 > poly_line_3d.EndParam) param2 = poly_line_3d.EndParam;

                                                double z_obj = poly_line_3d.GetPointAtParameter(param2).Z;

                                                dt1.Rows[dt1.Rows.Count - 1]["Z on object"] = z_obj;
                                                add_object_data_to_datatable(dt1, Tables1, line_int.ObjectId);
                                                poly_line_3d.Erase();
                                            }
                                        }
                                    }

                                    Autodesk.Civil.DatabaseServices.Pipe pipe1 = Trans1.GetObject(id1, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Pipe;
                                    if (pipe1 != null)
                                    {
                                        Point3d pt1 = pipe1.StartPoint;
                                        Point3d pt2 = pipe1.EndPoint;

                                        Polyline poly_pipe_2d = new Polyline();
                                        poly_pipe_2d.AddVertexAt(0, new Point2d(pt1.X, pt1.Y), 0, 0, 0);
                                        poly_pipe_2d.AddVertexAt(1, new Point2d(pt2.X, pt2.Y), 0, 0, 0);
                                        poly_pipe_2d.Elevation = poly2d.Elevation;


                                        System.Data.DataTable dt3 = new System.Data.DataTable();
                                        dt3.Columns.Add("X", typeof(double));
                                        dt3.Columns.Add("Y", typeof(double));
                                        dt3.Columns.Add("Z", typeof(double));


                                        dt3.Rows.Add();
                                        dt3.Rows[0]["X"] = pt1.X;
                                        dt3.Rows[0]["Y"] = pt1.Y;
                                        dt3.Rows[0]["Z"] = pt1.Z;
                                        dt3.Rows.Add();
                                        dt3.Rows[1]["X"] = pt2.X;
                                        dt3.Rows[1]["Y"] = pt2.Y;
                                        dt3.Rows[1]["Z"] = pt2.Z;



                                        Polyline3d poly_pipe_3d = Functions.Build_3d_poly_for_scanning(dt3);


                                        Point3dCollection col_int = Functions.Intersect_on_both_operands(poly2d, poly_pipe_2d);
                                        if (col_int.Count > 0)
                                        {
                                            for (int i = 0; i < col_int.Count; ++i)
                                            {
                                                dt1.Rows.Add();
                                                dt1.Rows[dt1.Rows.Count - 1]["Type"] = "PipeNetwork";
                                                dt1.Rows[dt1.Rows.Count - 1]["Layer"] = pipe1.Layer;
                                                dt1.Rows[dt1.Rows.Count - 1]["X"] = col_int[i].X;
                                                dt1.Rows[dt1.Rows.Count - 1]["Y"] = col_int[i].Y;

                                                Point3d p1 = poly2d.GetClosestPointTo(col_int[i], Vector3d.ZAxis, false);
                                                double param1 = poly2d.GetParameterAtPoint(p1);
                                                if (param1 > poly3d.EndParam) param1 = poly3d.EndParam;

                                                Point3d p2 = poly_pipe_2d.GetClosestPointTo(col_int[i], Vector3d.ZAxis, false);
                                                double param2 = poly_pipe_2d.GetParameterAtPoint(p2);
                                                if (param2 > poly_pipe_3d.EndParam) param2 = poly_pipe_3d.EndParam;

                                                double z_obj = poly_pipe_3d.GetPointAtParameter(param2).Z;
                                                double diam1 = pipe1.InnerDiameterOrWidth;

                                                dt1.Rows[dt1.Rows.Count - 1]["Z on CL"] = poly3d.GetPointAtParameter(param1).Z;
                                                dt1.Rows[dt1.Rows.Count - 1]["Sta"] = poly3d.GetDistanceAtParameter(param1);
                                                dt1.Rows[dt1.Rows.Count - 1]["Z on object"] = z_obj;


                                                if (dt1.Columns.Contains("Pipe Description") == false)
                                                {
                                                    dt1.Columns.Add("Pipe Description", typeof(string));
                                                }

                                                dt1.Rows[dt1.Rows.Count - 1]["Pipe Description"] = Convert.ToString(pipe1.Description);
                                                if (dt1.Columns.Contains("Pipe Style") == false)
                                                {
                                                    dt1.Columns.Add("Pipe Style", typeof(string));
                                                }
                                                dt1.Rows[dt1.Rows.Count - 1]["Pipe Style"] = Convert.ToString(pipe1.StyleName);

                                                if (dt1.Columns.Contains("Pipe Inner Diameter") == false)
                                                {
                                                    dt1.Columns.Add("Pipe Inner Diameter", typeof(double));
                                                }
                                                dt1.Rows[dt1.Rows.Count - 1]["Pipe Inner Diameter"] = diam1;

                                                if (dt1.Columns.Contains("Pipe TOP") == false)
                                                {
                                                    dt1.Columns.Add("Pipe TOP", typeof(double));
                                                }
                                                dt1.Rows[dt1.Rows.Count - 1]["Pipe TOP"] = z_obj + diam1 / 2;

                                                if (dt1.Columns.Contains("Pipe BOP") == false)
                                                {
                                                    dt1.Columns.Add("Pipe BOP", typeof(double));
                                                }
                                                dt1.Rows[dt1.Rows.Count - 1]["Pipe BOP"] = z_obj - diam1 / 2;

                                                add_object_data_to_datatable(dt1, Tables1, pipe1.ObjectId);

                                            }
                                        }



                                        //string msg = "_dwgunits linear = " + dsv.LinearUnit.ToString() ;
                                        //MessageBox.Show(msg);

                                        poly_pipe_3d.Erase();
                                    }


                                }



                            }




                        }


                        Functions.Transfer_datatable_to_new_excel_spreadsheet(dt1, Convert.ToString("int_" + DateTime.Now.Hour + "hr" + DateTime.Now.Minute + "min" + DateTime.Now.Second) + "sec");



                        if (delete_poly3d == true)
                        {
                            poly3d.Erase();
                        }
                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
        }

        private void add_object_data_to_datatable(System.Data.DataTable dt1, Autodesk.Gis.Map.ObjectData.Tables Tables1, ObjectId id1)
        {

            using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tables1.GetObjectRecords(Convert.ToUInt32(0), id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, false))
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
                                if (dt1.Columns.Contains(Nume_field) == false) dt1.Columns.Add(Nume_field, typeof(string));

                                dt1.Rows[dt1.Rows.Count - 1][Nume_field] = Convert.ToString(valoare1);

                            }

                        }
                    }

                }
            }
        }

    }
}
