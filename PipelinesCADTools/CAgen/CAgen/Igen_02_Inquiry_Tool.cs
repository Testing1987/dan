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

namespace Alignment_mdi
{
    /// <summary>
    /// 
    /// </summary>
    public partial class Igen_Inquiry_Tool : Form
    {


        bool Freeze_operations = false;

        System.Data.DataTable Data_table_centerline;



        string project_type = "2d";
        Igen_main_form IG_main = null;


        Point3d picked_pt = new Point3d(123.123, 123.123, 123.123);
        Point3d pt_on_poly = new Point3d(123.123, 123.123, 123.123);



        public Igen_Inquiry_Tool()
        {
            InitializeComponent();

        }


        public bool get_checkBox_temp_cl()
        {
            return checkBox_temp_cl.Checked;
        }
        public bool get_checkBox_temp_sta()
        {
            return checkBox_temp_sta.Checked;
        }

        private void comboBox_clear_at_index_changed(object sender, EventArgs e)
        {

            textBox_offset.Text = "";
            textBox_station.Text = "";
            textBox_mp.Text = "";
            textBox_zoom_to.Text = "";
            textBox_x.Text = "";
            textBox_y.Text = "";
            textBox_z.Text = "";
            textBox_lat.Text = "";
            textBox_long.Text = "";

        }

        private void button_select_centerline_Click(object sender, EventArgs e)
        {
            IG_main = this.MdiParent as Igen_main_form;

            if (Freeze_operations == false)
            {
                Freeze_operations = true;

                ObjectId[] Empty_array = null;

                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                HostApplicationServices.WorkingDatabase = ThisDrawing.Database;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                try
                {

                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {

                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.Database.TransactionManager.StartTransaction())
                        {

                            Data_table_centerline = Functions.Creaza_centerline_datatable_structure();
                            Data_table_centerline.Columns.Add("Bulge", typeof(double));
                            Set_centerline_label_to_red();
                            delete_station_labels();
                            Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_optionsCL = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect Centerline:");
                            Prompt_optionsCL.SetRejectMessage("\nYou did not selected a polyline (2d or 3d)");
                            Prompt_optionsCL.AddAllowedClass(typeof(Polyline), true);
                            Prompt_optionsCL.AddAllowedClass(typeof(Polyline3d), true);
                            IG_main.WindowState = FormWindowState.Minimized;

                            PromptEntityResult Rezultat_CL = Editor1.GetEntity(Prompt_optionsCL);
                            if (Rezultat_CL.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                            {

                                Freeze_operations = false;
                                Editor1.SetImpliedSelection(Empty_array);
                                Editor1.WriteMessage("\nCommand:");
                                IG_main.WindowState = FormWindowState.Normal;
                                return;
                            }



                            Curve Curba1 = Trans1.GetObject(Rezultat_CL.ObjectId, OpenMode.ForRead) as Curve;

                            if (Curba1 == null)
                            {
                                MessageBox.Show("you did not select a polyline or a polyline3d");

                                Freeze_operations = false;
                                Editor1.SetImpliedSelection(Empty_array);
                                Editor1.WriteMessage("\nCommand:");

                                IG_main.WindowState = FormWindowState.Normal;
                                return;
                            }

                            Polyline Poly1 = null;
                            Polyline3d Poly3 = null;

                            if (Curba1 is Polyline)
                            {
                                Poly1 = (Polyline)Curba1;
                                Poly3 = null;
                                project_type = "2d";
                            }

                            else if (Curba1 is Polyline3d)
                            {
                                Poly3 = (Polyline3d)Curba1;
                                Poly1 = Functions.Build_2dpoly_from_3d(Poly3);
                                project_type = "3d";
                            }
                            else
                            {
                                MessageBox.Show("you did not select a polyline or a polyline3d");
                                Freeze_operations = false;
                                Editor1.SetImpliedSelection(Empty_array);
                                Editor1.WriteMessage("\nCommand:");
                                IG_main.WindowState = FormWindowState.Normal;
                                return;
                            }


                            if (checkBox_reverse_direction.Checked == false)
                            {
                                for (int i = 0; i < Poly1.NumberOfVertices; ++i)
                                {
                                    double x2 = Poly1.GetPointAtParameter(i).X;
                                    double y2 = Poly1.GetPointAtParameter(i).Y;
                                    double z2 = Poly1.GetPointAtParameter(i).Z;
                                    if (Poly3 != null)
                                    {
                                        z2 = Poly3.GetPointAtParameter(i).Z;
                                    }

                                    Data_table_centerline.Rows.Add();
                                    Data_table_centerline.Rows[Data_table_centerline.Rows.Count - 1][_AGEN_mainform.Col_x] = x2;
                                    Data_table_centerline.Rows[Data_table_centerline.Rows.Count - 1][_AGEN_mainform.Col_y] = y2;
                                    Data_table_centerline.Rows[Data_table_centerline.Rows.Count - 1][_AGEN_mainform.Col_z] = z2;

                                    Data_table_centerline.Rows[Data_table_centerline.Rows.Count - 1]["Bulge"] = Poly1.GetBulgeAt(i);
                                }
                            }
                            else
                            {
                                for (int i = Poly1.NumberOfVertices - 1; i >= 0; --i)
                                {
                                    double x2 = Poly1.GetPointAtParameter(i).X;
                                    double y2 = Poly1.GetPointAtParameter(i).Y;
                                    double z2 = Poly1.GetPointAtParameter(i).Z;
                                    if (Poly3 != null)
                                    {
                                        z2 = Poly3.GetPointAtParameter(i).Z;
                                    }

                                    Data_table_centerline.Rows.Add();
                                    Data_table_centerline.Rows[Data_table_centerline.Rows.Count - 1][_AGEN_mainform.Col_x] = x2;
                                    Data_table_centerline.Rows[Data_table_centerline.Rows.Count - 1][_AGEN_mainform.Col_y] = y2;
                                    Data_table_centerline.Rows[Data_table_centerline.Rows.Count - 1][_AGEN_mainform.Col_z] = z2;

                                    if (i - 1 >= 0)
                                    {
                                        Data_table_centerline.Rows[Data_table_centerline.Rows.Count - 1]["Bulge"] = -Poly1.GetBulgeAt(i - 1);
                                    }
                                    else
                                    {
                                        Data_table_centerline.Rows[Data_table_centerline.Rows.Count - 1]["Bulge"] = 0;
                                    }

                                }
                            }


                            Trans1.Commit();
                            Set_centerline_label_to_green(Curba1.ObjectId.Handle.Value.ToString());
                        }
                    }

                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            Freeze_operations = false;
            IG_main.WindowState = FormWindowState.Normal;

            Autodesk.Gis.Map.Platform.AcMapMap Acmap = Autodesk.Gis.Map.Platform.AcMapMap.GetCurrentMap();

            string Curent_system = Acmap.GetMapSRS();
            if (Curent_system != "")
            {
                OSGeo.MapGuide.MgCoordinateSystemFactory Coord_factory1 = new OSGeo.MapGuide.MgCoordinateSystemFactory();
                OSGeo.MapGuide.MgCoordinateSystem CoordSys1 = Coord_factory1.Create(Curent_system);
                set_label_cs_to_green(CoordSys1.CsCode);
            }
            else
            {
                set_label_cs_to_red();
            }
        }

        private void button_station_at_point_inquiry_Click(object sender, EventArgs e)
        {

            if (Data_table_centerline == null)
            {
                MessageBox.Show("you did not selected any centerline");
                picked_pt = new Point3d(123.123, 123.123, 123.123);
                pt_on_poly = new Point3d(123.123, 123.123, 123.123);
                return;
            }
            if (Data_table_centerline.Rows.Count == 0)
            {
                MessageBox.Show("you did not selected any centerline");
                picked_pt = new Point3d(123.123, 123.123, 123.123);
                pt_on_poly = new Point3d(123.123, 123.123, 123.123);
                return;
            }

            string start_ammount = textBox_start_station.Text;

            if (Functions.IsNumeric(start_ammount.Replace("+", "")) == false)
            {
                MessageBox.Show("the start station is not specified");
                picked_pt = new Point3d(123.123, 123.123, 123.123);
                pt_on_poly = new Point3d(123.123, 123.123, 123.123);
                return;
            }

            double start1 = Math.Round(Convert.ToDouble(start_ammount.Replace("+", "")), 2);
            delete_zoom_labels();

            if (Freeze_operations == false)
            {
                ObjectId[] Empty_array = null;
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                HostApplicationServices.WorkingDatabase = ThisDrawing.Database;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                try
                {
                    Freeze_operations = true;
                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.Database.TransactionManager.StartTransaction())
                        {
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                            Polyline Poly2D = null;
                            Polyline3d Poly3D = null;
                            if (project_type == "2d")
                            {
                                Poly2D = creeaza_poly2d(Data_table_centerline);

                            }
                            else
                            {
                                Poly3D = Functions.Build_3d_poly_for_scanning(Data_table_centerline);
                                Poly2D = creeaza_poly2d(Data_table_centerline);
                            }

                            IG_main.WindowState = FormWindowState.Minimized;

                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                            Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                            PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify point:");
                            PP1.AllowNone = true;
                            Point_res1 = Editor1.GetPoint(PP1);

                            textBox_offset.Text = "";
                            textBox_station.Text = "";
                            textBox_mp.Text = "";
                            textBox_zoom_to.Text = "";
                            textBox_x.Text = "";
                            textBox_y.Text = "";
                            textBox_z.Text = "";
                            textBox_lat.Text = "";
                            textBox_long.Text = "";

                            if (Point_res1.Status != PromptStatus.OK)
                            {
                                Freeze_operations = false;
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                picked_pt = new Point3d(123.123, 123.123, 123.123);
                                pt_on_poly = new Point3d(123.123, 123.123, 123.123);
                                IG_main.WindowState = FormWindowState.Normal;
                                return;
                            }



                            picked_pt = Point_res1.Value.TransformBy(curent_ucs_matrix);
                            pt_on_poly = Poly2D.GetClosestPointTo(picked_pt, Vector3d.ZAxis, false);
                            double Station1 = Math.Round(start1 + Poly2D.GetDistAtPoint(pt_on_poly), 2);
                            delete_cl(Poly2D);


                            if (project_type == "3d")
                            {
                                double param1 = Poly2D.GetParameterAtPoint(pt_on_poly);
                                Station1 = Math.Round(start1 + Poly3D.GetDistanceAtParameter(param1), 2);
                                delete_cl(Poly3D);
                            }

                            textBox_station.Text = Functions.Get_chainage_from_double(Station1, "f", 2);
                            textBox_mp.Text = Functions.Get_String_Rounded(Station1 / 5280, 2);
                            textBox_offset.Text = Functions.Get_String_Rounded(Math.Round(new Point3d(picked_pt.X, picked_pt.Y, 0).DistanceTo(new Point3d(pt_on_poly.X, pt_on_poly.Y, 0)), 2), 2);
                            textBox_x.Text = Functions.Get_String_Rounded(picked_pt.X, 4);
                            textBox_y.Text = Functions.Get_String_Rounded(picked_pt.Y, 4);
                            textBox_z.Text = Functions.Get_String_Rounded(picked_pt.Z, 4);
                            Point3d LL = Convert_point_to_new_CS(picked_pt, "LL84");

                            textBox_long.Text = Functions.Get_DMS(LL.X, 2);
                            textBox_lat.Text = Functions.Get_DMS(LL.Y, 2);

                            if (checkBox_always_create_label.Checked == true)
                            {
                                Freeze_operations = false;
                                button_create_label_Click(sender, e);
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
                Freeze_operations = false;
                IG_main.WindowState = FormWindowState.Normal;
            }


        }

        private void button_create_label_Click(object sender, EventArgs e)
        {
            if (Data_table_centerline == null)
            {
                MessageBox.Show("you did not selected any centerline");
                return;
            }
            if (Data_table_centerline.Rows.Count == 0)
            {
                MessageBox.Show("you did not selected any centerline");
                return;
            }

            string start_ammount = textBox_start_station.Text;

            if (Functions.IsNumeric(start_ammount.Replace("+", "")) == false)
            {
                MessageBox.Show("the start station is not specified");
                return;
            }


            if (textBox_station.Text == "")
            {
                MessageBox.Show("First you have to specify a point");
                return;
            }



            string sta_ammount = textBox_station.Text;
            string mp_ammount = textBox_mp.Text;
            double sta1 = Convert.ToDouble(sta_ammount.Replace("+", ""));
            double mp1 = Convert.ToDouble(mp_ammount);
            double start1 = Math.Round(Convert.ToDouble(start_ammount.Replace("+", "")), 2);


            if (Freeze_operations == false)
            {

                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                HostApplicationServices.WorkingDatabase = ThisDrawing.Database;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                try
                {
                    Freeze_operations = true;
                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {
                        Functions.Creaza_layer(_AGEN_mainform.layer_no_plot, 30, false);
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.Database.TransactionManager.StartTransaction())
                        {
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                            Polyline Poly2D = creeaza_poly2d(Data_table_centerline);
                            Polyline3d Poly3D = null;

                            if (project_type == "3d")
                            {

                                Poly3D = Functions.Build_3d_poly_for_scanning(Data_table_centerline);

                            }



                            string continut = "";

                            string sta_string = "";
                            string mp_string = "";
                            string offset_string = "";
                            string x_string = "";
                            string y_string = "";
                            string z_string = "";
                            string lat_string = "";
                            string long_string = "";

                            Point3d mleader_ins_pt;



                            double dist1 = new Point3d(picked_pt.X, picked_pt.Y, 0).DistanceTo(new Point3d(pt_on_poly.X, pt_on_poly.Y, 0));

                            offset_string = "Offset=" + Functions.Get_String_Rounded(dist1, 2);

                            if (dist1 > 0.009)
                            {
                                mleader_ins_pt = picked_pt;
                                Line line1 = new Line(new Point3d(picked_pt.X, picked_pt.Y, 0), new Point3d(pt_on_poly.X, pt_on_poly.Y, 0));
                                line1.Layer = _AGEN_mainform.layer_no_plot;
                                BTrecord.AppendEntity(line1);
                                Trans1.AddNewlyCreatedDBObject(line1, true);
                            }
                            else
                            {
                                mleader_ins_pt = pt_on_poly;
                            }

                            if (project_type == "2d")
                            {
                                if (Poly2D.Length < sta1 - start1)
                                {
                                    MessageBox.Show("the station you specified minus the starting station is larger than polyline length ");
                                    Freeze_operations = false;
                                    return;
                                }



                                sta_string = "STA2d=" + Functions.Get_chainage_from_double(sta1, "f", 2);
                                mp_string = "MP2d=" + Functions.Get_String_Rounded(mp1, 2);

                            }
                            else
                            {
                                if (Poly3D.Length < sta1 - start1)
                                {
                                    MessageBox.Show("the station you specified minus the starting station is larger than polyline length ");
                                    Freeze_operations = false;
                                    return;
                                }


                                sta_string = "STA3d=" + Functions.Get_chainage_from_double(sta1, "f", 2);
                                mp_string = "MP3d=" + Functions.Get_String_Rounded(mp1, 2);
                            }


                            x_string = "X=" + textBox_x.Text;
                            y_string = "Y=" + textBox_y.Text;
                            z_string = "Z=" + textBox_z.Text;
                            lat_string = "Lat=" + textBox_lat.Text;
                            long_string = "Long=" + textBox_long.Text;

                            if (checkBox_station.Checked == true)
                            {
                                continut = sta_string;
                            }

                            if (checkBox_mp.Checked == true)
                            {
                                if (continut == "")
                                {
                                    continut = mp_string;
                                }
                                else
                                {
                                    continut = continut + "\r\n" + mp_string;
                                }
                            }

                            if (checkBox_offset.Checked == true)
                            {
                                if (continut == "")
                                {
                                    continut = offset_string;
                                }
                                else
                                {
                                    continut = continut + "\r\n" + offset_string;
                                }
                            }

                            if (checkBox_xyz.Checked == true)
                            {
                                if (continut == "")
                                {
                                    continut = x_string + " " + y_string + " " + z_string;
                                }
                                else
                                {
                                    continut = continut + "\r\n" + x_string + " " + y_string + " " + z_string;
                                }
                            }



                            if (checkBox_ll.Checked == true)
                            {
                                if (continut == "")
                                {
                                    continut = lat_string + " " + long_string;
                                }
                                else
                                {
                                    continut = continut + "\r\n" + lat_string + " " + long_string;
                                }
                            }

                            if (checkBox_custom.Checked == true)
                            {
                                if (textBox_custom.Text != "")
                                {
                                    continut = textBox_custom.Text + "\r\n" + continut;
                                }
                            }


                            MLeader ml1 = Functions.creaza_mleader(mleader_ins_pt, continut, 10, 50, 200, 5, 10, 10);
                            ml1.Layer = _AGEN_mainform.layer_no_plot;


                            zoom_to_Point(mleader_ins_pt, 200);
                            if (project_type == "3d") delete_cl(Poly3D);
                            delete_cl(Poly2D);
                            Trans1.Commit();
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }


                Autodesk.Gis.Map.Platform.AcMapMap Acmap = Autodesk.Gis.Map.Platform.AcMapMap.GetCurrentMap();
                string Curent_system = Acmap.GetMapSRS();
                if (Curent_system != "")
                {
                    OSGeo.MapGuide.MgCoordinateSystemFactory Coord_factory1 = new OSGeo.MapGuide.MgCoordinateSystemFactory();
                    OSGeo.MapGuide.MgCoordinateSystem CoordSys1 = Coord_factory1.Create(Curent_system);
                    set_label_cs_to_green(CoordSys1.CsCode);
                }
                else
                {
                    set_label_cs_to_red();
                }


                delete_zoom_labels();
                Editor1.WriteMessage("\nCommand:");
                Freeze_operations = false;
            }
        }

        private void zoom_to_Point(Point3d pt, double zoom_delta_distance)
        {

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;


            try
            {
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.Database.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);


                        try
                        {



                            Point3d minx = new Point3d(pt.X - zoom_delta_distance, pt.Y - zoom_delta_distance, 0);
                            Point3d maxx = new Point3d(pt.X + zoom_delta_distance, pt.Y + zoom_delta_distance, 0);

                            using (Autodesk.AutoCAD.GraphicsSystem.Manager GraphicsManager = ThisDrawing.GraphicsManager)
                            {

                                int Cvport = Convert.ToInt32(Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("CVPORT"));

                                //from here 2015 dlls:
                                Autodesk.AutoCAD.GraphicsSystem.KernelDescriptor kd = new Autodesk.AutoCAD.GraphicsSystem.KernelDescriptor();
                                kd.addRequirement(Autodesk.AutoCAD.UniqueString.Intern("3D Drawing"));
                                Autodesk.AutoCAD.GraphicsSystem.View view = GraphicsManager.ObtainAcGsView(Cvport, kd);
                                // to here 2015 dlls

                                //from here 2013 dlls:

                                //Autodesk.AutoCAD.GraphicsSystem.View view = GraphicsManager.GetGsView(Cvport, true);

                                // to here 2013 dlls

                                if (view != null)
                                {
                                    using (view)
                                    {

                                        view.ZoomExtents(minx, maxx);

                                        view.Zoom(0.95);//<--optional 
                                        GraphicsManager.SetViewportFromView(Cvport, view, true, true, false);

                                    }
                                }
                                Trans1.Commit();
                            }


                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }

                    }
                }
            }







            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        private void button_zoom_to_Click(object sender, EventArgs e)
        {
            if (Data_table_centerline == null)
            {
                MessageBox.Show("you did not selected any centerline");
                return;
            }
            if (Data_table_centerline.Rows.Count == 0)
            {
                MessageBox.Show("you did not selected any centerline");
                return;
            }

            string start_ammount = textBox_start_station.Text;

            if (Functions.IsNumeric(start_ammount.Replace("+", "")) == false)
            {
                MessageBox.Show("the start station is not specified");
                return;
            }


            string station_ammount = textBox_zoom_to.Text;
            if (Functions.IsNumeric(textBox_zoom_to.Text.Replace("+", "")) == false)
            {
                MessageBox.Show("station is not specified properly");
                return;
            }

            delete_zoom_labels();


            double start1 = Math.Round(Convert.ToDouble(start_ammount.Replace("+", "")), 2);


            double Sta_pt = Convert.ToDouble(station_ammount.Replace("+", ""));


            if (Freeze_operations == false)
            {
                ObjectId[] Empty_array = null;
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                HostApplicationServices.WorkingDatabase = ThisDrawing.Database;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                try
                {
                    Freeze_operations = true;
                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {
                        Functions.Creaza_layer(_AGEN_mainform.layer_no_plot, 30, false);

                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.Database.TransactionManager.StartTransaction())
                        {
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                            Polyline Poly2D = creeaza_poly2d(Data_table_centerline);
                            Polyline3d Poly3D = null;

                            if (project_type == "3d")

                            {
                                Poly3D = Functions.Build_3d_poly_for_scanning(Data_table_centerline);

                            }



                            string continut = "";

                            double Station1 = -123.123;

                            if (comboBox_label_type.Text == "by STA")
                            {
                                Station1 = Sta_pt;

                                if (project_type == "2d")
                                {
                                    if (Poly2D.Length < Sta_pt - start1)
                                    {
                                        MessageBox.Show("the station you specified minus the starting station is larger than polyline length ");
                                        Freeze_operations = false;
                                        return;
                                    }
                                    pt_on_poly = Poly2D.GetPointAtDist(Sta_pt - start1);
                                    continut = "STA2d=" + Functions.Get_chainage_from_double(Sta_pt, "f", 2) + "\r\nMP2d=" + Functions.Get_String_Rounded(Sta_pt / 5280, 2);

                                }
                                else
                                {
                                    if (Poly3D.Length < Sta_pt - start1)
                                    {
                                        MessageBox.Show("the station you specified minus the starting station is larger than polyline length ");
                                        Freeze_operations = false;
                                        return;
                                    }

                                    pt_on_poly = Poly3D.GetPointAtDist(Sta_pt - start1);
                                    continut = "STA3d=" + Functions.Get_chainage_from_double(Sta_pt, "f", 2) + "\r\nMP3d=" + Functions.Get_String_Rounded(Sta_pt / 5280, 2);
                                }
                            }
                            else
                            {
                                Station1 = Sta_pt * 5280;

                                if (project_type == "2d")
                                {
                                    if (Poly2D.Length < Sta_pt * 5280 - start1)
                                    {
                                        MessageBox.Show("the mp you specified minus the starting station is larger than polyline length ");
                                        Freeze_operations = false;
                                        return;
                                    }
                                    pt_on_poly = Poly2D.GetPointAtDist(Sta_pt * 5280 - start1);
                                    continut = "STA2d=" + Functions.Get_chainage_from_double(Sta_pt * 5280, "f", 2) + "\r\nMP2d=" + Functions.Get_String_Rounded(Sta_pt, 2);
                                }
                                else
                                {
                                    if (Poly3D.Length < Sta_pt * 5280 - start1)
                                    {
                                        MessageBox.Show("the mp you specified minus the starting station is larger than polyline length ");
                                        Freeze_operations = false;
                                        return;
                                    }
                                    pt_on_poly = Poly3D.GetPointAtDist(Sta_pt * 5280 - start1);
                                    continut = "STA3d=" + Functions.Get_chainage_from_double(Sta_pt * 5280, "f", 2) + "\r\nMP3d=" + Functions.Get_String_Rounded(Sta_pt, 2);
                                }
                            }

                            textBox_station.Text = Functions.Get_chainage_from_double(Station1, "f", 2);
                            textBox_mp.Text = Functions.Get_String_Rounded(Station1 / 5280, 2);
                            textBox_offset.Text = "0";
                            textBox_zoom_to.Text = "";


                            textBox_x.Text = Functions.Get_String_Rounded(pt_on_poly.X, 4);
                            textBox_y.Text = Functions.Get_String_Rounded(pt_on_poly.Y, 4);
                            textBox_z.Text = Functions.Get_String_Rounded(pt_on_poly.Z, 4);
                            Point3d LL = Convert_point_to_new_CS(pt_on_poly, "LL84");

                            textBox_long.Text = Functions.Get_DMS(LL.X, 2);
                            textBox_lat.Text = Functions.Get_DMS(LL.Y, 2);



                            MText mt1 = Functions.creaza_mtext_label(pt_on_poly, continut, 0.1);
                            mt1.Layer = _AGEN_mainform.layer_no_plot;
                            BTrecord.AppendEntity(mt1);
                            Trans1.AddNewlyCreatedDBObject(mt1, true);
                            Igen_main_form.col_labels_zoom.Add(mt1.ObjectId.Handle.Value.ToString());

                            zoom_to_Point(pt_on_poly, 40);
                            picked_pt = pt_on_poly;
                            if (project_type == "3d") delete_cl(Poly3D);
                            delete_cl(Poly2D);
                            Trans1.Commit();
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                Autodesk.Gis.Map.Platform.AcMapMap Acmap = Autodesk.Gis.Map.Platform.AcMapMap.GetCurrentMap();

                string Curent_system = Acmap.GetMapSRS();
                if (Curent_system != "")
                {
                    OSGeo.MapGuide.MgCoordinateSystemFactory Coord_factory1 = new OSGeo.MapGuide.MgCoordinateSystemFactory();
                    OSGeo.MapGuide.MgCoordinateSystem CoordSys1 = Coord_factory1.Create(Curent_system);
                    set_label_cs_to_green(CoordSys1.CsCode);
                }
                else
                {
                    set_label_cs_to_red();
                }


                Editor1.SetImpliedSelection(Empty_array);
                Editor1.WriteMessage("\nCommand:");
                Freeze_operations = false;
            }
        }






        private void Set_centerline_label_to_red()
        {
            label_cl_loaded.Text = "CL not loaded";
            label_cl_loaded.ForeColor = Color.Red;
            label_sta.Text = "2D Station:";
            label_mp.Text = "MP(2D):";
            textBox_offset.Text = "";
            textBox_station.Text = "";
            textBox_mp.Text = "";
            textBox_zoom_to.Text = "";
            textBox_x.Text = "";
            textBox_y.Text = "";
            textBox_z.Text = "";
            textBox_lat.Text = "";
            textBox_long.Text = "";
        }

        private void Set_centerline_label_to_green(string handle1)
        {
            label_cl_loaded.Text = "CL loaded \r\nhandle# - " + handle1;
            label_cl_loaded.ForeColor = Color.LimeGreen;

            if (project_type == "2d")
            {
                label_sta.Text = "2D Station:";
                label_mp.Text = "MP(2D):";
            }
            else
            {
                label_sta.Text = "3D Station:";
                label_mp.Text = "MP(3D):";
            }

            textBox_offset.Text = "";
            textBox_station.Text = "";
            textBox_mp.Text = "";
            textBox_zoom_to.Text = "";
            textBox_x.Text = "";
            textBox_y.Text = "";
            textBox_z.Text = "";
            textBox_lat.Text = "";
            textBox_long.Text = "";

        }

        private Polyline creeaza_poly2d(System.Data.DataTable dt_cl)
        {
            Polyline Poly2D = new Polyline();

            if (dt_cl.Rows.Count > 0)
            {

                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;



                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {

                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;


                        int index1 = 0;

                        for (int i = 0; i < dt_cl.Rows.Count; ++i)
                        {
                            double x = 0;
                            double y = 0;

                            if (dt_cl.Rows[i][Alignment_mdi._AGEN_mainform.Col_x] != DBNull.Value)
                            {
                                x = (double)dt_cl.Rows[i][Alignment_mdi._AGEN_mainform.Col_x];
                                if (dt_cl.Rows[i][Alignment_mdi._AGEN_mainform.Col_y] != DBNull.Value)
                                {
                                    y = (double)dt_cl.Rows[i][Alignment_mdi._AGEN_mainform.Col_y];

                                    double bulge1 = 0;
                                    if (dt_cl.Rows[i][Alignment_mdi._AGEN_mainform.Col_MMid] != DBNull.Value)
                                    {
                                        string blg = Convert.ToString(dt_cl.Rows[i][Alignment_mdi._AGEN_mainform.Col_MMid]);
                                        if (Functions.IsNumeric(blg) == true)
                                        {
                                            bulge1 = Convert.ToDouble(blg);
                                        }
                                    }


                                    Poly2D.AddVertexAt(index1, new Point2d(x, y), bulge1, 0, 0);
                                    Poly2D.Elevation = 0;

                                    index1 = index1 + 1;
                                }
                            }
                        }

                        BTrecord.AppendEntity(Poly2D);
                        Trans1.AddNewlyCreatedDBObject(Poly2D, true);

                        Trans1.Commit();

                    }
                }

            }


            return Poly2D;


        }

        private void delete_cl(Entity Poly1)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock lock1 = ThisDrawing.LockDocument())
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {

                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                    Entity ent1 = Trans1.GetObject(Poly1.ObjectId, OpenMode.ForWrite) as Entity;
                    if (ent1 != null)
                    {
                        ent1.Erase();
                        Trans1.Commit();
                    }
                }
            }
        }
        public void delete_cl_from_redraw(string handle1)
        {
            if (handle1 != null)
            {

                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {

                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                        ObjectId id1 = Functions.GetObjectId(ThisDrawing.Database, handle1);
                        if (id1 != ObjectId.Null)
                        {
                            try
                            {
                                Entity ent1 = Trans1.GetObject(id1, OpenMode.ForWrite) as Entity;
                                if (ent1 != null)
                                {
                                    ent1.Erase();
                                    Trans1.Commit();
                                }
                            }
                            catch (System.Exception ex)
                            {
                            }
                        }

                    }

                }
            }
        }


        private void button_redraw_centerline_Click(object sender, EventArgs e)
        {
            if (Freeze_operations == false)
            {
                Freeze_operations = true;

                if (Data_table_centerline != null)
                {
                    if (Data_table_centerline.Rows.Count > 0)
                    {
                        if (checkBox_reverse_direction.Checked == true)
                        {
                            Data_table_centerline = reverse_cl();
                        }

                        try
                        {
                            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                            using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                            {
                                Functions.Creaza_layer(_AGEN_mainform.layer_no_plot, 30, false);
                                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                                {
                                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                                    if (project_type == "2d")
                                    {
                                        Polyline new_poly2d = new Polyline();
                                        for (int i = 0; i < Data_table_centerline.Rows.Count; ++i)
                                        {
                                            double x = 0;
                                            double y = 0;
                                            double bulge1 = 0;

                                            if (Data_table_centerline.Rows[i][_AGEN_mainform.Col_x] != DBNull.Value)
                                            {
                                                x = (double)Data_table_centerline.Rows[i][_AGEN_mainform.Col_x];
                                            }
                                            else
                                            {

                                                Freeze_operations = false;
                                                MessageBox.Show("no X value for centerline in row " + (i).ToString());
                                                return;
                                            }
                                            if (Data_table_centerline.Rows[i][_AGEN_mainform.Col_y] != DBNull.Value)
                                            {
                                                y = (double)Data_table_centerline.Rows[i][_AGEN_mainform.Col_y];
                                            }
                                            else
                                            {

                                                Freeze_operations = false;
                                                MessageBox.Show("no Y value for centerline in row " + (i).ToString());
                                                return;
                                            }
                                     
                                            if (Data_table_centerline.Rows[i][Alignment_mdi._AGEN_mainform.Col_MMid] != DBNull.Value)
                                            {
                                                string blg = Convert.ToString(Data_table_centerline.Rows[i][Alignment_mdi._AGEN_mainform.Col_MMid]);
                                                if (Functions.IsNumeric(blg) == true)
                                                {
                                                    bulge1 = Convert.ToDouble(blg);
                                                }
                                            }
                                            new_poly2d.AddVertexAt(i, new Point2d(x, y), bulge1, 0, 0);

                                        }

                                        new_poly2d.Layer = _AGEN_mainform.layer_no_plot;
                                        BTrecord.AppendEntity(new_poly2d);
                                        Trans1.AddNewlyCreatedDBObject(new_poly2d, true);

                                        Igen_main_form.cl_id_for_temp = new_poly2d.ObjectId.Handle.Value.ToString();

                                    }
                                    else
                                    {
                                        Polyline3d new_poly3d = new Polyline3d();
                                        new_poly3d.SetDatabaseDefaults();
                                        new_poly3d.Layer = _AGEN_mainform.layer_no_plot;
                                        BTrecord.AppendEntity(new_poly3d);
                                        Trans1.AddNewlyCreatedDBObject(new_poly3d, true);

                                        for (int i = 0; i < Data_table_centerline.Rows.Count; ++i)
                                        {
                                            double x = 0;
                                            double y = 0;
                                            double z = 0;

                                            if (Data_table_centerline.Rows[i][_AGEN_mainform.Col_x] != DBNull.Value)
                                            {
                                                x = (double)Data_table_centerline.Rows[i][_AGEN_mainform.Col_x];
                                            }
                                            else
                                            {

                                                Freeze_operations = false;
                                                MessageBox.Show("no X value for centerline in row " + (i).ToString());
                                                return;
                                            }
                                            if (Data_table_centerline.Rows[i][_AGEN_mainform.Col_y] != DBNull.Value)
                                            {
                                                y = (double)Data_table_centerline.Rows[i][_AGEN_mainform.Col_y];
                                            }
                                            else
                                            {

                                                Freeze_operations = false;
                                                MessageBox.Show("no Y value for centerline in row " + (i).ToString());
                                                return;
                                            }
                                            if (Data_table_centerline.Rows[i][_AGEN_mainform.Col_z] != DBNull.Value)
                                            {
                                                z = (double)Data_table_centerline.Rows[i][_AGEN_mainform.Col_z];
                                            }
                                            else
                                            {

                                                Freeze_operations = false;
                                                MessageBox.Show("no Y value for centerline in row " + (i).ToString());
                                                return;
                                            }

                                            PolylineVertex3d Vertex_new = new PolylineVertex3d(new Point3d(x, y, z));
                                            new_poly3d.AppendVertex(Vertex_new);
                                            Trans1.AddNewlyCreatedDBObject(Vertex_new, true);

                                        }

                                        Igen_main_form.cl_id_for_temp = new_poly3d.ObjectId.Handle.Value.ToString();

                                    }
                                    Trans1.Commit();

                                    if (checkBox_reverse_direction.Checked == true)
                                    {
                                        Data_table_centerline = reverse_cl();
                                    }



                                }
                            }
                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }

                    Freeze_operations = false;
                }
            }
        }

        private void button_create_station_label_Click(object sender, EventArgs e)
        {

            double spacing_major = 500;
            double spacing_minor = 100;
            double tick_major = 20;
            double tick_minor = 10;
            double texth = 4;
            double gap1 = 4;
            string start_ammount = textBox_start_station.Text;

            if (Functions.IsNumeric(start_ammount.Replace("+", "")) == false)
            {
                MessageBox.Show("the start station is not specified");
                return;
            }

            double start1 = Math.Round(Convert.ToDouble(start_ammount.Replace("+", "")), 2);

            if (Freeze_operations == false)
            {
                Freeze_operations = true;

                if (Data_table_centerline != null)
                {
                    if (Data_table_centerline.Rows.Count > 0)
                    {
                        try
                        {
                            delete_station_labels();

                            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                            using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                            {
                                Functions.Creaza_layer(_AGEN_mainform.layer_no_plot, 30, false);
                                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                                {
                                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                                    Polyline Poly2D = creeaza_poly2d(Data_table_centerline);
                                    Polyline3d Poly3D = null;

                                    if (project_type == "3d") Poly3D = Functions.Build_3d_poly_for_scanning(Data_table_centerline);

                                    if (project_type == "2d")
                                    {
                                        if (Poly2D.Length >= spacing_major)
                                        {
                                            int no_major = Convert.ToInt32(Math.Floor((Poly2D.Length) / spacing_major)) + 2;



                                            double first_label_major = spacing_major * Math.Ceiling(start1 / spacing_major);
                                            double len_stationed_major = Poly2D.Length - (first_label_major - start1);


                                            do
                                            {
                                                if (no_major * spacing_major >= len_stationed_major)
                                                {
                                                    no_major = no_major - 1;
                                                }
                                            } while (no_major * spacing_major >= len_stationed_major);



                                            if (no_major > 0)
                                            {
                                                for (int i = 0; i <= no_major; ++i)
                                                {
                                                    Point3d pt0 = Poly2D.GetPointAtDist((first_label_major - start1) + i * spacing_major);
                                                    double label_major = first_label_major + i * spacing_major;
                                                    Line Big1 = new Line(new Point3d(pt0.X - tick_major / 2, pt0.Y, 0), new Point3d(pt0.X + tick_major / 2, pt0.Y, 0));

                                                    double param1 = Poly2D.GetParameterAtPoint(Poly2D.GetClosestPointTo(pt0, Vector3d.ZAxis, false));
                                                    double param2 = param1 + 1;
                                                    if (Poly2D.EndParam < param2)
                                                    {
                                                        param1 = Poly2D.EndParam - 1;
                                                        param2 = Poly2D.EndParam;
                                                    }


                                                    Point3d point1 = Poly2D.GetPointAtParameter(Math.Floor(param1));

                                                    Point3d point2 = Poly2D.GetPointAtParameter(Math.Floor(param2));

                                                    double bear1 = Functions.GET_Bearing_rad(point1.X, point1.Y, point2.X, point2.Y);

                                                    double rot1 = bear1 - Math.PI / 2;



                                                    Big1.TransformBy(Matrix3d.Rotation(rot1, Vector3d.ZAxis, pt0));

                                                    Big1.Layer = _AGEN_mainform.layer_no_plot;
                                                    Big1.ColorIndex = 256;
                                                    BTrecord.AppendEntity(Big1);
                                                    Trans1.AddNewlyCreatedDBObject(Big1, true);

                                                    Igen_main_form.col_station_labels.Add(Big1.ObjectId.Handle.Value.ToString());

                                                    Line l_t = new Line(Big1.StartPoint, Big1.EndPoint);
                                                    l_t.TransformBy(Matrix3d.Scaling((Big1.Length + gap1) / Big1.Length, Big1.EndPoint));

                                                    MText mt1 = creaza_mtext_sta(l_t.StartPoint, Functions.Get_chainage_from_double(first_label_major + i * spacing_major, "f", 0), texth, bear1);

                                                    mt1.Layer = _AGEN_mainform.layer_no_plot;
                                                    BTrecord.AppendEntity(mt1);
                                                    Trans1.AddNewlyCreatedDBObject(mt1, true);
                                                    Igen_main_form.col_station_labels.Add(mt1.ObjectId.Handle.Value.ToString());
                                                }
                                            }
                                        }

                                        if (Poly2D.Length >= spacing_minor)
                                        {

                                            int no_minor = Convert.ToInt32(Math.Floor((Poly2D.Length) / spacing_minor)) + 2;




                                            double first_label_minor = spacing_minor * Math.Ceiling(start1 / spacing_minor);
                                            double len_stationed_minor = Poly2D.Length - (first_label_minor - start1);



                                            do
                                            {
                                                if (no_minor * spacing_minor >= len_stationed_minor)
                                                {
                                                    no_minor = no_minor - 1;
                                                }
                                            } while (no_minor * spacing_minor >= len_stationed_minor);


                                            if (no_minor > 0)
                                            {
                                                for (int i = 0; i <= no_minor; ++i)
                                                {
                                                    Point3d pt0 = Poly2D.GetPointAtDist((first_label_minor - start1) + i * spacing_minor);
                                                    double label_major = first_label_minor + i * spacing_minor;
                                                    Line small1 = new Line(new Point3d(pt0.X - tick_minor / 2, pt0.Y, 0), new Point3d(pt0.X + tick_minor / 2, pt0.Y, 0));

                                                    double param1 = Poly2D.GetParameterAtPoint(Poly2D.GetClosestPointTo(pt0, Vector3d.ZAxis, false));
                                                    double param2 = param1 + 1;
                                                    if (Poly2D.EndParam < param2)
                                                    {
                                                        param1 = Poly2D.EndParam - 1;
                                                        param2 = Poly2D.EndParam;
                                                    }


                                                    Point3d point1 = Poly2D.GetPointAtParameter(Math.Floor(param1));

                                                    Point3d point2 = Poly2D.GetPointAtParameter(Math.Floor(param2));

                                                    double bear1 = Functions.GET_Bearing_rad(point1.X, point1.Y, point2.X, point2.Y);

                                                    double rot1 = bear1 - Math.PI / 2;

                                                    small1.TransformBy(Matrix3d.Rotation(rot1, Vector3d.ZAxis, pt0));

                                                    small1.Layer = _AGEN_mainform.layer_no_plot;
                                                    small1.ColorIndex = 256;
                                                    BTrecord.AppendEntity(small1);
                                                    Trans1.AddNewlyCreatedDBObject(small1, true);

                                                    Igen_main_form.col_station_labels.Add(small1.ObjectId.Handle.Value.ToString());

                                                }

                                            }

                                        }

                                    }
                                    if (project_type == "3d")
                                    {
                                        if (Poly3D.Length >= spacing_major)
                                        {
                                            int no_major = Convert.ToInt32(Math.Floor((Poly3D.Length) / spacing_major)) + 2;


                                            double first_label_major = spacing_major * Math.Ceiling(start1 / spacing_major);
                                            double len_stationed_major = Poly3D.Length - (first_label_major - start1);



                                            do
                                            {
                                                if (no_major * spacing_major >= len_stationed_major)
                                                {
                                                    no_major = no_major - 1;
                                                }
                                            } while (no_major * spacing_major >= len_stationed_major);




                                            if (no_major > 0)
                                            {
                                                for (int i = 0; i <= no_major; ++i)
                                                {
                                                    Point3d pt0 = Poly3D.GetPointAtDist((first_label_major - start1) + i * spacing_major);


                                                    double label_major = first_label_major + i * spacing_major;
                                                    Line Big1 = new Line(new Point3d(pt0.X - tick_major / 2, pt0.Y, 0), new Point3d(pt0.X + tick_major / 2, pt0.Y, 0));

                                                    double param1 = Poly2D.GetParameterAtPoint(Poly2D.GetClosestPointTo(pt0, Vector3d.ZAxis, false));
                                                    double param2 = param1 + 1;
                                                    if (Poly2D.EndParam < param2)
                                                    {
                                                        param1 = Poly2D.EndParam - 1;
                                                        param2 = Poly2D.EndParam;
                                                    }


                                                    Point3d point1 = Poly3D.GetPointAtParameter(Math.Floor(param1));

                                                    Point3d point2 = Poly3D.GetPointAtParameter(Math.Floor(param2));

                                                    double bear1 = Functions.GET_Bearing_rad(point1.X, point1.Y, point2.X, point2.Y);

                                                    double rot1 = bear1 - Math.PI / 2;

                                                    Big1.TransformBy(Matrix3d.Rotation(rot1, Vector3d.ZAxis, pt0));

                                                    Big1.Layer = _AGEN_mainform.layer_no_plot;
                                                    Big1.ColorIndex = 256;
                                                    BTrecord.AppendEntity(Big1);
                                                    Trans1.AddNewlyCreatedDBObject(Big1, true);

                                                    Igen_main_form.col_station_labels.Add(Big1.ObjectId.Handle.Value.ToString());

                                                    Line l_t = new Line(Big1.StartPoint, Big1.EndPoint);
                                                    l_t.TransformBy(Matrix3d.Scaling((Big1.Length + gap1) / Big1.Length, Big1.EndPoint));

                                                    MText mt1 = creaza_mtext_sta(l_t.StartPoint, Functions.Get_chainage_from_double(first_label_major + i * spacing_major, "f", 0), texth, bear1);

                                                    mt1.Layer = _AGEN_mainform.layer_no_plot;
                                                    BTrecord.AppendEntity(mt1);
                                                    Trans1.AddNewlyCreatedDBObject(mt1, true);
                                                    Igen_main_form.col_station_labels.Add(mt1.ObjectId.Handle.Value.ToString());

                                                }
                                            }
                                        }

                                        if (Poly3D.Length >= spacing_minor)
                                        {
                                            int no_minor = Convert.ToInt32(Math.Floor((Poly3D.Length) / spacing_minor)) + 2;


                                            double first_label_minor = spacing_minor * Math.Ceiling(start1 / spacing_minor);
                                            double len_stationed_minor = Poly3D.Length - (first_label_minor - start1);

                                            do
                                            {
                                                if (no_minor * spacing_minor >= len_stationed_minor)
                                                {
                                                    no_minor = no_minor - 1;
                                                }
                                            } while (no_minor * spacing_minor >= len_stationed_minor);


                                            if (no_minor > 0)
                                            {
                                                for (int i = 0; i <= no_minor; ++i)
                                                {
                                                    Point3d pt0 = Poly3D.GetPointAtDist((first_label_minor - start1) + i * spacing_minor);
                                                    double label_major = first_label_minor + i * spacing_minor;
                                                    Line small1 = new Line(new Point3d(pt0.X - tick_minor / 2, pt0.Y, 0), new Point3d(pt0.X + tick_minor / 2, pt0.Y, 0));

                                                    double param1 = Poly2D.GetParameterAtPoint(Poly2D.GetClosestPointTo(pt0, Vector3d.ZAxis, false));
                                                    double param2 = param1 + 1;
                                                    if (Poly2D.EndParam < param2)
                                                    {
                                                        param1 = Poly2D.EndParam - 1;
                                                        param2 = Poly2D.EndParam;
                                                    }


                                                    Point3d point1 = Poly3D.GetPointAtParameter(Math.Floor(param1));

                                                    Point3d point2 = Poly3D.GetPointAtParameter(Math.Floor(param2));

                                                    double bear1 = Functions.GET_Bearing_rad(point1.X, point1.Y, point2.X, point2.Y);

                                                    double rot1 = bear1 - Math.PI / 2;

                                                    small1.TransformBy(Matrix3d.Rotation(rot1, Vector3d.ZAxis, pt0));

                                                    small1.Layer = _AGEN_mainform.layer_no_plot;
                                                    small1.ColorIndex = 256;
                                                    BTrecord.AppendEntity(small1);
                                                    Trans1.AddNewlyCreatedDBObject(small1, true);

                                                    Igen_main_form.col_station_labels.Add(small1.ObjectId.Handle.Value.ToString());
                                                }
                                            }
                                        }

                                    }

                                    if (project_type == "3d") delete_cl(Poly3D);
                                    delete_cl(Poly2D);
                                    Trans1.Commit();
                                }
                            }
                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }

                    Freeze_operations = false;
                }
            }
        }

        public static MText creaza_mtext_sta(Point3d pt_ins, string continut, double texth, double rot1)
        {


            MText mtext1 = new MText();
            mtext1.Attachment = AttachmentPoint.BottomCenter;
            mtext1.Contents = continut;
            mtext1.TextHeight = texth;
            mtext1.BackgroundFill = true;
            mtext1.UseBackgroundColor = true;
            mtext1.BackgroundScaleFactor = 1.2;
            mtext1.Location = pt_ins;
            mtext1.Rotation = rot1;
            mtext1.ColorIndex = 256;


            return mtext1;


        }


        public void delete_zoom_labels()
        {
            if (Igen_main_form.col_labels_zoom.Count > 0)
            {
                using (ObjectIdCollection col1 = new ObjectIdCollection())
                {

                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.Database.TransactionManager.StartTransaction())
                        {
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                            foreach (string handle1 in Igen_main_form.col_labels_zoom)
                            {
                                ObjectId id1 = Functions.GetObjectId(ThisDrawing.Database, handle1);
                                if (id1 != ObjectId.Null)
                                {
                                    try
                                    {
                                        MText mt_zoom = Trans1.GetObject(id1, OpenMode.ForRead) as MText;
                                        if (mt_zoom != null)
                                        {
                                            if (mt_zoom.Layer == _AGEN_mainform.layer_no_plot)
                                            {
                                                if (mt_zoom.TextHeight == 0.1)
                                                {
                                                    mt_zoom.UpgradeOpen();
                                                    mt_zoom.Erase();
                                                    col1.Add(id1);
                                                }
                                            }
                                        }
                                    }
                                    catch (System.Exception ex)
                                    {
                                    }
                                }
                            }
                            Trans1.Commit();
                        }
                    }

                    if (col1.Count > 0)
                    {
                        foreach (ObjectId id1 in col1)
                        {
                            string handle1 = id1.Handle.Value.ToString();
                            if (Igen_main_form.col_labels_zoom.Contains(handle1) == true)
                            {
                                Igen_main_form.col_labels_zoom.Remove(handle1);
                            }
                        }
                    }
                }
            }
        }


        public void delete_station_labels()
        {
            if (Igen_main_form.col_station_labels.Count > 0)
            {
                using (ObjectIdCollection col1 = new ObjectIdCollection())
                {

                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.Database.TransactionManager.StartTransaction())
                        {
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                            foreach (string handle1 in Igen_main_form.col_station_labels)
                            {
                                ObjectId id1 = Functions.GetObjectId(ThisDrawing.Database, handle1);
                                if (id1 != ObjectId.Null)
                                {
                                    try
                                    {
                                        Entity ent1 = Trans1.GetObject(id1, OpenMode.ForRead) as Entity;
                                        if (ent1 != null)
                                        {

                                            ent1.UpgradeOpen();
                                            ent1.Erase();
                                            col1.Add(id1);


                                        }
                                    }
                                    catch (System.Exception ex)
                                    {
                                    }
                                }
                            }
                            Trans1.Commit();
                        }
                    }

                    if (col1.Count > 0)
                    {
                        foreach (ObjectId id1 in col1)
                        {
                            string handle1 = id1.Handle.Value.ToString();
                            if (Igen_main_form.col_station_labels.Contains(handle1) == true)
                            {
                                Igen_main_form.col_station_labels.Remove(handle1);
                            }
                        }
                    }
                }
            }
        }

        public Point3d Convert_point_to_new_CS(Point3d Point1, string to_coord_system)
        {
            Autodesk.Gis.Map.Platform.AcMapMap Acmap = Autodesk.Gis.Map.Platform.AcMapMap.GetCurrentMap();

            string Curent_system = Acmap.GetMapSRS();
            if (Curent_system == "")
            {
                set_label_cs_to_red();
                return new Point3d();
            }

            Point3d Point2 = new Point3d();
            OSGeo.MapGuide.MgCoordinateSystemFactory Coord_factory1 = new OSGeo.MapGuide.MgCoordinateSystemFactory();
            OSGeo.MapGuide.MgCoordinateSystemCatalog Catalog1 = Coord_factory1.GetCatalog();
            OSGeo.MapGuide.MgCoordinateSystemDictionary Dictionary1 = Catalog1.GetCoordinateSystemDictionary();
            OSGeo.MapGuide.MgCoordinateSystemEnum Enum1 = Dictionary1.GetEnum();

            OSGeo.MapGuide.MgCoordinateSystem CoordSys1 = Coord_factory1.Create(Curent_system);

            OSGeo.MapGuide.MgCoordinateSystem CoordSys2 = Dictionary1.GetCoordinateSystem(to_coord_system);

            OSGeo.MapGuide.MgCoordinateSystemTransform Transform1 = Coord_factory1.GetTransform(CoordSys1, CoordSys2);
            OSGeo.MapGuide.MgCoordinate Coord1 = Transform1.Transform(Point1.X, Point1.Y);

            Point2 = new Point3d(Coord1.X, Coord1.Y, 0);

            set_label_cs_to_green(CoordSys1.CsCode);

            return Point2;
        }

        public void set_label_cs_to_red()
        {
            label_cs.Text = "NO coordinate system set";
            label_cs.ForeColor = Color.Red;
        }

        public void set_label_cs_to_green(string cs_name)
        {
            label_cs.Text = cs_name;
            label_cs.ForeColor = Color.LimeGreen;
        }

        private void checkBox_reverse_direction_CheckedChanged(object sender, EventArgs e)
        {
            Data_table_centerline = reverse_cl();

            button_create_station_label_Click(sender, e);

        }


        private System.Data.DataTable reverse_cl()
        {
            System.Data.DataTable dt1 = Functions.Creaza_centerline_datatable_structure();
            dt1.Columns.Add("Bulge", typeof(double));

            if (Data_table_centerline != null)
            {
                if (Data_table_centerline.Rows.Count > 0)
                {
                    for (int i = Data_table_centerline.Rows.Count - 1; i >= 0; --i)
                    {
                        dt1.Rows.Add();
                        for (int j = 0; j < Data_table_centerline.Columns.Count; ++j)
                        {
                            dt1.Rows[dt1.Rows.Count - 1][j] = Data_table_centerline.Rows[i][j];
                        }
                    }

                }
            }

            return dt1;
        }
    }
}
