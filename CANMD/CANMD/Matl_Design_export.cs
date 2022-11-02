using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.Gis.Map;
using Autodesk.Gis.Map.ImportExport;
using Autodesk.Gis.Map.ObjectData;
using Autodesk.Gis.Map.Project;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace Alignment_mdi
{
    public partial class Export_form : Form
    {
        //Global Variables

        public static Export_form tpage_export = null;

        string pipes_layer = "Pipes";
        string pipes_od = "Pipes";
        string buoyancy_layer = "Buoyancy";
        string buoyancy_od = "Buoyancy";
        string cpac_layer = "CPAC";
        string cpac_od = "CPAC";
        string es_layer = "EandS";
        string es_od = "EandS";

        string xing_layer = "Crossing";
        string xing_od = "Crossing";
        string transition_layer = "Transition";
        string transition_od = "Transition";
        string elbows_layer = "Elbows";
        string elbows_od = "Elbows";
        string fab_layer = "Assembly";
        string fab_od = "Assembly";
        string class_layer = "Class";
        string class_od = "Class";
        string buoyancy_pt_layer = "BuoyancyPT";
        string buoyancy_pt_od = "BuoyancyPT";
        string geotech_layer = "Geohazard";
        string geotech_od = "Geohazard";
        string geotech_layer_pt = "GeohazardPT";
        string geotech_od_pt = "GeohazardPT";
        string muskeg_layer = "Musk_drainage";
        string muskeg_od = "Musk_drainage";
        string doc_layer = "Depth_of_Cover";
        string doc_od = "Depth_of_Cover";

        string hydrotest_layerPT = "HydrostaticTestPT";
        string hydrotest_odPT = "HydrostaticTestPT";

        string hydrotest_layer = "HydrostaticTest";
        string hydrotest_od = "HydrostaticTest";

        string preexisting_layer = "Pre_existing_pipe";
        string preexisting_od = "Pre_existing_pipe";

        public Export_form()
        {
            InitializeComponent();
            tpage_export = this;

        }





        #region set enable true or false    
        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();

            lista_butoane.Add(button_export);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = false;
            }
        }
        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_export);


            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }
        #endregion

        private void button_export_Click(object sender, EventArgs e)
        {







            string dt99_layer = "polygon";
            string dt99_od = "polygon";


            System.Data.DataTable dt1 = new System.Data.DataTable();
            dt1.Columns.Add("id", typeof(ObjectId));
            dt1.Columns.Add("layer", typeof(string));

            System.Data.DataTable dt2 = new System.Data.DataTable();
            dt2 = dt1.Clone();

            System.Data.DataTable dt3 = new System.Data.DataTable();
            dt3 = dt1.Clone();

            System.Data.DataTable dt4 = new System.Data.DataTable();
            dt4 = dt1.Clone();

            System.Data.DataTable dt5 = new System.Data.DataTable();
            dt5 = dt1.Clone();

            System.Data.DataTable dt6 = new System.Data.DataTable();
            dt6 = dt1.Clone();

            System.Data.DataTable dt7 = new System.Data.DataTable();
            dt7 = dt1.Clone();

            System.Data.DataTable dt8 = new System.Data.DataTable();
            dt8 = dt1.Clone();

            System.Data.DataTable dt9 = new System.Data.DataTable();
            dt9 = dt1.Clone();
            System.Data.DataTable dt10 = new System.Data.DataTable();
            dt10 = dt1.Clone();
            System.Data.DataTable dt11 = new System.Data.DataTable();
            dt11 = dt1.Clone();
            System.Data.DataTable dt12 = new System.Data.DataTable();
            dt12 = dt1.Clone();
            System.Data.DataTable dt13 = new System.Data.DataTable();
            dt13 = dt1.Clone();
            System.Data.DataTable dt14 = new System.Data.DataTable();
            dt14 = dt1.Clone();
            System.Data.DataTable dt15 = new System.Data.DataTable();
            dt15 = dt1.Clone();

            System.Data.DataTable dt16 = new System.Data.DataTable();
            dt16 = dt1.Clone();

            System.Data.DataTable dt17 = new System.Data.DataTable();
            dt17 = dt1.Clone();


            System.Data.DataTable dt99 = new System.Data.DataTable();
            dt99 = dt1.Clone();


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
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;
                        Autodesk.Gis.Map.MapApplication mapApp = Autodesk.Gis.Map.HostMapApplicationServices.Application;

                        foreach (ObjectId id1 in BTrecord)
                        {
                            Entity Ent1 = Trans1.GetObject(id1, OpenMode.ForRead) as Entity;
                            if (Ent1 is Curve && Ent1.Layer == pipes_layer)
                            {
                                dt1.Rows.Add();
                                dt1.Rows[dt1.Rows.Count - 1][0] = id1;
                                dt1.Rows[dt1.Rows.Count - 1][1] = Ent1.Layer;
                            }

                            if (Ent1 is Curve && Ent1.Layer == buoyancy_layer)
                            {
                                dt2.Rows.Add();
                                dt2.Rows[dt2.Rows.Count - 1][0] = id1;
                                dt2.Rows[dt2.Rows.Count - 1][1] = Ent1.Layer;
                            }

                            if ((Ent1 is DBPoint || Ent1 is BlockReference) &&
                                (Ent1.Layer == cpac_layer || Ent1.Layer == es_layer ||
                                Ent1.Layer == hydrotest_layerPT || Ent1.Layer == xing_layer || Ent1.Layer == transition_layer))
                            {
                                dt2.Rows.Add();
                                dt2.Rows[dt2.Rows.Count - 1][0] = id1;
                                dt2.Rows[dt2.Rows.Count - 1][1] = Ent1.Layer;
                            }

                            if ((Ent1 is DBPoint || Ent1 is BlockReference) && Ent1.Layer == cpac_layer)
                            {
                                dt3.Rows.Add();
                                dt3.Rows[dt3.Rows.Count - 1][0] = id1;
                                dt3.Rows[dt3.Rows.Count - 1][1] = Ent1.Layer;
                            }

                            if ((Ent1 is DBPoint || Ent1 is BlockReference) && Ent1.Layer == es_layer)
                            {
                                dt4.Rows.Add();
                                dt4.Rows[dt4.Rows.Count - 1][0] = id1;
                                dt4.Rows[dt4.Rows.Count - 1][1] = Ent1.Layer;
                            }

                            if ((Ent1 is DBPoint || Ent1 is BlockReference) && Ent1.Layer == hydrotest_layerPT)
                            {
                                dt5.Rows.Add();
                                dt5.Rows[dt5.Rows.Count - 1][0] = id1;
                                dt5.Rows[dt5.Rows.Count - 1][1] = Ent1.Layer;
                            }

                            if ((Ent1 is DBPoint || Ent1 is BlockReference) && Ent1.Layer == xing_layer)
                            {
                                dt6.Rows.Add();
                                dt6.Rows[dt6.Rows.Count - 1][0] = id1;
                                dt6.Rows[dt6.Rows.Count - 1][1] = Ent1.Layer;
                            }

                            if ((Ent1 is DBPoint || Ent1 is BlockReference) && Ent1.Layer == transition_layer)
                            {
                                dt7.Rows.Add();
                                dt7.Rows[dt7.Rows.Count - 1][0] = id1;
                                dt7.Rows[dt7.Rows.Count - 1][1] = Ent1.Layer;
                            }

                            if (Ent1 is Curve && Ent1.Layer == elbows_layer)
                            {
                                dt8.Rows.Add();
                                dt8.Rows[dt8.Rows.Count - 1][0] = id1;
                                dt8.Rows[dt8.Rows.Count - 1][1] = Ent1.Layer;
                            }

                            if (Ent1 is Curve && Ent1.Layer == fab_layer)
                            {
                                dt9.Rows.Add();
                                dt9.Rows[dt9.Rows.Count - 1][0] = id1;
                                dt9.Rows[dt9.Rows.Count - 1][1] = Ent1.Layer;
                            }

                            if (Ent1 is Curve && Ent1.Layer == class_layer)
                            {
                                dt10.Rows.Add();
                                dt10.Rows[dt10.Rows.Count - 1][0] = id1;
                                dt10.Rows[dt10.Rows.Count - 1][1] = Ent1.Layer;
                            }

                            if (Ent1 is Curve && Ent1.Layer == geotech_layer)
                            {
                                dt11.Rows.Add();
                                dt11.Rows[dt11.Rows.Count - 1][0] = id1;
                                dt11.Rows[dt11.Rows.Count - 1][1] = Ent1.Layer;
                            }

                            if (Ent1 is Curve && Ent1.Layer == doc_layer)
                            {
                                dt12.Rows.Add();
                                dt12.Rows[dt12.Rows.Count - 1][0] = id1;
                                dt12.Rows[dt12.Rows.Count - 1][1] = Ent1.Layer;
                            }

                            if ((Ent1 is DBPoint || Ent1 is BlockReference) && Ent1.Layer == buoyancy_pt_layer)
                            {
                                dt13.Rows.Add();
                                dt13.Rows[dt13.Rows.Count - 1][0] = id1;
                                dt13.Rows[dt13.Rows.Count - 1][1] = Ent1.Layer;
                            }

                            if ((Ent1 is DBPoint || Ent1 is BlockReference) && Ent1.Layer == geotech_layer_pt)
                            {
                                dt14.Rows.Add();
                                dt14.Rows[dt14.Rows.Count - 1][0] = id1;
                                dt14.Rows[dt14.Rows.Count - 1][1] = Ent1.Layer;
                            }

                            if ((Ent1 is DBPoint || Ent1 is BlockReference) && Ent1.Layer == muskeg_layer)
                            {
                                dt15.Rows.Add();
                                dt15.Rows[dt15.Rows.Count - 1][0] = id1;
                                dt15.Rows[dt15.Rows.Count - 1][1] = Ent1.Layer;
                            }

                            if (Ent1 is Curve && Ent1.Layer == hydrotest_layer)
                            {
                                dt16.Rows.Add();
                                dt16.Rows[dt16.Rows.Count - 1][0] = id1;
                                dt16.Rows[dt16.Rows.Count - 1][1] = Ent1.Layer;
                            }


                            if (Ent1 is Curve && Ent1.Layer == preexisting_layer)
                            {
                                dt17.Rows.Add();
                                dt17.Rows[dt17.Rows.Count - 1][0] = id1;
                                dt17.Rows[dt17.Rows.Count - 1][1] = Ent1.Layer;
                            }
                        }

                        if (dt1.Rows.Count > 0)
                        {
                            ObjectIdCollection col_objectID = new ObjectIdCollection();
                            for (int i = 0; i < dt1.Rows.Count; ++i)
                            {

                                if (dt1.Rows[i][0] != DBNull.Value)
                                {
                                    col_objectID.Add((ObjectId)dt1.Rows[i][0]);
                                }
                            }
                            string filename = textBox_output_folder.Text;

                            if (filename.Substring(filename.Length - 1, 1) != "\\")
                            {
                                filename = filename + "\\";
                            }

                            if (System.IO.Directory.Exists(filename) == true)
                            {
                                int incr = 0;
                                string suff1 = "";
                                bool exista = true;
                                do
                                {

                                    if (System.IO.File.Exists(filename + pipes_layer + suff1 + ".shp") == false)
                                    {
                                        filename = filename + pipes_layer + suff1 + ".shp";
                                        exista = false;
                                    }
                                    else
                                    {

                                        ++incr;
                                        suff1 = incr.ToString();
                                    }

                                } while (exista == true);
                            }
                            ExportSHP("SHP", filename, pipes_od, true, false, "line", col_objectID);
                        }


                        if (dt2.Rows.Count > 0)
                        {
                            ObjectIdCollection col_objectID = new ObjectIdCollection();
                            for (int i = 0; i < dt2.Rows.Count; ++i)
                            {

                                if (dt2.Rows[i][0] != DBNull.Value)
                                {
                                    col_objectID.Add((ObjectId)dt2.Rows[i][0]);
                                }
                            }
                            string filename = textBox_output_folder.Text;

                            if (filename.Substring(filename.Length - 1, 1) != "\\")
                            {
                                filename = filename + "\\";
                            }

                            if (System.IO.Directory.Exists(filename) == true)
                            {
                                int incr = 0;
                                string suff1 = "";
                                bool exista = true;
                                do
                                {

                                    if (System.IO.File.Exists(filename + buoyancy_layer + suff1 + ".shp") == false)
                                    {
                                        filename = filename + buoyancy_layer + suff1 + ".shp";
                                        exista = false;
                                    }
                                    else
                                    {

                                        ++incr;
                                        suff1 = incr.ToString();
                                    }

                                } while (exista == true);
                            }
                            ExportSHP("SHP", filename, buoyancy_od, true, false, "line", col_objectID);
                        }


                        if (dt3.Rows.Count > 0)
                        {
                            ObjectIdCollection col_objectID = new ObjectIdCollection();
                            for (int i = 0; i < dt3.Rows.Count; ++i)
                            {

                                if (dt3.Rows[i][0] != DBNull.Value)
                                {
                                    col_objectID.Add((ObjectId)dt3.Rows[i][0]);
                                }
                            }
                            string filename = textBox_output_folder.Text;

                            if (filename.Substring(filename.Length - 1, 1) != "\\")
                            {
                                filename = filename + "\\";
                            }

                            if (System.IO.Directory.Exists(filename) == true)
                            {
                                int incr = 0;
                                string suff1 = "";
                                bool exista = true;
                                do
                                {

                                    if (System.IO.File.Exists(filename + cpac_layer + suff1 + ".shp") == false)
                                    {
                                        filename = filename + cpac_layer + suff1 + ".shp";
                                        exista = false;
                                    }
                                    else
                                    {

                                        ++incr;
                                        suff1 = incr.ToString();
                                    }

                                } while (exista == true);
                            }
                            ExportSHP("SHP", filename, cpac_od, true, false, "point", col_objectID);
                        }


                        if (dt4.Rows.Count > 0)
                        {
                            ObjectIdCollection col_objectID = new ObjectIdCollection();
                            for (int i = 0; i < dt4.Rows.Count; ++i)
                            {

                                if (dt4.Rows[i][0] != DBNull.Value)
                                {
                                    col_objectID.Add((ObjectId)dt4.Rows[i][0]);
                                }
                            }
                            string filename = textBox_output_folder.Text;

                            if (filename.Substring(filename.Length - 1, 1) != "\\")
                            {
                                filename = filename + "\\";
                            }

                            if (System.IO.Directory.Exists(filename) == true)
                            {
                                int incr = 0;
                                string suff1 = "";
                                bool exista = true;
                                do
                                {

                                    if (System.IO.File.Exists(filename + es_layer + suff1 + ".shp") == false)
                                    {
                                        filename = filename + es_layer + suff1 + ".shp";
                                        exista = false;
                                    }
                                    else
                                    {

                                        ++incr;
                                        suff1 = incr.ToString();
                                    }

                                } while (exista == true);
                            }
                            ExportSHP("SHP", filename, es_od, true, false, "point", col_objectID);
                        }


                        if (dt5.Rows.Count > 0)
                        {
                            ObjectIdCollection col_objectID = new ObjectIdCollection();
                            for (int i = 0; i < dt5.Rows.Count; ++i)
                            {

                                if (dt5.Rows[i][0] != DBNull.Value)
                                {
                                    col_objectID.Add((ObjectId)dt5.Rows[i][0]);
                                }
                            }
                            string filename = textBox_output_folder.Text;

                            if (filename.Substring(filename.Length - 1, 1) != "\\")
                            {
                                filename = filename + "\\";
                            }

                            if (System.IO.Directory.Exists(filename) == true)
                            {
                                int incr = 0;
                                string suff1 = "";
                                bool exista = true;
                                do
                                {

                                    if (System.IO.File.Exists(filename + hydrotest_layerPT + suff1 + ".shp") == false)
                                    {
                                        filename = filename + hydrotest_layerPT + suff1 + ".shp";
                                        exista = false;
                                    }
                                    else
                                    {

                                        ++incr;
                                        suff1 = incr.ToString();
                                    }

                                } while (exista == true);
                            }
                            ExportSHP("SHP", filename, hydrotest_odPT, true, false, "point", col_objectID);
                        }

                        if (dt6.Rows.Count > 0)
                        {
                            ObjectIdCollection col_objectID = new ObjectIdCollection();
                            for (int i = 0; i < dt6.Rows.Count; ++i)
                            {

                                if (dt6.Rows[i][0] != DBNull.Value)
                                {
                                    col_objectID.Add((ObjectId)dt6.Rows[i][0]);
                                }
                            }
                            string filename = textBox_output_folder.Text;

                            if (filename.Substring(filename.Length - 1, 1) != "\\")
                            {
                                filename = filename + "\\";
                            }

                            if (System.IO.Directory.Exists(filename) == true)
                            {
                                int incr = 0;
                                string suff1 = "";
                                bool exista = true;
                                do
                                {

                                    if (System.IO.File.Exists(filename + xing_layer + suff1 + ".shp") == false)
                                    {
                                        filename = filename + xing_layer + suff1 + ".shp";
                                        exista = false;
                                    }
                                    else
                                    {

                                        ++incr;
                                        suff1 = incr.ToString();
                                    }

                                } while (exista == true);
                            }
                            ExportSHP("SHP", filename, xing_od, true, false, "point", col_objectID);
                        }

                        if (dt7.Rows.Count > 0)
                        {
                            ObjectIdCollection col_objectID = new ObjectIdCollection();
                            for (int i = 0; i < dt7.Rows.Count; ++i)
                            {

                                if (dt7.Rows[i][0] != DBNull.Value)
                                {
                                    col_objectID.Add((ObjectId)dt7.Rows[i][0]);
                                }
                            }
                            string filename = textBox_output_folder.Text;

                            if (filename.Substring(filename.Length - 1, 1) != "\\")
                            {
                                filename = filename + "\\";
                            }

                            if (System.IO.Directory.Exists(filename) == true)
                            {
                                int incr = 0;
                                string suff1 = "";
                                bool exista = true;
                                do
                                {

                                    if (System.IO.File.Exists(filename + transition_layer + suff1 + ".shp") == false)
                                    {
                                        filename = filename + transition_layer + suff1 + ".shp";
                                        exista = false;
                                    }
                                    else
                                    {

                                        ++incr;
                                        suff1 = incr.ToString();
                                    }

                                } while (exista == true);
                            }
                            ExportSHP("SHP", filename, transition_od, true, false, "point", col_objectID);
                        }


                        if (dt8.Rows.Count > 0)
                        {
                            ObjectIdCollection col_objectID = new ObjectIdCollection();
                            for (int i = 0; i < dt8.Rows.Count; ++i)
                            {

                                if (dt8.Rows[i][0] != DBNull.Value)
                                {
                                    col_objectID.Add((ObjectId)dt8.Rows[i][0]);
                                }
                            }
                            string filename = textBox_output_folder.Text;

                            if (filename.Substring(filename.Length - 1, 1) != "\\")
                            {
                                filename = filename + "\\";
                            }

                            if (System.IO.Directory.Exists(filename) == true)
                            {
                                int incr = 0;
                                string suff1 = "";
                                bool exista = true;
                                do
                                {

                                    if (System.IO.File.Exists(filename + elbows_layer + suff1 + ".shp") == false)
                                    {
                                        filename = filename + elbows_layer + suff1 + ".shp";
                                        exista = false;
                                    }
                                    else
                                    {

                                        ++incr;
                                        suff1 = incr.ToString();
                                    }

                                } while (exista == true);
                            }
                            ExportSHP("SHP", filename, elbows_od, true, false, "line", col_objectID);
                        }

                        if (dt9.Rows.Count > 0)
                        {
                            ObjectIdCollection col_objectID = new ObjectIdCollection();
                            for (int i = 0; i < dt9.Rows.Count; ++i)
                            {

                                if (dt9.Rows[i][0] != DBNull.Value)
                                {
                                    col_objectID.Add((ObjectId)dt9.Rows[i][0]);
                                }
                            }
                            string filename = textBox_output_folder.Text;

                            if (filename.Substring(filename.Length - 1, 1) != "\\")
                            {
                                filename = filename + "\\";
                            }

                            if (System.IO.Directory.Exists(filename) == true)
                            {
                                int incr = 0;
                                string suff1 = "";
                                bool exista = true;
                                do
                                {

                                    if (System.IO.File.Exists(filename + fab_layer + suff1 + ".shp") == false)
                                    {
                                        filename = filename + fab_layer + suff1 + ".shp";
                                        exista = false;
                                    }
                                    else
                                    {

                                        ++incr;
                                        suff1 = incr.ToString();
                                    }

                                } while (exista == true);
                            }
                            ExportSHP("SHP", filename, fab_od, true, false, "line", col_objectID);
                        }


                        if (dt10.Rows.Count > 0)
                        {
                            ObjectIdCollection col_objectID = new ObjectIdCollection();
                            for (int i = 0; i < dt10.Rows.Count; ++i)
                            {

                                if (dt10.Rows[i][0] != DBNull.Value)
                                {
                                    col_objectID.Add((ObjectId)dt10.Rows[i][0]);
                                }
                            }
                            string filename = textBox_output_folder.Text;

                            if (filename.Substring(filename.Length - 1, 1) != "\\")
                            {
                                filename = filename + "\\";
                            }

                            if (System.IO.Directory.Exists(filename) == true)
                            {
                                int incr = 0;
                                string suff1 = "";
                                bool exista = true;
                                do
                                {

                                    if (System.IO.File.Exists(filename + class_layer + suff1 + ".shp") == false)
                                    {
                                        filename = filename + class_layer + suff1 + ".shp";
                                        exista = false;
                                    }
                                    else
                                    {

                                        ++incr;
                                        suff1 = incr.ToString();
                                    }

                                } while (exista == true);
                            }
                            ExportSHP("SHP", filename, class_od, true, false, "line", col_objectID);
                        }


                        if (dt11.Rows.Count > 0)
                        {
                            ObjectIdCollection col_objectID = new ObjectIdCollection();
                            for (int i = 0; i < dt11.Rows.Count; ++i)
                            {

                                if (dt11.Rows[i][0] != DBNull.Value)
                                {
                                    col_objectID.Add((ObjectId)dt11.Rows[i][0]);
                                }
                            }
                            string filename = textBox_output_folder.Text;

                            if (filename.Substring(filename.Length - 1, 1) != "\\")
                            {
                                filename = filename + "\\";
                            }

                            if (System.IO.Directory.Exists(filename) == true)
                            {
                                int incr = 0;
                                string suff1 = "";
                                bool exista = true;
                                do
                                {

                                    if (System.IO.File.Exists(filename + geotech_layer + suff1 + ".shp") == false)
                                    {
                                        filename = filename + geotech_layer + suff1 + ".shp";
                                        exista = false;
                                    }
                                    else
                                    {

                                        ++incr;
                                        suff1 = incr.ToString();
                                    }

                                } while (exista == true);
                            }
                            ExportSHP("SHP", filename, geotech_od, true, false, "line", col_objectID);
                        }

                        if (dt12.Rows.Count > 0)
                        {
                            ObjectIdCollection col_objectID = new ObjectIdCollection();
                            for (int i = 0; i < dt12.Rows.Count; ++i)
                            {

                                if (dt12.Rows[i][0] != DBNull.Value)
                                {
                                    col_objectID.Add((ObjectId)dt12.Rows[i][0]);
                                }
                            }
                            string filename = textBox_output_folder.Text;

                            if (filename.Substring(filename.Length - 1, 1) != "\\")
                            {
                                filename = filename + "\\";
                            }

                            if (System.IO.Directory.Exists(filename) == true)
                            {
                                int incr = 0;
                                string suff1 = "";
                                bool exista = true;
                                do
                                {

                                    if (System.IO.File.Exists(filename + doc_layer + suff1 + ".shp") == false)
                                    {
                                        filename = filename + doc_layer + suff1 + ".shp";
                                        exista = false;
                                    }
                                    else
                                    {

                                        ++incr;
                                        suff1 = incr.ToString();
                                    }

                                } while (exista == true);
                            }
                            ExportSHP("SHP", filename, doc_od, true, false, "line", col_objectID);
                        }

                        if (dt13.Rows.Count > 0)
                        {
                            ObjectIdCollection col_objectID = new ObjectIdCollection();
                            for (int i = 0; i < dt13.Rows.Count; ++i)
                            {

                                if (dt13.Rows[i][0] != DBNull.Value)
                                {
                                    col_objectID.Add((ObjectId)dt13.Rows[i][0]);
                                }
                            }
                            string filename = textBox_output_folder.Text;

                            if (filename.Substring(filename.Length - 1, 1) != "\\")
                            {
                                filename = filename + "\\";
                            }

                            if (System.IO.Directory.Exists(filename) == true)
                            {
                                int incr = 0;
                                string suff1 = "";
                                bool exista = true;
                                do
                                {

                                    if (System.IO.File.Exists(filename + buoyancy_pt_layer + suff1 + ".shp") == false)
                                    {
                                        filename = filename + buoyancy_pt_layer + suff1 + ".shp";
                                        exista = false;
                                    }
                                    else
                                    {

                                        ++incr;
                                        suff1 = incr.ToString();
                                    }

                                } while (exista == true);
                            }
                            ExportSHP("SHP", filename, buoyancy_pt_od, true, false, "point", col_objectID);
                        }


                        if (dt14.Rows.Count > 0)
                        {
                            ObjectIdCollection col_objectID = new ObjectIdCollection();
                            for (int i = 0; i < dt14.Rows.Count; ++i)
                            {

                                if (dt14.Rows[i][0] != DBNull.Value)
                                {
                                    col_objectID.Add((ObjectId)dt14.Rows[i][0]);
                                }
                            }
                            string filename = textBox_output_folder.Text;

                            if (filename.Substring(filename.Length - 1, 1) != "\\")
                            {
                                filename = filename + "\\";
                            }

                            if (System.IO.Directory.Exists(filename) == true)
                            {
                                int incr = 0;
                                string suff1 = "";
                                bool exista = true;
                                do
                                {

                                    if (System.IO.File.Exists(filename + geotech_layer_pt + suff1 + ".shp") == false)
                                    {
                                        filename = filename + geotech_layer_pt + suff1 + ".shp";
                                        exista = false;
                                    }
                                    else
                                    {

                                        ++incr;
                                        suff1 = incr.ToString();
                                    }

                                } while (exista == true);
                            }
                            ExportSHP("SHP", filename, geotech_od_pt, true, false, "point", col_objectID);
                        }

                        if (dt15.Rows.Count > 0)
                        {
                            ObjectIdCollection col_objectID = new ObjectIdCollection();
                            for (int i = 0; i < dt15.Rows.Count; ++i)
                            {

                                if (dt15.Rows[i][0] != DBNull.Value)
                                {
                                    col_objectID.Add((ObjectId)dt15.Rows[i][0]);
                                }
                            }
                            string filename = textBox_output_folder.Text;

                            if (filename.Substring(filename.Length - 1, 1) != "\\")
                            {
                                filename = filename + "\\";
                            }

                            if (System.IO.Directory.Exists(filename) == true)
                            {
                                int incr = 0;
                                string suff1 = "";
                                bool exista = true;
                                do
                                {

                                    if (System.IO.File.Exists(filename + muskeg_layer + suff1 + ".shp") == false)
                                    {
                                        filename = filename + muskeg_layer + suff1 + ".shp";
                                        exista = false;
                                    }
                                    else
                                    {

                                        ++incr;
                                        suff1 = incr.ToString();
                                    }

                                } while (exista == true);
                            }
                            ExportSHP("SHP", filename, muskeg_od, true, false, "point", col_objectID);
                        }


                        if (dt16.Rows.Count > 0)
                        {
                            ObjectIdCollection col_objectID = new ObjectIdCollection();
                            for (int i = 0; i < dt16.Rows.Count; ++i)
                            {

                                if (dt16.Rows[i][0] != DBNull.Value)
                                {
                                    col_objectID.Add((ObjectId)dt16.Rows[i][0]);
                                }
                            }
                            string filename = textBox_output_folder.Text;

                            if (filename.Substring(filename.Length - 1, 1) != "\\")
                            {
                                filename = filename + "\\";
                            }

                            if (System.IO.Directory.Exists(filename) == true)
                            {
                                int incr = 0;
                                string suff1 = "";
                                bool exista = true;
                                do
                                {

                                    if (System.IO.File.Exists(filename + hydrotest_layer + suff1 + ".shp") == false)
                                    {
                                        filename = filename + hydrotest_layer + suff1 + ".shp";
                                        exista = false;
                                    }
                                    else
                                    {

                                        ++incr;
                                        suff1 = incr.ToString();
                                    }

                                } while (exista == true);
                            }
                            ExportSHP("SHP", filename, hydrotest_od, true, false, "line", col_objectID);
                        }

                        if (dt17.Rows.Count > 0)
                        {
                            ObjectIdCollection col_objectID = new ObjectIdCollection();
                            for (int i = 0; i < dt17.Rows.Count; ++i)
                            {

                                if (dt17.Rows[i][0] != DBNull.Value)
                                {
                                    col_objectID.Add((ObjectId)dt17.Rows[i][0]);
                                }
                            }
                            string filename = textBox_output_folder.Text;

                            if (filename.Substring(filename.Length - 1, 1) != "\\")
                            {
                                filename = filename + "\\";
                            }

                            if (System.IO.Directory.Exists(filename) == true)
                            {
                                int incr = 0;
                                string suff1 = "";
                                bool exista = true;
                                do
                                {

                                    if (System.IO.File.Exists(filename + preexisting_layer + suff1 + ".shp") == false)
                                    {
                                        filename = filename + preexisting_layer + suff1 + ".shp";
                                        exista = false;
                                    }
                                    else
                                    {

                                        ++incr;
                                        suff1 = incr.ToString();
                                    }

                                } while (exista == true);
                            }
                            ExportSHP("SHP", filename, preexisting_od, true, false, "line", col_objectID);
                        }


                        if (dt99.Rows.Count > 0)
                        {
                            ObjectIdCollection col_objectID = new ObjectIdCollection();
                            for (int i = 0; i < dt99.Rows.Count; ++i)
                            {

                                if (dt99.Rows[i][0] != DBNull.Value)
                                {
                                    col_objectID.Add((ObjectId)dt99.Rows[i][0]);
                                }
                            }
                            string filename = textBox_output_folder.Text;

                            if (filename.Substring(filename.Length - 1, 1) != "\\")
                            {
                                filename = filename + "\\";
                            }

                            if (System.IO.Directory.Exists(filename) == true)
                            {
                                int incr = 0;
                                string suff1 = "";
                                bool exista = true;
                                do
                                {

                                    if (System.IO.File.Exists(filename + dt99_layer + suff1 + ".shp") == false)
                                    {
                                        filename = filename + dt99_layer + suff1 + ".shp";
                                        exista = false;
                                    }
                                    else
                                    {

                                        ++incr;
                                        suff1 = incr.ToString();
                                    }

                                } while (exista == true);
                            }
                            ExportSHP("SHP", filename, dt99_od, true, false, "polygon", col_objectID);
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
            set_enable_true();

        }



        public void ExportSHP(string format, string filename, string table_name, bool isODTable, bool isLinkTemplate, string geomtryType, ObjectIdCollection object_id_col)
        {
            Exporter exporter = null;
            try
            {
                MapApplication mapApp = HostMapApplicationServices.Application;
                exporter = mapApp.Exporter;
                // Initiate the exporter
                exporter.Init(format, filename);
                if (object_id_col != null && object_id_col.Count > 0)
                {
                    exporter.ExportAll = false;

                    exporter.SetSelectionSet(object_id_col);
                }
                else
                {
                    //exporter.ExportAll = true;
                }

                GeometryType geomtryType1 = (GeometryType)Enum.Parse(typeof(GeometryType), geomtryType, true);
                exporter.SetStorageOptions(StorageType.FileOneEntityType, geomtryType1, string.Empty);

                // Get Data mapping object
                ExpressionTargetCollection dataMapping = null;
                dataMapping = exporter.GetExportDataMappings();
                dataMapping.Clear();

                // Set ObjectData data mapping if isODTable is true
                if (isODTable == true && MapODData(dataMapping, table_name) == true)
                {
                    // Reset Data mapping with Object data and Link template keys		
                    exporter.SetExportDataMappings(dataMapping);
                }
                // If layerFilter isn't null, set the layer filter to export layer by layer
                if (null != table_name)
                {
                    exporter.LayerFilter = table_name;
                }

                // Do the exporting and log the result
                ExportResults results;
                results = exporter.Export(true);

                Utility.ShowMsg("\nExporting succeeded.");
            }
            catch (MapException e)
            {

            }
            finally
            {

            }
        }


        public bool MapODData(ExpressionTargetCollection mapping, string tablename)
        {
            MapApplication mapApi = HostMapApplicationServices.Application;
            ProjectModel proj = mapApi.ActiveProject;
            Tables tables = proj.ODTables;
            if (tables.IsTableDefined(tablename) == true)
            {
                try
                {
                    Autodesk.Gis.Map.ObjectData.Table table = tables[tablename];
                    FieldDefinitions definitions = table.FieldDefinitions;
                    for (int j = 0; j < definitions.Count; j++)
                    {
                        FieldDefinition column = null;
                        column = definitions[j];
                        // fieldName is the OD table field name in the data mapping. It should be 
                        // in the format:fieldName&tableName. 
                        // newFieldName is the attribute field name of exported-to file
                        string ODfieldName = ":" + column.Name + "@" + tablename;
                        string shpFieldName = column.Name;

                        mapping.Add(ODfieldName, shpFieldName);
                    }
                }
                catch (MapImportExportException)
                {
                    return false;
                }
                return true;
            }
            else
            {
                return false;
            }

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
                    if (Name_of_layer.Contains("|") == false && Name_of_layer.Contains("$") == false && Layer1.IsFrozen == false && Layer1.IsOff == false)
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


    }
    public sealed class Utility
    {
        private Utility()
        {
        }

        public static void ShowMsg(string msg)
        {
            AcadEditor.WriteMessage(msg);
        }

        public static Autodesk.AutoCAD.EditorInput.Editor AcadEditor
        {
            get
            {
                return Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            }
        }
    }
}

