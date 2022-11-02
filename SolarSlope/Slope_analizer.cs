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
using Font = System.Drawing.Font;
using Autodesk.Civil.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;

namespace Alignment_mdi
{
    public partial class SolarSlope_form : Form
    {

        private ContextMenuStrip ContextMenuStrip_go_to_error;

        string col_y = "Northing\r\n(Y)";
        string col_x = "Easting\r\n(X)";
        string col_z = "Elevation\r\n(Z)";
        string col_pn = "Point Number";
        string col_descr = "Description\r\n(D)";
        string col_slope_NORTH = "North\r\nSlope %";
        string col_slopeN1 = "North\r\nSlope\r\nNew %";
        string col_NORTH_PN = "North\r\nPoint number";
        string col_slope_SOUTH = "South\r\nSlope %";
        string col_slopeS1 = "South\r\nSlope\r\nNew %";

        string col_SOUTH_PN = "South\r\nPoint number";
        string col_slope_EAST = "East\r\nSlope %";
        string col_slopeE1 = "East\r\nSlope\r\nNew %";

        string col_EAST_PN = "East\r\nPoint number";
        string col_slope_WEST = "West\r\nSlope %";
        string col_slopeW1 = "West\r\nSlope\r\nNew %";

        string col_WEST_PN = "West\r\nPoint number";
        string col_newZ_NS = "NS\r\nNew Elevation";
        string col_newZ_EW = "EW\r\nNew Elevation";
        string col_max_EW = "EW\r\nMax Slope";
        string col_max_NS = "NS\r\nMax Slope";
        string col_new_elev = "NEW\r\nElevation";
        string col_new_description = "NEW\r\nDescription";

        string col_NS_processed = "NS\r\nProcessed";
        string col_EW_processed = "EW\r\nProcessed";
        string col_id = "objectid";
        string col_anchor = "Anchor";

        double max_NS1 = 300;
        double max_EW1 = 300;
        double min_NS2 = -300;
        double min_EW2 = -300;

        double gridH = 0;
        double gridV = 0;


        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();

            lista_butoane.Add(button_calc);
            lista_butoane.Add(button_load_layers_to_combobox_point_layer1);
            lista_butoane.Add(button_load_layers_to_combobox_point_layer2);
            lista_butoane.Add(button_load_point_label_style1);
            lista_butoane.Add(button_load_point_label_style2);
            lista_butoane.Add(button_load_point_style1);
            lista_butoane.Add(button_load_point_style2);
            lista_butoane.Add(button_udp1);
            lista_butoane.Add(button_udp2);



            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {

                bt1.Enabled = false;

            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();

            lista_butoane.Add(button_calc);
            lista_butoane.Add(button_load_layers_to_combobox_point_layer1);
            lista_butoane.Add(button_load_layers_to_combobox_point_layer2);
            lista_butoane.Add(button_load_point_label_style1);
            lista_butoane.Add(button_load_point_label_style2);
            lista_butoane.Add(button_load_point_style1);
            lista_butoane.Add(button_load_point_style2);
            lista_butoane.Add(button_udp1);
            lista_butoane.Add(button_udp2);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }

        public SolarSlope_form()
        {
            InitializeComponent();
            var toolStripMenuItem2 = new ToolStripMenuItem { Text = "tool strip menu" };
            //toolStripMenuItem2.Click += go_to_excel_point;


            ContextMenuStrip_go_to_error = new ContextMenuStrip();
            ContextMenuStrip_go_to_error.Items.AddRange(new ToolStripItem[] { toolStripMenuItem2 });


        }


        private void button_calc_Click(object sender, EventArgs e)
        {
            if (Functions.IsNumeric(textBox_H.Text) == true)
            {
                gridH = Math.Abs(Convert.ToDouble(textBox_H.Text));
            }

            if (Functions.IsNumeric(textBox_V.Text) == true)
            {
                gridV = Math.Abs(Convert.ToDouble(textBox_V.Text));
            }

            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.Civil.ApplicationServices.CivilDocument CivilDrawing = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument;

            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            Editor1.SetImpliedSelection(Empty_array);
            try
            {

                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                        set_enable_false();
                        this.MdiParent.WindowState = FormWindowState.Minimized;

                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez.MessageForAdding = "\nSelect points";
                        Prompt_rez.SingleOnly = false;





                        Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            set_enable_true();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            this.MdiParent.WindowState = FormWindowState.Normal;

                            return;
                        }

                        System.Data.DataTable dt1 = new System.Data.DataTable();
                        dt1.Columns.Add(col_pn, typeof(string));
                        dt1.Columns.Add(col_y, typeof(double));
                        dt1.Columns.Add(col_x, typeof(double));
                        dt1.Columns.Add(col_z, typeof(double));
                        dt1.Columns.Add(col_descr, typeof(string));
                        dt1.Columns.Add(col_slope_NORTH, typeof(double));
                        dt1.Columns.Add(col_NORTH_PN, typeof(string));
                        dt1.Columns.Add(col_slope_SOUTH, typeof(double));
                        dt1.Columns.Add(col_SOUTH_PN, typeof(string));

                        dt1.Columns.Add(col_slope_EAST, typeof(double));
                        dt1.Columns.Add(col_EAST_PN, typeof(string));
                        dt1.Columns.Add(col_slope_WEST, typeof(double));
                        dt1.Columns.Add(col_WEST_PN, typeof(string));

                        dt1.Columns.Add(col_newZ_NS, typeof(double));
                        dt1.Columns.Add(col_max_NS, typeof(double));

                        dt1.Columns.Add(col_newZ_EW, typeof(double));
                        dt1.Columns.Add(col_max_EW, typeof(double));
                        dt1.Columns.Add(col_new_elev, typeof(double));

                        dt1.Columns.Add(col_slopeN1, typeof(double));
                        dt1.Columns.Add(col_slopeS1, typeof(double));
                        dt1.Columns.Add(col_slopeE1, typeof(double));
                        dt1.Columns.Add(col_slopeW1, typeof(double));

                        dt1.Columns.Add(col_new_description, typeof(string));
                        dt1.Columns.Add(col_NS_processed, typeof(bool));

                        dt1.Columns.Add(col_EW_processed, typeof(bool));

                        dt1.Columns.Add(col_id, typeof(ObjectId));


                        System.Data.DataTable dt3 = new System.Data.DataTable();
                        dt3.Columns.Add(col_pn, typeof(string));
                        dt3.Columns.Add(col_y, typeof(double));
                        dt3.Columns.Add(col_x, typeof(double));
                        dt3.Columns.Add(col_z, typeof(double));
                        dt3.Columns.Add(col_descr, typeof(string));
                        dt3.Columns.Add(col_new_description, typeof(string));
                        dt3.Columns.Add(col_new_elev, typeof(double));



                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {
                            CogoPoint cg1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as CogoPoint;
                            if (cg1 != null)
                            {
                                dt1.Rows.Add();
                                dt1.Rows[dt1.Rows.Count - 1][col_pn] = cg1.PointNumber;
                                dt1.Rows[dt1.Rows.Count - 1][col_id] = cg1.ObjectId;

                                dt1.Rows[dt1.Rows.Count - 1][col_x] = Math.Round(cg1.Location.X, 4);
                                dt1.Rows[dt1.Rows.Count - 1][col_y] = Math.Round(cg1.Location.Y, 4);

                                dt1.Rows[dt1.Rows.Count - 1][col_NS_processed] = false;
                                dt1.Rows[dt1.Rows.Count - 1][col_EW_processed] = false;



                                string udp1_string = "";
                                double udp2_double = -1.234;
                                UDPString udp1 = Functions.Find_udp_string(comboBox_udp_field_1.Text);
                                if (udp1 != null)
                                {
                                    udp1_string = Convert.ToString(cg1.GetUDPValue(udp1));
                                }
                                UDPDouble udp2 = Functions.Find_udp_double(comboBox_udp_field_2.Text);
                                if (udp2 != null)
                                {
                                    udp2_double = Convert.ToDouble(cg1.GetUDPValue(udp2));
                                }

                                if (udp1_string != "")
                                {
                                    dt1.Rows[dt1.Rows.Count - 1][col_descr] = udp1_string;
                                    dt1.Rows[dt1.Rows.Count - 1][col_new_description] = udp1_string;
                                }
                                else
                                {
                                    dt1.Rows[dt1.Rows.Count - 1][col_descr] = cg1.RawDescription;
                                    dt1.Rows[dt1.Rows.Count - 1][col_new_description] = cg1.RawDescription;
                                }

                                if (udp2_double != -1.234)
                                {
                                    dt1.Rows[dt1.Rows.Count - 1][col_z] = Math.Round(udp2_double, 4);
                                }
                                else
                                {
                                    dt1.Rows[dt1.Rows.Count - 1][col_z] = Math.Round(cg1.Elevation, 4);
                                }

                            }
                        }

                        System.Data.DataTable dt2 = Functions.Sort_data_table(dt1, col_y);

                        for (int i = 0; i < dt2.Rows.Count; ++i)
                        {
                            double z1 = Convert.ToDouble(dt2.Rows[i][col_z]);
                            double x1 = Convert.ToDouble(dt2.Rows[i][col_x]);
                            double y1 = Convert.ToDouble(dt2.Rows[i][col_y]);

                            for (int j = i + 1; j < dt2.Rows.Count; ++j)
                            {
                                double z2 = Convert.ToDouble(dt2.Rows[j][col_z]);
                                double x2 = Convert.ToDouble(dt2.Rows[j][col_x]);
                                double y2 = Convert.ToDouble(dt2.Rows[j][col_y]);
                                string pn2 = Convert.ToString(dt2.Rows[j][col_pn]);
                                double Run = Math.Abs(y2 - y1);

                                if (z2 != z1 && x1 == x2 && Math.Round(Run, 4) == Math.Round(gridV, 4))
                                {
                                    double Rise = z1 - z2;
                                    double Slope = Math.Round(100 * Rise / Run, 4);
                                    dt2.Rows[i][col_slope_NORTH] = Slope;
                                    dt2.Rows[i][col_NORTH_PN] = pn2;

                                    j = dt2.Rows.Count;
                                }
                            }
                        }

                        dt2 = Functions.Sort_data_table_desc(dt2, col_y);

                        for (int i = 0; i < dt2.Rows.Count; ++i)
                        {
                            double z1 = Convert.ToDouble(dt2.Rows[i][col_z]);
                            double x1 = Convert.ToDouble(dt2.Rows[i][col_x]);
                            double y1 = Convert.ToDouble(dt2.Rows[i][col_y]);

                            for (int j = i + 1; j < dt2.Rows.Count; ++j)
                            {
                                double z2 = Convert.ToDouble(dt2.Rows[j][col_z]);
                                double x2 = Convert.ToDouble(dt2.Rows[j][col_x]);
                                double y2 = Convert.ToDouble(dt2.Rows[j][col_y]);
                                string pn2 = Convert.ToString(dt2.Rows[j][col_pn]);

                                double Run = Math.Abs(y2 - y1);

                                if (z2 != z1 && x1 == x2 && Math.Round(Run, 4) == Math.Round(gridV, 4))
                                {
                                    double Rise = z1 - z2;
                                    double Slope = Math.Round(100 * Rise / Run, 4);

                                    dt2.Rows[i][col_slope_SOUTH] = Slope;
                                    dt2.Rows[i][col_SOUTH_PN] = pn2;

                                    j = dt2.Rows.Count;
                                }
                            }


                        }


                        dt2 = Functions.Sort_data_table(dt2, col_x);

                        for (int i = 0; i < dt2.Rows.Count; ++i)
                        {
                            double z1 = Convert.ToDouble(dt2.Rows[i][col_z]);
                            double x1 = Convert.ToDouble(dt2.Rows[i][col_x]);
                            double y1 = Convert.ToDouble(dt2.Rows[i][col_y]);


                            for (int j = i + 1; j < dt2.Rows.Count; ++j)
                            {
                                double z2 = Convert.ToDouble(dt2.Rows[j][col_z]);
                                double x2 = Convert.ToDouble(dt2.Rows[j][col_x]);
                                double y2 = Convert.ToDouble(dt2.Rows[j][col_y]);
                                string pn2 = Convert.ToString(dt2.Rows[j][col_pn]);
                                double Run = Math.Abs(x2 - x1);

                                if (z2 != z1 && y1 == y2 && Math.Round(Run, 4) == Math.Round(gridH, 4))
                                {

                                    double Rise = z1 - z2;
                                    double Slope = Math.Round(100 * Rise / Run, 4);

                                    dt2.Rows[i][col_slope_EAST] = Slope;
                                    dt2.Rows[i][col_EAST_PN] = pn2;

                                    j = dt2.Rows.Count;
                                }
                            }

                        }


                        dt2 = Functions.Sort_data_table_desc(dt2, col_x);

                        for (int i = 0; i < dt2.Rows.Count; ++i)
                        {
                            double z1 = Convert.ToDouble(dt2.Rows[i][col_z]);
                            double x1 = Convert.ToDouble(dt2.Rows[i][col_x]);
                            double y1 = Convert.ToDouble(dt2.Rows[i][col_y]);



                            for (int j = i + 1; j < dt2.Rows.Count; ++j)
                            {
                                double z2 = Convert.ToDouble(dt2.Rows[j][col_z]);
                                double x2 = Convert.ToDouble(dt2.Rows[j][col_x]);
                                double y2 = Convert.ToDouble(dt2.Rows[j][col_y]);
                                string pn2 = Convert.ToString(dt2.Rows[j][col_pn]);

                                double Run = Math.Abs(x2 - x1);

                                if (z2 != z1 && y1 == y2 && Math.Round(Run, 4) == Math.Round(gridH, 4))
                                {
                                    double Rise = z1 - z2;
                                    double Slope = Math.Round(100 * Rise / Run, 4);

                                    dt2.Rows[i][col_slope_WEST] = Slope;
                                    dt2.Rows[i][col_WEST_PN] = pn2;

                                    j = dt2.Rows.Count;
                                }

                            }

                        }






                        if (Functions.IsNumeric(textBox_max_NS.Text) == true)
                        {
                            max_NS1 = Math.Abs(Convert.ToDouble(textBox_max_NS.Text));
                            min_NS2 = -max_NS1;
                        }
                        if (Functions.IsNumeric(textBox_max_EW.Text) == true)
                        {
                            max_EW1 = Math.Abs(Convert.ToDouble(textBox_max_EW.Text));
                            min_EW2 = -max_EW1;
                        }



                        for (int i = 0; i < dt2.Rows.Count; ++i)
                        {
                            double z1 = Convert.ToDouble(dt2.Rows[i][col_z]);
                            string pn1 = Convert.ToString(dt2.Rows[i][col_pn]);

                            double Nslope = 0;

                            if (dt2.Rows[i][col_slope_NORTH] != DBNull.Value)
                            {
                                Nslope = Convert.ToDouble(dt2.Rows[i][col_slope_NORTH]);
                            }

                            double Sslope = 0;

                            if (dt2.Rows[i][col_slope_SOUTH] != DBNull.Value)
                            {
                                Sslope = Convert.ToDouble(dt2.Rows[i][col_slope_SOUTH]);
                            }

                            double NSslope = Nslope;
                            string NSpoint = Convert.ToString(dt2.Rows[i][col_NORTH_PN]);

                            if (Math.Abs(Sslope) > Math.Abs(Nslope))
                            {
                                NSslope = Sslope;
                                NSpoint = Convert.ToString(dt2.Rows[i][col_SOUTH_PN]);
                            }


                            if (NSslope >= min_NS2 && NSslope <= max_NS1)
                            {

                            }
                            else
                            {
                                double calc_slope = max_NS1;

                                if (NSslope < 0) calc_slope = min_NS2;
                                double x1 = Convert.ToDouble(dt2.Rows[i][col_x]);
                                double y1 = Convert.ToDouble(dt2.Rows[i][col_y]);

                                bool processed1 = Convert.ToBoolean(dt2.Rows[i][col_NS_processed]);
                                if (processed1 == false)
                                {
                                    for (int j = 0; j < dt2.Rows.Count; ++j)
                                    {
                                        string pn2 = Convert.ToString(dt2.Rows[j][col_pn]);
                                        if (pn2 == NSpoint)
                                        {
                                            double z2 = Convert.ToDouble(dt2.Rows[j][col_z]);
                                            double x2 = Convert.ToDouble(dt2.Rows[j][col_x]);
                                            double y2 = Convert.ToDouble(dt2.Rows[j][col_y]);
                                            double run = Math.Abs(y2 - y1);

                                            dt2.Rows[i][col_newZ_NS] = z2 + calc_slope * run / 100;
                                            dt2.Rows[i][col_max_NS] = calc_slope;
                                            dt2.Rows[i][col_NS_processed] = true;
                                            dt2.Rows[j][col_NS_processed] = true;
                                            j = dt2.Rows.Count;
                                        }
                                    }
                                }
                            }


                            double Eslope = 0;
                            if (dt2.Rows[i][col_slope_EAST] != DBNull.Value)
                            {
                                Eslope = Convert.ToDouble(dt2.Rows[i][col_slope_EAST]);
                            }

                            double Wslope = 0;

                            if (dt2.Rows[i][col_slope_WEST] != DBNull.Value)
                            {
                                Wslope = Convert.ToDouble(dt2.Rows[i][col_slope_WEST]);
                            }

                            double EWslope = Eslope;
                            string EWpoint = Convert.ToString(dt2.Rows[i][col_EAST_PN]);

                            if (Math.Abs(Wslope) > Math.Abs(Eslope))
                            {
                                EWslope = Wslope;
                                EWpoint = Convert.ToString(dt2.Rows[i][col_WEST_PN]);
                            }


                            if (EWslope >= min_EW2 && EWslope <= max_EW1)
                            {

                            }
                            else
                            {
                                double calc_slope = max_EW1;

                                if (EWslope < 0) calc_slope = min_EW2;

                                double x1 = Convert.ToDouble(dt2.Rows[i][col_x]);
                                double y1 = Convert.ToDouble(dt2.Rows[i][col_y]);
                                bool processed1 = Convert.ToBoolean(dt2.Rows[i][col_EW_processed]);
                                if (processed1 == false)
                                {
                                    for (int j = 0; j < dt2.Rows.Count; ++j)
                                    {
                                        string pn2 = Convert.ToString(dt2.Rows[j][col_pn]);
                                        if (pn2 == EWpoint)
                                        {
                                            double z2 = Convert.ToDouble(dt2.Rows[j][col_z]);
                                            double x2 = Convert.ToDouble(dt2.Rows[j][col_x]);
                                            double y2 = Convert.ToDouble(dt2.Rows[j][col_y]);
                                            double run = Math.Abs(x2 - x1);
                                            dt2.Rows[i][col_newZ_EW] = z2 + calc_slope * run / 100;
                                            dt2.Rows[i][col_max_EW] = calc_slope;
                                            dt2.Rows[i][col_EW_processed] = true;
                                            dt2.Rows[j][col_EW_processed] = true;
                                            j = dt2.Rows.Count;
                                        }
                                    }
                                }
                            }


                            double delta_NS = 0;
                            double delta_EW = 0;
                            if (dt2.Rows[i][col_newZ_NS] != DBNull.Value)
                            {
                                delta_NS = Math.Abs(z1 - Convert.ToDouble(dt2.Rows[i][col_newZ_NS]));
                            }

                            if (dt2.Rows[i][col_newZ_EW] != DBNull.Value)
                            {
                                delta_EW = Math.Abs(z1 - Convert.ToDouble(dt2.Rows[i][col_newZ_EW]));
                            }



                            if (delta_EW > delta_NS)
                            {

                                dt2.Rows[i][col_new_elev] = dt2.Rows[i][col_newZ_EW];
                            }
                            if (delta_EW < delta_NS)
                            {
                                dt2.Rows[i][col_new_elev] = dt2.Rows[i][col_newZ_NS];
                            }

                            if (delta_EW == delta_NS && delta_NS > 0)
                            {
                                dt2.Rows[i][col_new_elev] = dt2.Rows[i][col_newZ_NS];
                            }

                            if (dt2.Rows[i][col_new_elev] != DBNull.Value)
                            {
                                double z2 = Convert.ToDouble(dt2.Rows[i][col_new_elev]);
                                if (z2 > z1)
                                {
                                    dt2.Rows[i][col_new_description] = "FILL";
                                }
                                else
                                {
                                    dt2.Rows[i][col_new_description] = "CUT";

                                }
                            }

                            string new_descr = Convert.ToString(dt2.Rows[i][col_new_description]);
                            ObjectId id1 = (ObjectId)dt2.Rows[i][col_id];
                            CogoPoint cg1 = Trans1.GetObject(id1, OpenMode.ForWrite) as CogoPoint;
                           
                            if (new_descr == "FILL" || new_descr == "CUT")
                            {
                                cg1.RawDescription = new_descr;
                                Autodesk.Civil.DatabaseServices.Styles.PointStyleCollection col_point_styles = CivilDrawing.Styles.PointStyles;
                                IEnumerator<ObjectId> enum1 = col_point_styles.GetEnumerator();

                                while (enum1.MoveNext())
                                {
                                    Autodesk.Civil.DatabaseServices.Styles.PointStyle pst1 = Trans1.GetObject(enum1.Current, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.PointStyle;

                                    if (new_descr == "FILL")
                                    {
                                        if (pst1.Name == comboBox_point_style1.Text)
                                        {
                                            cg1.StyleId = pst1.ObjectId;
                                        }
                                    }

                                    if (new_descr == "CUT")
                                    {
                                        if (pst1.Name == comboBox_point_style2.Text)
                                        {
                                            cg1.StyleId = pst1.ObjectId;
                                        }
                                    }
                                }

                                Autodesk.Civil.DatabaseServices.Styles.LabelStyleCollection col_point_label_styles = CivilDrawing.Styles.LabelStyles.PointLabelStyles.LabelStyles;

                                IEnumerator<ObjectId> enum2 = col_point_label_styles.GetEnumerator();

                                while (enum2.MoveNext())
                                {
                                    Autodesk.Civil.DatabaseServices.Styles.LabelStyle plst2 = Trans1.GetObject(enum2.Current, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.LabelStyle;

                                    if (new_descr == "FILL")
                                    {
                                        if (plst2.Name == comboBox_point_label_style1.Text)
                                        {
                                            cg1.LabelStyleId = plst2.ObjectId;
                                        }
                                    }
                                    if (new_descr == "CUT")
                                    {
                                        if (plst2.Name == comboBox_point_label_style2.Text)
                                        {
                                            cg1.LabelStyleId = plst2.ObjectId;
                                        }
                                    }
                                }

                                string layer_pt = "0";
                                if (comboBox_point_layer1.Text.Length > 0)
                                {
                                    if (new_descr == "FILL")
                                    {
                                        layer_pt = comboBox_point_layer1.Text;
                                    }
                                    if (new_descr == "CUT")
                                    {
                                        layer_pt = comboBox_point_layer2.Text;
                                    }
                                }
                                cg1.Layer = layer_pt;




                            }

                            string udp1_string = "";
                            double udp2_double = -1.234;
                            UDPString udp1 = Functions.Find_udp_string(comboBox_udp_field_1.Text);
                            if (udp1 != null)
                            {
                                udp1_string = Convert.ToString(cg1.GetUDPValue(udp1));
                            }
                            UDPDouble udp2 = Functions.Find_udp_double(comboBox_udp_field_2.Text);
                            if (udp2 != null)
                            {
                                udp2_double = Convert.ToDouble(cg1.GetUDPValue(udp2));
                            }

                            if (checkBox_output_to_excel1.Checked == true)
                            {
                                dt3.Rows.Add();
                                dt3.Rows[dt3.Rows.Count - 1][col_pn] = pn1;
                                dt3.Rows[dt3.Rows.Count - 1][col_y] = cg1.Northing;
                                dt3.Rows[dt3.Rows.Count - 1][col_x] = cg1.Easting;
                                dt3.Rows[dt3.Rows.Count - 1][col_z] = cg1.Elevation;
                                dt3.Rows[dt3.Rows.Count - 1][col_descr] = cg1.RawDescription;



                                if (udp1_string != "")
                                {
                                    dt3.Rows[dt3.Rows.Count - 1][col_new_description] = udp1_string;
                                }

                                if (udp2_double != -1.234)
                                {
                                    dt3.Rows[dt3.Rows.Count - 1][col_new_elev] = udp2_double;
                                }
                            }


                        }

                        dt2 = Functions.Sort_data_table(dt2, col_y);

                        for (int i = 0; i < dt2.Rows.Count; ++i)
                        {
                            if (dt2.Rows[i][col_new_elev] != DBNull.Value && dt2.Rows[i][col_NORTH_PN] != DBNull.Value)
                            {
                                double z1 = Convert.ToDouble(dt2.Rows[i][col_new_elev]);
                                double x1 = Convert.ToDouble(dt2.Rows[i][col_x]);
                                double y1 = Convert.ToDouble(dt2.Rows[i][col_y]);
                                string pn1 = Convert.ToString(dt2.Rows[i][col_NORTH_PN]);


                                for (int j = i + 1; j < dt2.Rows.Count; ++j)
                                {
                                    string pn2 = Convert.ToString(dt2.Rows[j][col_pn]);

                                    if (pn1 == pn2)
                                    {
                                        double z2 = Convert.ToDouble(dt2.Rows[j][col_z]);
                                        double x2 = Convert.ToDouble(dt2.Rows[j][col_x]);
                                        double y2 = Convert.ToDouble(dt2.Rows[j][col_y]);
                                        if (dt2.Rows[j][col_new_elev] != DBNull.Value)
                                        {
                                            z2 = Convert.ToDouble(dt2.Rows[j][col_new_elev]);
                                        }
                                        double Rise = z1 - z2;
                                        double Run = Math.Abs(y2 - y1);
                                        double Slope = Math.Round(100 * Rise / Run, 4);

                                        dt2.Rows[i][col_slopeN1] = Slope;

                                        j = dt2.Rows.Count;
                                    }
                                }
                            }
                        }



                        dt2 = Functions.Sort_data_table_desc(dt2, col_y);

                        for (int i = 0; i < dt2.Rows.Count; ++i)
                        {
                            if (dt2.Rows[i][col_new_elev] != DBNull.Value && dt2.Rows[i][col_SOUTH_PN] != DBNull.Value)
                            {
                                double z1 = Convert.ToDouble(dt2.Rows[i][col_new_elev]);
                                double x1 = Convert.ToDouble(dt2.Rows[i][col_x]);
                                double y1 = Convert.ToDouble(dt2.Rows[i][col_y]);
                                string pn1 = Convert.ToString(dt2.Rows[i][col_SOUTH_PN]);

                                for (int j = i + 1; j < dt2.Rows.Count; ++j)
                                {
                                    string pn2 = Convert.ToString(dt2.Rows[j][col_pn]);

                                    if (pn1 == pn2)
                                    {
                                        double z2 = Convert.ToDouble(dt2.Rows[j][col_z]);
                                        double x2 = Convert.ToDouble(dt2.Rows[j][col_x]);
                                        double y2 = Convert.ToDouble(dt2.Rows[j][col_y]);
                                        if (dt2.Rows[j][col_new_elev] != DBNull.Value)
                                        {
                                            z2 = Convert.ToDouble(dt2.Rows[j][col_new_elev]);
                                        }
                                        double Rise = z1 - z2;
                                        double Run = Math.Abs(y2 - y1);
                                        double Slope = Math.Round(100 * Rise / Run, 4);

                                        dt2.Rows[i][col_slopeS1] = Slope;

                                        j = dt2.Rows.Count;
                                    }
                                }
                            }


                        }


                        dt2 = Functions.Sort_data_table(dt2, col_x);

                        for (int i = 0; i < dt2.Rows.Count; ++i)
                        {

                            if (dt2.Rows[i][col_new_elev] != DBNull.Value && dt2.Rows[i][col_EAST_PN] != DBNull.Value)
                            {
                                double z1 = Convert.ToDouble(dt2.Rows[i][col_new_elev]);
                                double x1 = Convert.ToDouble(dt2.Rows[i][col_x]);
                                double y1 = Convert.ToDouble(dt2.Rows[i][col_y]);
                                string pn1 = Convert.ToString(dt2.Rows[i][col_EAST_PN]);

                                for (int j = i + 1; j < dt2.Rows.Count; ++j)
                                {
                                    string pn2 = Convert.ToString(dt2.Rows[j][col_pn]);
                                    if (pn1 == pn2)
                                    {
                                        double z2 = Convert.ToDouble(dt2.Rows[j][col_z]);
                                        double x2 = Convert.ToDouble(dt2.Rows[j][col_x]);
                                        double y2 = Convert.ToDouble(dt2.Rows[j][col_y]);
                                        if (dt2.Rows[j][col_new_elev] != DBNull.Value)
                                        {
                                            z2 = Convert.ToDouble(dt2.Rows[j][col_new_elev]);
                                        }
                                        double Rise = z1 - z2;
                                        double Run = Math.Abs(x2 - x1);
                                        double Slope = Math.Round(100 * Rise / Run, 4);
                                        dt2.Rows[i][col_slopeE1] = Slope;
                                        j = dt2.Rows.Count;
                                    }
                                }
                            }
                        }


                        dt2 = Functions.Sort_data_table_desc(dt2, col_x);

                        for (int i = 0; i < dt2.Rows.Count; ++i)
                        {
                            if (dt2.Rows[i][col_new_elev] != DBNull.Value && dt2.Rows[i][col_WEST_PN] != DBNull.Value)
                            {
                                double z1 = Convert.ToDouble(dt2.Rows[i][col_new_elev]);
                                double x1 = Convert.ToDouble(dt2.Rows[i][col_x]);
                                double y1 = Convert.ToDouble(dt2.Rows[i][col_y]);
                                string pn1 = Convert.ToString(dt2.Rows[i][col_WEST_PN]);

                                for (int j = i + 1; j < dt2.Rows.Count; ++j)
                                {

                                    string pn2 = Convert.ToString(dt2.Rows[j][col_pn]);

                                    if (pn1 == pn2)
                                    {
                                        double z2 = Convert.ToDouble(dt2.Rows[j][col_z]);
                                        double x2 = Convert.ToDouble(dt2.Rows[j][col_x]);
                                        double y2 = Convert.ToDouble(dt2.Rows[j][col_y]);
                                        if (dt2.Rows[j][col_new_elev] != DBNull.Value)
                                        {
                                            z2 = Convert.ToDouble(dt2.Rows[j][col_new_elev]);
                                        }

                                        double Rise = z1 - z2;
                                        double Run = Math.Abs(x2 - x1);
                                        double Slope = Math.Round(100 * Rise / Run, 4);

                                        dt2.Rows[i][col_slopeW1] = Slope;

                                        j = dt2.Rows.Count;
                                    }
                                }
                            }
                        }

                        using (System.Data.DataTable Data_table_temp = dt2.Clone())
                        {
                            DataView dv = new DataView(dt2);
                            dv.Sort = col_x + "," + col_y;
                            for (int i = 0; i < dv.Count; ++i)
                            {
                                System.Data.DataRow Data_row1 = dv[i].Row;
                                Data_table_temp.Rows.Add();
                                for (int j = 0; j < dt2.Columns.Count; ++j)
                                {
                                    Data_table_temp.Rows[Data_table_temp.Rows.Count - 1][j] = Data_row1[j];
                                }
                            }

                            dt2 = Data_table_temp.Copy();
                            dt2.Columns.Remove(col_EW_processed);
                            dt2.Columns.Remove(col_NS_processed);
                            dt2.Columns.Remove(col_id);
                        }

                        string nume1 = Environment.UserName + " " + DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + " at " + DateTime.Now.Hour + "h" + DateTime.Now.Minute + "m";

                        Microsoft.Office.Interop.Excel.Worksheet W1 = Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt2);
                        W1.Name = nume1;

                        int hide1 = 10;
                        W1.Range["A1:W1"].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        W1.Range["A1:W1"].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                        W1.Range["A1:W1"].WrapText = true;
                        W1.Range["A:A"].ColumnWidth = 13;
                        W1.Range["B:E"].ColumnWidth = 11;
                        W1.Range["F:F"].ColumnWidth = 10;
                        W1.Range["G:G"].ColumnWidth = hide1;
                        W1.Range["H:H"].ColumnWidth = 10;
                        W1.Range["I:I"].ColumnWidth = hide1;
                        W1.Range["J:J"].ColumnWidth = 10;
                        W1.Range["K:K"].ColumnWidth = hide1;
                        W1.Range["L:L"].ColumnWidth = 10;
                        W1.Range["M:Q"].ColumnWidth = hide1;
                        W1.Range["R:V"].ColumnWidth = 10;
                        W1.Range["W:W"].ColumnWidth = 12;


                        dt1.Dispose();
                        dt2.Dispose();

                        Trans1.Commit();




                        Microsoft.Office.Interop.Excel.Worksheet W2 = Functions.Transfer_datatable_to_new_excel_spreadsheet(dt3, nume1);
                        if (W2 != null)
                        {
                            W2.Rows[1].Delete();
                        }



                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            set_enable_true();
            this.MdiParent.WindowState = FormWindowState.Normal;
            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
        }

    


        private void button_load_point_style_Click(object sender, EventArgs e)
        {
            comboBox_point_style1.Items.Clear();
            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.Civil.ApplicationServices.CivilDocument CivilDrawing = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            try
            {
                set_enable_false();
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.Civil.DatabaseServices.Styles.PointStyleCollection col_point_styles = CivilDrawing.Styles.PointStyles;
                        IEnumerator<ObjectId> enum1 = col_point_styles.GetEnumerator();
                        while (enum1.MoveNext())
                        {
                            Autodesk.Civil.DatabaseServices.Styles.PointStyle pst1 = Trans1.GetObject(enum1.Current, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.PointStyle;
                            comboBox_point_style1.Items.Add(pst1.Name);
                        }
                        if (comboBox_point_style1.Items.Contains("Fill") == true)
                        {
                            comboBox_point_style1.SelectedIndex = comboBox_point_style1.Items.IndexOf("Fill");
                        }
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


        private void button_load_point_label_style_Click(object sender, EventArgs e)
        {
            comboBox_point_label_style1.Items.Clear();

            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.Civil.ApplicationServices.CivilDocument CivilDrawing = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            try
            {
                set_enable_false();
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.Civil.DatabaseServices.Styles.LabelStyleCollection col_point_label_styles = CivilDrawing.Styles.LabelStyles.PointLabelStyles.LabelStyles;
                        IEnumerator<ObjectId> enum2 = col_point_label_styles.GetEnumerator();

                        while (enum2.MoveNext())
                        {
                            Autodesk.Civil.DatabaseServices.Styles.LabelStyle plst2 = Trans1.GetObject(enum2.Current, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.LabelStyle;
                            comboBox_point_label_style1.Items.Add(plst2.Name);
                        }

                        if (comboBox_point_label_style1.Items.Contains("Point#-Elevation-Description [FILL]") == true)
                        {
                            comboBox_point_label_style1.SelectedIndex = comboBox_point_label_style1.Items.IndexOf("Point#-Elevation-Description [FILL]");
                        }
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

        private void button_load_layers_to_combobox_point_layer_Click(object sender, EventArgs e)
        {
            Functions.Incarca_existing_layers_to_combobox(comboBox_point_layer1);

            if (comboBox_point_layer1.Items.Contains("PTS-Fill") == true)
            {
                comboBox_point_layer1.SelectedIndex = comboBox_point_layer1.Items.IndexOf("PTS-Fill");
            }

        }

        private void button_load_point_style2_Click(object sender, EventArgs e)
        {
            comboBox_point_style2.Items.Clear();
            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.Civil.ApplicationServices.CivilDocument CivilDrawing = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            try
            {
                set_enable_false();
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.Civil.DatabaseServices.Styles.PointStyleCollection col_point_styles = CivilDrawing.Styles.PointStyles;
                        IEnumerator<ObjectId> enum1 = col_point_styles.GetEnumerator();
                        while (enum1.MoveNext())
                        {
                            Autodesk.Civil.DatabaseServices.Styles.PointStyle pst1 = Trans1.GetObject(enum1.Current, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.PointStyle;
                            comboBox_point_style2.Items.Add(pst1.Name);
                        }
                        if (comboBox_point_style2.Items.Contains("Cut") == true)
                        {
                            comboBox_point_style2.SelectedIndex = comboBox_point_style2.Items.IndexOf("Cut");
                        }
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

        private void button_load_point_label_style2_Click(object sender, EventArgs e)
        {
            comboBox_point_label_style2.Items.Clear();

            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.Civil.ApplicationServices.CivilDocument CivilDrawing = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            try
            {
                set_enable_false();
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.Civil.DatabaseServices.Styles.LabelStyleCollection col_point_label_styles = CivilDrawing.Styles.LabelStyles.PointLabelStyles.LabelStyles;
                        IEnumerator<ObjectId> enum2 = col_point_label_styles.GetEnumerator();

                        while (enum2.MoveNext())
                        {
                            Autodesk.Civil.DatabaseServices.Styles.LabelStyle plst2 = Trans1.GetObject(enum2.Current, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.LabelStyle;
                            comboBox_point_label_style2.Items.Add(plst2.Name);
                        }

                        if (comboBox_point_label_style2.Items.Contains("Point#-Elevation-Description [CUT]") == true)
                        {
                            comboBox_point_label_style2.SelectedIndex = comboBox_point_label_style2.Items.IndexOf("Point#-Elevation-Description [CUT]");
                        }
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

        private void button_load_layers_to_combobox_point_layer2_Click(object sender, EventArgs e)
        {
            Functions.Incarca_existing_layers_to_combobox(comboBox_point_layer2);

            if (comboBox_point_layer2.Items.Contains("PTS-Cut") == true)
            {
                comboBox_point_layer2.SelectedIndex = comboBox_point_layer2.Items.IndexOf("PTS-Cut");
            }

        }


        private void button_udp1_Click(object sender, EventArgs e)
        {
            comboBox_udp_field_1.Items.Clear();
            foreach (UDP udp1 in Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument.PointUDPs)
            {
                if(udp1!=null)
                {
comboBox_udp_field_1.Items.Add(udp1.Name);
                }
                
            }
            if (comboBox_udp_field_1.Items.Contains("ATT1") == true)
            {
                comboBox_udp_field_1.SelectedIndex = comboBox_udp_field_1.Items.IndexOf("ATT1");
            }
        }
        private void button_udp2_Click(object sender, EventArgs e)
        {
            comboBox_udp_field_2.Items.Clear();
            foreach (UDP udp2 in Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument.PointUDPs)
            {
                if (udp2 != null)
                {
                    comboBox_udp_field_2.Items.Add(udp2.Name);
                }
            }
            if (comboBox_udp_field_2.Items.Contains("ATT2") == true)
            {
                comboBox_udp_field_2.SelectedIndex = comboBox_udp_field_2.Items.IndexOf("ATT2");
            }
        }

        private void button_global_load_Click(object sender, EventArgs e)
        {
            comboBox_point_style1.Items.Clear();
            comboBox_point_label_style1.Items.Clear();
            comboBox_point_style2.Items.Clear();
            comboBox_point_label_style2.Items.Clear();

            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.Civil.ApplicationServices.CivilDocument CivilDrawing = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            try
            {
                set_enable_false();
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.Civil.DatabaseServices.Styles.PointStyleCollection col_point_styles = CivilDrawing.Styles.PointStyles;
                        IEnumerator<ObjectId> enum1 = col_point_styles.GetEnumerator();
                        while (enum1.MoveNext())
                        {
                            Autodesk.Civil.DatabaseServices.Styles.PointStyle pst1 = Trans1.GetObject(enum1.Current, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.PointStyle;
                            comboBox_point_style1.Items.Add(pst1.Name);
                            comboBox_point_style2.Items.Add(pst1.Name);
                        }
                        if (comboBox_point_style1.Items.Contains("Fill") == true)
                        {
                            comboBox_point_style1.SelectedIndex = comboBox_point_style1.Items.IndexOf("Fill");
                        }
                        if (comboBox_point_style2.Items.Contains("Cut") == true)
                        {
                            comboBox_point_style2.SelectedIndex = comboBox_point_style2.Items.IndexOf("Cut");
                        }
                        Autodesk.Civil.DatabaseServices.Styles.LabelStyleCollection col_point_label_styles = CivilDrawing.Styles.LabelStyles.PointLabelStyles.LabelStyles;
                        IEnumerator<ObjectId> enum2 = col_point_label_styles.GetEnumerator();

                        while (enum2.MoveNext())
                        {
                            Autodesk.Civil.DatabaseServices.Styles.LabelStyle plst2 = Trans1.GetObject(enum2.Current, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.LabelStyle;
                            comboBox_point_label_style1.Items.Add(plst2.Name);
                            comboBox_point_label_style2.Items.Add(plst2.Name);
                        }
                        if (comboBox_point_label_style1.Items.Contains("Point#-Elevation-Description [FILL]") == true)
                        {
                            comboBox_point_label_style1.SelectedIndex = comboBox_point_label_style1.Items.IndexOf("Point#-Elevation-Description [FILL]");
                        }
                        if (comboBox_point_label_style2.Items.Contains("Point#-Elevation-Description [CUT]") == true)
                        {
                            comboBox_point_label_style2.SelectedIndex = comboBox_point_label_style2.Items.IndexOf("Point#-Elevation-Description [CUT]");
                        }

                        comboBox_udp_field_1.Items.Clear();
                        foreach (UDP udp1 in Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument.PointUDPs)
                        {
                            if (udp1 != null)
                            {
                                comboBox_udp_field_1.Items.Add(udp1.Name);
                            }
                        }
                        if (comboBox_udp_field_1.Items.Contains("ATT1") == true)
                        {
                            comboBox_udp_field_1.SelectedIndex = comboBox_udp_field_1.Items.IndexOf("ATT1");
                        }

                        comboBox_udp_field_2.Items.Clear();
                        foreach (UDP udp2 in Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument.PointUDPs)
                        {
                            if (udp2 != null)
                            {
                                comboBox_udp_field_2.Items.Add(udp2.Name);
                            }
                        }
                        if (comboBox_udp_field_2.Items.Contains("ATT2") == true)
                        {
                            comboBox_udp_field_2.SelectedIndex = comboBox_udp_field_2.Items.IndexOf("ATT2");
                        }

                        Functions.Incarca_existing_layers_to_combobox(comboBox_point_layer1);

                        if (comboBox_point_layer1.Items.Contains("PTS-Fill") == true)
                        {
                            comboBox_point_layer1.SelectedIndex = comboBox_point_layer1.Items.IndexOf("PTS-Fill");
                        }
                        Functions.Incarca_existing_layers_to_combobox(comboBox_point_layer2);

                        if (comboBox_point_layer2.Items.Contains("PTS-Cut") == true)
                        {
                            comboBox_point_layer2.SelectedIndex = comboBox_point_layer2.Items.IndexOf("PTS-Cut");
                        }


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
    }
}
