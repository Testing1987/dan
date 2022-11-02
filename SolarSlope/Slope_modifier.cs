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
    public partial class Slope_modify_form : Form
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
        string col_orig_elev = "Original\r\nElevation";
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

        string col_col = "col";
        string col_row = "row";
        string col_dist = "Distance";
        ObjectIdCollection col_sel_ids = null;
        System.Data.DataTable dt_analize = null;

        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();


            lista_butoane.Add(button_adjust);
            lista_butoane.Add(button_load_layers_to_combobox1);
            lista_butoane.Add(button_load_layers_to_combobox2);
            lista_butoane.Add(button_load_layers_to_combobox3);
            lista_butoane.Add(button_load_layers_to_combobox4);
            lista_butoane.Add(button_load_layers_to_combobox5);
            lista_butoane.Add(button_load_point_label_style1);
            lista_butoane.Add(button_load_point_label_style2);
            lista_butoane.Add(button_load_point_label_style3);
            lista_butoane.Add(button_load_point_label_style4);
            lista_butoane.Add(button_load_point_label_style5);
            lista_butoane.Add(button_load_point_style1);
            lista_butoane.Add(button_load_point_style2);
            lista_butoane.Add(button_load_point_style3);
            lista_butoane.Add(button_load_point_style4);
            lista_butoane.Add(button_load_point_style5);
            lista_butoane.Add(button_global_load);
            lista_butoane.Add(button_udp1);
            lista_butoane.Add(button_udp2);
            lista_butoane.Add(button_reset_to_dt_analise);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {

                bt1.Enabled = false;

            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();

            lista_butoane.Add(button_adjust);
            lista_butoane.Add(button_load_layers_to_combobox1);
            lista_butoane.Add(button_load_layers_to_combobox2);
            lista_butoane.Add(button_load_layers_to_combobox3);
            lista_butoane.Add(button_load_layers_to_combobox4);
            lista_butoane.Add(button_load_layers_to_combobox5);
            lista_butoane.Add(button_load_point_label_style1);
            lista_butoane.Add(button_load_point_label_style2);
            lista_butoane.Add(button_load_point_label_style3);
            lista_butoane.Add(button_load_point_label_style4);
            lista_butoane.Add(button_load_point_label_style5);
            lista_butoane.Add(button_load_point_style1);
            lista_butoane.Add(button_load_point_style2);
            lista_butoane.Add(button_load_point_style3);
            lista_butoane.Add(button_load_point_style4);
            lista_butoane.Add(button_load_point_style5);
            lista_butoane.Add(button_global_load);
            lista_butoane.Add(button_udp1);
            lista_butoane.Add(button_udp2);
            lista_butoane.Add(button_reset_to_dt_analise);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }

        public Slope_modify_form()
        {
            InitializeComponent();
            var toolStripMenuItem2 = new ToolStripMenuItem { Text = "tool strip menu" };
            //toolStripMenuItem2.Click += go_to_excel_point;


            ContextMenuStrip_go_to_error = new ContextMenuStrip();
            ContextMenuStrip_go_to_error.Items.AddRange(new ToolStripItem[] { toolStripMenuItem2 });


        }



        private void button_adjust_Click(object sender, EventArgs e)
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

                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_cogopts = null;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez.MessageForAdding = "\nSelect cogo points";
                        Prompt_rez.SingleOnly = false;

                        set_enable_false();
                        this.MdiParent.WindowState = FormWindowState.Minimized;

                        if (col_sel_ids == null)
                        {
                            Rezultat_cogopts = ThisDrawing.Editor.GetSelection(Prompt_rez);

                            if (Rezultat_cogopts.Status != PromptStatus.OK)
                            {
                                set_enable_true();
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                this.MdiParent.WindowState = FormWindowState.Normal;
                                return;
                            }
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
                        dt1.Columns.Add(col_anchor, typeof(bool));
                        dt1.Columns.Add(col_id, typeof(ObjectId));
                        dt1.Columns.Add(col_orig_elev, typeof(double));

                        #region get styles and layers

                        Autodesk.Civil.DatabaseServices.Styles.PointStyleCollection col_point_styles = CivilDrawing.Styles.PointStyles;
                        IEnumerator<ObjectId> enum1 = col_point_styles.GetEnumerator();

                        Autodesk.Civil.DatabaseServices.Styles.PointStyle pst1 = null;
                        Autodesk.Civil.DatabaseServices.Styles.PointStyle pst2 = null;

                        while (enum1.MoveNext())
                        {
                            Autodesk.Civil.DatabaseServices.Styles.PointStyle pst3 = Trans1.GetObject(enum1.Current, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.PointStyle;
                            if (pst3.Name == comboBox_point_style1.Text)
                            {
                                pst1 = pst3;
                            }
                            if (pst3.Name == comboBox_point_style2.Text)
                            {
                                pst2 = pst3;
                            }
                        }

                        Autodesk.Civil.DatabaseServices.Styles.LabelStyleCollection col_point_label_styles = CivilDrawing.Styles.LabelStyles.PointLabelStyles.LabelStyles;
                        Autodesk.Civil.DatabaseServices.Styles.LabelStyle plst1 = null;
                        Autodesk.Civil.DatabaseServices.Styles.LabelStyle plst2 = null;

                        IEnumerator<ObjectId> enum2 = col_point_label_styles.GetEnumerator();

                        while (enum2.MoveNext())
                        {
                            Autodesk.Civil.DatabaseServices.Styles.LabelStyle plst3 = Trans1.GetObject(enum2.Current, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.LabelStyle;
                            if (plst3.Name == comboBox_point_label_style1.Text)
                            {
                                plst1 = plst3;
                            }
                            if (plst3.Name == comboBox_point_label_style2.Text)
                            {
                                plst2 = plst3;
                            }
                        }

                        string layer_pt1 = "0";
                        if (comboBox_point_layer1.Text.Length > 0)
                        {
                            layer_pt1 = comboBox_point_layer1.Text;
                        }
                        string layer_pt2 = "0";
                        if (comboBox_point_layer2.Text.Length > 0)
                        {
                            layer_pt2 = comboBox_point_layer2.Text;
                        }

                        #endregion

                        if (col_sel_ids == null && Rezultat_cogopts != null)
                        {
                            for (int i = 0; i < Rezultat_cogopts.Value.Count; ++i)
                            {
                                CogoPoint cg1 = Trans1.GetObject(Rezultat_cogopts.Value[i].ObjectId, OpenMode.ForRead) as CogoPoint;
                                if (cg1 != null)
                                {
                                    if (col_sel_ids == null) col_sel_ids = new ObjectIdCollection();
                                    col_sel_ids.Add(cg1.ObjectId);
                                }
                            }
                        }


                        for (int i = 0; i < col_sel_ids.Count; ++i)
                        {
                            CogoPoint cg1 = Trans1.GetObject(col_sel_ids[i], OpenMode.ForRead) as CogoPoint;
                            if (cg1 != null)
                            {
                                dt1.Rows.Add();
                                dt1.Rows[dt1.Rows.Count - 1][col_pn] = cg1.PointNumber;

                                dt1.Rows[dt1.Rows.Count - 1][col_descr] = cg1.RawDescription;
                                dt1.Rows[dt1.Rows.Count - 1][col_x] = Math.Round(cg1.Location.X, 4);
                                dt1.Rows[dt1.Rows.Count - 1][col_y] = Math.Round(cg1.Location.Y, 4);
                                dt1.Rows[dt1.Rows.Count - 1][col_z] = Math.Round(cg1.Elevation, 4);
                                dt1.Rows[dt1.Rows.Count - 1][col_new_description] = dt1.Rows[dt1.Rows.Count - 1][col_descr];
                                dt1.Rows[dt1.Rows.Count - 1][col_NS_processed] = false;
                                dt1.Rows[dt1.Rows.Count - 1][col_EW_processed] = false;
                                dt1.Rows[dt1.Rows.Count - 1][col_id] = cg1.ObjectId;
                                if (cg1.RawDescription == "ANCHOR")
                                {
                                    dt1.Rows[dt1.Rows.Count - 1][col_anchor] = true;
                                    if (radioButton_EW.Checked == true)
                                    {
                                        dt1.Rows[dt1.Rows.Count - 1][col_EW_processed] = true;
                                        dt1.Rows[dt1.Rows.Count - 1][col_new_elev] = dt1.Rows[dt1.Rows.Count - 1][col_z];
                                    }
                                    if (radioButton_NS.Checked == true)
                                    {
                                        dt1.Rows[dt1.Rows.Count - 1][col_NS_processed] = true;

                                        dt1.Rows[dt1.Rows.Count - 1][col_new_elev] = dt1.Rows[dt1.Rows.Count - 1][col_z];
                                    }
                                }
                                else
                                {
                                    dt1.Rows[dt1.Rows.Count - 1][col_anchor] = false;
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

                                if (udp2_double != -1.234)
                                {
                                    dt1.Rows[dt1.Rows.Count - 1][col_orig_elev] = Math.Round(udp2_double, 4);
                                }
                            }
                        }

                        #region read existing slopes
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
                        #endregion




                        if (radioButton_NS.Checked == true)
                        {

                            if (Functions.IsNumeric(textBox_max_NS.Text) == true)
                            {
                                max_NS1 = Math.Abs(Convert.ToDouble(textBox_max_NS.Text));
                                min_NS2 = -max_NS1;
                            }
                            else
                            {
                                MessageBox.Show("maximum slope north-south is not specified", "SLOPEZ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                set_enable_true();
                                this.MdiParent.WindowState = FormWindowState.Normal;
                                Editor1.SetImpliedSelection(Empty_array);
                                Editor1.WriteMessage("\nCommand:");
                            }
                        }

                        if (radioButton_EW.Checked == true)
                        {
                            if (Functions.IsNumeric(textBox_max_EW.Text) == true)
                            {
                                max_EW1 = Math.Abs(Convert.ToDouble(textBox_max_EW.Text));
                                min_EW2 = -max_EW1;
                            }
                            else
                            {
                                MessageBox.Show("maximum slope east-west is not specified", "SLOPEZ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                set_enable_true();
                                this.MdiParent.WindowState = FormWindowState.Normal;
                                Editor1.SetImpliedSelection(Empty_array);
                                Editor1.WriteMessage("\nCommand:");
                            }
                        }

                        if (radioButton_EW.Checked == true)
                        {

                            for (int i = 0; i < dt2.Rows.Count; ++i)
                            {
                                if (dt2.Rows[i][col_anchor] != DBNull.Value && Convert.ToBoolean(dt2.Rows[i][col_anchor]) == true)
                                {
                                    if (dt2.Rows[i][col_EW_processed] != DBNull.Value && Convert.ToBoolean(dt2.Rows[i][col_EW_processed]) == true)
                                    {
                                        double x1 = Convert.ToDouble(dt2.Rows[i][col_x]);
                                        double y1 = Convert.ToDouble(dt2.Rows[i][col_y]);
                                        double z1 = Convert.ToDouble(dt2.Rows[i][col_z]);
                                        string pn1 = Convert.ToString(dt2.Rows[i][col_pn]);
                                        ObjectId id1 = (ObjectId)dt2.Rows[i][col_id];
                                        CogoPoint cg_anchor = Trans1.GetObject(id1, OpenMode.ForWrite) as CogoPoint;

                                        #region EAST
                                        if (dt2.Rows[i][col_EAST_PN] != DBNull.Value)
                                        {
                                            string pn2 = Convert.ToString(dt2.Rows[i][col_EAST_PN]);

                                            double slope_EAST = Convert.ToDouble(dt2.Rows[i][col_slope_EAST]);

                                            bool run1 = true;

                                            run1 = adjust_slope_towards_EAST(Trans1, dt2, ref pn2, ref slope_EAST, ref x1, ref y1, ref z1, pst1, pst2, plst1, plst2, layer_pt1, layer_pt2);

                                            if (run1 == true)
                                            {
                                                if (pn2 != "")
                                                {
                                                    do
                                                    {
                                                        run1 = adjust_slope_towards_EAST(Trans1, dt2, ref pn2, ref slope_EAST, ref x1, ref y1, ref z1, pst1, pst2, plst1, plst2, layer_pt1, layer_pt2);
                                                    } while (run1 == true);
                                                }

                                            }

                                        }
                                        #endregion
                                        x1 = Convert.ToDouble(dt2.Rows[i][col_x]);
                                        y1 = Convert.ToDouble(dt2.Rows[i][col_y]);
                                        z1 = Convert.ToDouble(dt2.Rows[i][col_z]);
                                        #region WEST
                                        if (dt2.Rows[i][col_WEST_PN] != DBNull.Value)
                                        {
                                            string pn2 = Convert.ToString(dt2.Rows[i][col_WEST_PN]);

                                            double slope_WEST = Convert.ToDouble(dt2.Rows[i][col_slope_WEST]);

                                            bool run1 = true;

                                            run1 = adjust_slope_towards_WEST(Trans1, dt2, ref pn2, ref slope_WEST, ref x1, ref y1, ref z1, pst1, pst2, plst1, plst2, layer_pt1, layer_pt2);

                                            if (run1 == true)
                                            {
                                                if (pn2 != "")
                                                {
                                                    do
                                                    {
                                                        run1 = adjust_slope_towards_WEST(Trans1, dt2, ref pn2, ref slope_WEST, ref x1, ref y1, ref z1, pst1, pst2, plst1, plst2, layer_pt1, layer_pt2);
                                                    } while (run1 == true);
                                                }

                                            }

                                        }
                                        #endregion
                                    }
                                }
                            }
                        }


                        if (radioButton_NS.Checked == true)
                        {




                            for (int i = 0; i < dt2.Rows.Count; ++i)
                            {
                                if (dt2.Rows[i][col_anchor] != DBNull.Value && Convert.ToBoolean(dt2.Rows[i][col_anchor]) == true)
                                {
                                    if (dt2.Rows[i][col_NS_processed] != DBNull.Value && Convert.ToBoolean(dt2.Rows[i][col_NS_processed]) == true)
                                    {
                                        double x1 = Convert.ToDouble(dt2.Rows[i][col_x]);
                                        double y1 = Convert.ToDouble(dt2.Rows[i][col_y]);
                                        double z1 = Convert.ToDouble(dt2.Rows[i][col_z]);
                                        string pn1 = Convert.ToString(dt2.Rows[i][col_pn]);


                                        ObjectId id1 = (ObjectId)dt2.Rows[i][col_id];
                                        CogoPoint cg_anchor = Trans1.GetObject(id1, OpenMode.ForWrite) as CogoPoint;
                                        #region NORTH
                                        if (dt2.Rows[i][col_NORTH_PN] != DBNull.Value)
                                        {
                                            string pn2 = Convert.ToString(dt2.Rows[i][col_NORTH_PN]);

                                            double slope_NORTH = Convert.ToDouble(dt2.Rows[i][col_slope_NORTH]);

                                            bool run1 = true;

                                            run1 = adjust_slope_towards_NORTH(Trans1, dt2, ref pn2, ref slope_NORTH, ref x1, ref y1, ref z1, pst1, pst2, plst1, plst2, layer_pt1, layer_pt2);

                                            if (run1 == true)
                                            {
                                                if (pn2 != "")
                                                {
                                                    do
                                                    {
                                                        run1 = adjust_slope_towards_NORTH(Trans1, dt2, ref pn2, ref slope_NORTH, ref x1, ref y1, ref z1, pst1, pst2, plst1, plst2, layer_pt1, layer_pt2);
                                                    } while (run1 == true);
                                                }

                                            }

                                        }
                                        #endregion

                                        x1 = Convert.ToDouble(dt2.Rows[i][col_x]);
                                        y1 = Convert.ToDouble(dt2.Rows[i][col_y]);
                                        z1 = Convert.ToDouble(dt2.Rows[i][col_z]);


                                        #region SOUTH
                                        if (dt2.Rows[i][col_SOUTH_PN] != DBNull.Value)
                                        {
                                            string pn2 = Convert.ToString(dt2.Rows[i][col_SOUTH_PN]);

                                            double slope_SOUTH = Convert.ToDouble(dt2.Rows[i][col_slope_SOUTH]);

                                            bool run1 = true;

                                            run1 = adjust_slope_towards_SOUTH(Trans1, dt2, ref pn2, ref slope_SOUTH, ref x1, ref y1, ref z1, pst1, pst2, plst1, plst2, layer_pt1, layer_pt2);

                                            if (run1 == true)
                                            {
                                                if (pn2 != "")
                                                {
                                                    do
                                                    {
                                                        run1 = adjust_slope_towards_SOUTH(Trans1, dt2, ref pn2, ref slope_SOUTH, ref x1, ref y1, ref z1, pst1, pst2, plst1, plst2, layer_pt1, layer_pt2);
                                                    } while (run1 == true);
                                                }

                                            }

                                        }
                                        #endregion

                                    }
                                }

                            }


                        }





                        #region calc new slopes

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
                        #endregion


                        #region final sort and  write to excel



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

                        
                        Microsoft.Office.Interop.Excel.Worksheet W1 = Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt2);
                        W1.Name = Environment.UserName + " " + DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + " at " + DateTime.Now.Hour + "h" + DateTime.Now.Minute + "m";

                        int hide1 = 0;
                        W1.Range["A1:X1"].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        W1.Range["A1:X1"].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                        W1.Range["A1:X1"].WrapText = true;
                        W1.Range["A:A"].ColumnWidth = 13;
                        W1.Range["B:E"].ColumnWidth = 11;
                        W1.Range["F:M"].ColumnWidth = 10;
                        //W1.Range["G:G"].ColumnWidth = hide1;
                        //W1.Range["H:H"].ColumnWidth = hide1;
                        //W1.Range["I:I"].ColumnWidth = hide1;
                        //W1.Range["J:J"].ColumnWidth = 10;
                        //W1.Range["K:K"].ColumnWidth = 10;
                        //W1.Range["L:M"].ColumnWidth = hide1;
                        W1.Range["N:Q"].ColumnWidth = hide1;
                        W1.Range["R:V"].ColumnWidth = 10;
                        W1.Range["W:X"].ColumnWidth = 12;


                        dt1.Dispose();
                        dt2.Dispose();
                        #endregion

                        Trans1.Commit();

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

        private bool adjust_slope_towards_SOUTH(Transaction Trans1, System.Data.DataTable dt2, ref string pn2, ref double slope_SOUTH, ref double x1, ref double y1, ref double z1,
            Autodesk.Civil.DatabaseServices.Styles.PointStyle pst1, Autodesk.Civil.DatabaseServices.Styles.PointStyle pst2,
            Autodesk.Civil.DatabaseServices.Styles.LabelStyle plst1, Autodesk.Civil.DatabaseServices.Styles.LabelStyle plst2,
            string layer_pt1, string layer_pt2)
        {


            double calc_slope = max_NS1;

            if (slope_SOUTH < 0) calc_slope = min_NS2;

            for (int j = 0; j < dt2.Rows.Count; ++j)
            {
                string pn3 = Convert.ToString(dt2.Rows[j][col_pn]);
                if (pn2 == pn3)
                {

                    double z2 = Convert.ToDouble(dt2.Rows[j][col_z]);
                    double x2 = Convert.ToDouble(dt2.Rows[j][col_x]);
                    double y2 = Convert.ToDouble(dt2.Rows[j][col_y]);

                    if (slope_SOUTH >= min_NS2 && slope_SOUTH <= max_NS1)
                    {

                        x1 = x2;
                        y1 = y2;
                        z1 = z2;

                        dt2.Rows[j][col_new_elev] = dt2.Rows[j][col_z];


                        if (dt2.Rows[j][col_SOUTH_PN] != DBNull.Value)
                        {
                            pn2 = Convert.ToString(dt2.Rows[j][col_SOUTH_PN]);
                            slope_SOUTH = Convert.ToDouble(dt2.Rows[j][col_slope_SOUTH]);
                            return true;
                        }

                    }
                    else
                    {
                        double Run = Math.Abs(y2 - y1);
                        double z3 = z1 - calc_slope * Run / 100;


                        dt2.Rows[j][col_newZ_NS] = z3;
                        dt2.Rows[j][col_new_elev] = z3;
                        dt2.Rows[j][col_max_NS] = calc_slope;
                        dt2.Rows[j][col_NS_processed] = true;


                        double z0 = Convert.ToDouble(dt2.Rows[j][col_z]);
                        if (dt2.Rows[j][col_orig_elev] != DBNull.Value)
                        {
                            z0 = Convert.ToDouble(dt2.Rows[j][col_orig_elev]);
                        }

                        if (dt2.Rows[j][col_new_elev] != DBNull.Value)
                        {
                            CogoPoint cg1 = null;
                            if (dt2.Rows[j][col_id] != DBNull.Value)
                            {
                                ObjectId id1 = (ObjectId)dt2.Rows[j][col_id];
                                cg1 = Trans1.GetObject(id1, OpenMode.ForWrite) as CogoPoint;
                                if (cg1 != null)
                                {
                                    cg1.Elevation = z3;
                                }

                            }

                            if (z3 > z0)
                            {
                                dt2.Rows[j][col_new_description] = "FILL";
                                if (cg1 != null)
                                {
                                    cg1.RawDescription = "FILL";
                                    cg1.Layer = layer_pt1;
                                    if (pst1 != null) cg1.StyleId = pst1.ObjectId;
                                    if (plst1 != null) cg1.LabelStyleId = plst1.ObjectId;
                                }
                            }
                            else
                            {
                                dt2.Rows[j][col_new_description] = "CUT";
                                if (cg1 != null)
                                {
                                    cg1.RawDescription = "CUT";
                                    cg1.Layer = layer_pt2;
                                    if (pst2 != null) cg1.StyleId = pst2.ObjectId;
                                    if (plst2 != null) cg1.LabelStyleId = plst2.ObjectId;
                                }

                            }
                        }


                        if (dt2.Rows[j][col_SOUTH_PN] != DBNull.Value)
                        {
                            pn2 = Convert.ToString(dt2.Rows[j][col_SOUTH_PN]);


                            for (int k = 0; k < dt2.Rows.Count; ++k)
                            {
                                string pn4 = Convert.ToString(dt2.Rows[k][col_pn]);
                                if (pn2 == pn4)
                                {
                                    double x4 = Convert.ToDouble(dt2.Rows[k][col_x]);
                                    double y4 = Convert.ToDouble(dt2.Rows[k][col_y]);


                                    double z4 = Convert.ToDouble(dt2.Rows[k][col_z]);
                                    double Run4 = Math.Abs(y2 - y4);
                                    double Rise = z3 - z4;
                                    slope_SOUTH = Math.Round(100 * Rise / Run4, 4);
                                    k = dt2.Rows.Count;

                                }

                            }
                            x1 = x2;
                            y1 = y2;
                            z1 = z3;
                            return true;
                        }

                        x1 = x2;
                        y1 = y2;
                        z1 = z3;

                    }

                }



            }



            pn2 = "";
            return false;
        }

        private bool adjust_slope_towards_NORTH(Transaction Trans1, System.Data.DataTable dt2, ref string pn2, ref double slope_NORTH, ref double x1, ref double y1, ref double z1,
            Autodesk.Civil.DatabaseServices.Styles.PointStyle pst1, Autodesk.Civil.DatabaseServices.Styles.PointStyle pst2,
            Autodesk.Civil.DatabaseServices.Styles.LabelStyle plst1, Autodesk.Civil.DatabaseServices.Styles.LabelStyle plst2,
            string layer_pt1, string layer_pt2)
        {




            double calc_slope = max_NS1;

            if (slope_NORTH < 0) calc_slope = min_NS2;

            for (int j = 0; j < dt2.Rows.Count; ++j)
            {
                string pn3 = Convert.ToString(dt2.Rows[j][col_pn]);
                if (pn2 == pn3)
                {

                    double z2 = Convert.ToDouble(dt2.Rows[j][col_z]);
                    double x2 = Convert.ToDouble(dt2.Rows[j][col_x]);
                    double y2 = Convert.ToDouble(dt2.Rows[j][col_y]);

                    if (slope_NORTH >= min_NS2 && slope_NORTH <= max_NS1)
                    {

                        x1 = x2;
                        y1 = y2;
                        z1 = z2;

                        dt2.Rows[j][col_new_elev] = dt2.Rows[j][col_z];


                        if (dt2.Rows[j][col_NORTH_PN] != DBNull.Value)
                        {
                            pn2 = Convert.ToString(dt2.Rows[j][col_NORTH_PN]);
                            slope_NORTH = Convert.ToDouble(dt2.Rows[j][col_slope_NORTH]);
                            return true;
                        }

                    }
                    else
                    {
                        double Run = Math.Abs(y2 - y1);
                        double z3 = z1 - calc_slope * Run / 100;


                        dt2.Rows[j][col_newZ_NS] = z3;
                        dt2.Rows[j][col_new_elev] = z3;
                        dt2.Rows[j][col_max_NS] = calc_slope;
                        dt2.Rows[j][col_NS_processed] = true;


                        double z0 = Convert.ToDouble(dt2.Rows[j][col_z]);
                        if (dt2.Rows[j][col_orig_elev] != DBNull.Value)
                        {
                            z0 = Convert.ToDouble(dt2.Rows[j][col_orig_elev]);
                        }

                        if (dt2.Rows[j][col_new_elev] != DBNull.Value)
                        {
                            CogoPoint cg1 = null;
                            if (dt2.Rows[j][col_id] != DBNull.Value)
                            {
                                ObjectId id1 = (ObjectId)dt2.Rows[j][col_id];
                                cg1 = Trans1.GetObject(id1, OpenMode.ForWrite) as CogoPoint;
                                if (cg1 != null)
                                {
                                    cg1.Elevation = z3;
                                }

                            }

                            if (z3 > z0)
                            {
                                dt2.Rows[j][col_new_description] = "FILL";
                                if (cg1 != null)
                                {
                                    cg1.RawDescription = "FILL";
                                    cg1.Layer = layer_pt1;
                                    if (pst1 != null) cg1.StyleId = pst1.ObjectId;
                                    if (plst1 != null) cg1.LabelStyleId = plst1.ObjectId;
                                }
                            }
                            else
                            {
                                dt2.Rows[j][col_new_description] = "CUT";
                                if (cg1 != null)
                                {
                                    cg1.RawDescription = "CUT";
                                    cg1.Layer = layer_pt2;
                                    if (pst2 != null) cg1.StyleId = pst2.ObjectId;
                                    if (plst2 != null) cg1.LabelStyleId = plst2.ObjectId;
                                }

                            }
                        }


                        if (dt2.Rows[j][col_NORTH_PN] != DBNull.Value)
                        {
                            pn2 = Convert.ToString(dt2.Rows[j][col_NORTH_PN]);


                            for (int k = 0; k < dt2.Rows.Count; ++k)
                            {
                                string pn4 = Convert.ToString(dt2.Rows[k][col_pn]);
                                if (pn2 == pn4)
                                {
                                    double x4 = Convert.ToDouble(dt2.Rows[k][col_x]);
                                    double y4 = Convert.ToDouble(dt2.Rows[k][col_y]);


                                    double z4 = Convert.ToDouble(dt2.Rows[k][col_z]);
                                    double Run4 = Math.Abs(y2 - y4);
                                    double Rise = z3 - z4;
                                    slope_NORTH = Math.Round(100 * Rise / Run4, 4);
                                    k = dt2.Rows.Count;

                                }

                            }
                            x1 = x2;
                            y1 = y2;
                            z1 = z3;
                            return true;
                        }

                        x1 = x2;
                        y1 = y2;
                        z1 = z3;

                    }

                }



            }



            pn2 = "";
            return false;
        }


        private bool adjust_slope_towards_WEST(Transaction Trans1, System.Data.DataTable dt2, ref string pn2, ref double slope_WEST, ref double x1, ref double y1, ref double z1,
            Autodesk.Civil.DatabaseServices.Styles.PointStyle pst1, Autodesk.Civil.DatabaseServices.Styles.PointStyle pst2,
            Autodesk.Civil.DatabaseServices.Styles.LabelStyle plst1, Autodesk.Civil.DatabaseServices.Styles.LabelStyle plst2,
            string layer_pt1, string layer_pt2)
        {




            double calc_slope = max_EW1;

            if (slope_WEST < 0) calc_slope = min_EW2;

            for (int j = 0; j < dt2.Rows.Count; ++j)
            {
                string pn3 = Convert.ToString(dt2.Rows[j][col_pn]);
                if (pn2 == pn3)
                {

                    double z2 = Convert.ToDouble(dt2.Rows[j][col_z]);
                    double x2 = Convert.ToDouble(dt2.Rows[j][col_x]);
                    double y2 = Convert.ToDouble(dt2.Rows[j][col_y]);

                    if (slope_WEST >= min_EW2 && slope_WEST <= max_EW1)
                    {

                        x1 = x2;
                        y1 = y2;
                        z1 = z2;

                        dt2.Rows[j][col_new_elev] = dt2.Rows[j][col_z];

                        slope_WEST = 0;
                        if (dt2.Rows[j][col_WEST_PN] != DBNull.Value)
                        {
                            pn2 = Convert.ToString(dt2.Rows[j][col_WEST_PN]);
                            slope_WEST = Convert.ToDouble(dt2.Rows[j][col_slope_WEST]);
                            return true;
                        }

                    }
                    else
                    {
                        double run = Math.Abs(x2 - x1);
                        double z3 = z1 - calc_slope * run / 100;


                        dt2.Rows[j][col_newZ_EW] = z3;
                        dt2.Rows[j][col_new_elev] = z3;
                        dt2.Rows[j][col_max_EW] = calc_slope;
                        dt2.Rows[j][col_EW_processed] = true;


                        double z0 = Convert.ToDouble(dt2.Rows[j][col_z]);
                        if (dt2.Rows[j][col_orig_elev] != DBNull.Value)
                        {
                            z0 = Convert.ToDouble(dt2.Rows[j][col_orig_elev]);
                        }

                        if (dt2.Rows[j][col_new_elev] != DBNull.Value)
                        {
                            CogoPoint cg1 = null;
                            if (dt2.Rows[j][col_id] != DBNull.Value)
                            {
                                ObjectId id1 = (ObjectId)dt2.Rows[j][col_id];
                                cg1 = Trans1.GetObject(id1, OpenMode.ForWrite) as CogoPoint;
                                if (cg1 != null)
                                {
                                    cg1.Elevation = z3;
                                }

                            }

                            if (z3 > z0)
                            {
                                dt2.Rows[j][col_new_description] = "FILL";
                                if (cg1 != null)
                                {
                                    cg1.RawDescription = "FILL";
                                    cg1.Layer = layer_pt1;
                                    if (pst1 != null) cg1.StyleId = pst1.ObjectId;
                                    if (plst1 != null) cg1.LabelStyleId = plst1.ObjectId;
                                }
                            }
                            else
                            {
                                dt2.Rows[j][col_new_description] = "CUT";
                                if (cg1 != null)
                                {
                                    cg1.RawDescription = "CUT";
                                    cg1.Layer = layer_pt2;
                                    if (pst2 != null) cg1.StyleId = pst2.ObjectId;
                                    if (plst2 != null) cg1.LabelStyleId = plst2.ObjectId;
                                }

                            }
                        }


                        if (dt2.Rows[j][col_WEST_PN] != DBNull.Value)
                        {
                            pn2 = Convert.ToString(dt2.Rows[j][col_WEST_PN]);


                            for (int k = 0; k < dt2.Rows.Count; ++k)
                            {
                                string pn4 = Convert.ToString(dt2.Rows[k][col_pn]);
                                if (pn2 == pn4)
                                {
                                    double x4 = Convert.ToDouble(dt2.Rows[k][col_x]);
                                    double y4 = Convert.ToDouble(dt2.Rows[k][col_y]);


                                    double z4 = Convert.ToDouble(dt2.Rows[k][col_z]);
                                    double Run = Math.Abs(x2 - x4);
                                    double Rise = z3 - z4;
                                    slope_WEST = Math.Round(100 * Rise / Run, 4);
                                    k = dt2.Rows.Count;

                                }

                            }
                            x1 = x2;
                            y1 = y2;
                            z1 = z3;
                            return true;
                        }

                        x1 = x2;
                        y1 = y2;
                        z1 = z3;

                    }

                }



            }



            pn2 = "";
            return false;
        }


        private bool adjust_slope_towards_EAST(Transaction Trans1, System.Data.DataTable dt2, ref string pn2, ref double slope_EAST, ref double x1, ref double y1, ref double z1,
                                                                            Autodesk.Civil.DatabaseServices.Styles.PointStyle pst1, Autodesk.Civil.DatabaseServices.Styles.PointStyle pst2,
                                                                                Autodesk.Civil.DatabaseServices.Styles.LabelStyle plst1, Autodesk.Civil.DatabaseServices.Styles.LabelStyle plst2,
                                                                                                                                                                        string layer_pt1, string layer_pt2)
        {




            double calc_slope = max_EW1;

            if (slope_EAST < 0) calc_slope = min_EW2;

            for (int j = 0; j < dt2.Rows.Count; ++j)
            {
                string pn3 = Convert.ToString(dt2.Rows[j][col_pn]);
                if (pn2 == pn3)
                {

                    double z2 = Convert.ToDouble(dt2.Rows[j][col_z]);
                    double x2 = Convert.ToDouble(dt2.Rows[j][col_x]);
                    double y2 = Convert.ToDouble(dt2.Rows[j][col_y]);

                    if (slope_EAST >= min_EW2 && slope_EAST <= max_EW1)
                    {

                        x1 = x2;
                        y1 = y2;
                        z1 = z2;

                        dt2.Rows[j][col_new_elev] = dt2.Rows[j][col_z];

                        slope_EAST = 0;
                        if (dt2.Rows[j][col_EAST_PN] != DBNull.Value)
                        {
                            pn2 = Convert.ToString(dt2.Rows[j][col_EAST_PN]);
                            slope_EAST = Convert.ToDouble(dt2.Rows[j][col_slope_EAST]);
                            return true;
                        }

                    }
                    else
                    {
                        double run = Math.Abs(x2 - x1);
                        double z3 = z1 - calc_slope * run / 100;


                        dt2.Rows[j][col_newZ_EW] = z3;
                        dt2.Rows[j][col_new_elev] = z3;
                        dt2.Rows[j][col_max_EW] = calc_slope;
                        dt2.Rows[j][col_EW_processed] = true;


                        double z0 = Convert.ToDouble(dt2.Rows[j][col_z]);
                        if (dt2.Rows[j][col_orig_elev] != DBNull.Value)
                        {
                            z0 = Convert.ToDouble(dt2.Rows[j][col_orig_elev]);
                        }

                        if (dt2.Rows[j][col_new_elev] != DBNull.Value)
                        {
                            CogoPoint cg1 = null;
                            if (dt2.Rows[j][col_id] != DBNull.Value)
                            {
                                ObjectId id1 = (ObjectId)dt2.Rows[j][col_id];
                                cg1 = Trans1.GetObject(id1, OpenMode.ForWrite) as CogoPoint;
                                if (cg1 != null)
                                {
                                    cg1.Elevation = z3;
                                }

                            }

                            if (z3 > z0)
                            {
                                dt2.Rows[j][col_new_description] = "FILL";
                                if (cg1 != null)
                                {
                                    cg1.RawDescription = "FILL";
                                    cg1.Layer = layer_pt1;
                                    if (pst1 != null) cg1.StyleId = pst1.ObjectId;
                                    if (plst1 != null) cg1.LabelStyleId = plst1.ObjectId;
                                }
                            }
                            else
                            {
                                dt2.Rows[j][col_new_description] = "CUT";
                                if (cg1 != null)
                                {
                                    cg1.RawDescription = "CUT";
                                    cg1.Layer = layer_pt2;
                                    if (pst2 != null) cg1.StyleId = pst2.ObjectId;
                                    if (plst2 != null) cg1.LabelStyleId = plst2.ObjectId;
                                }

                            }
                        }


                        if (dt2.Rows[j][col_EAST_PN] != DBNull.Value)
                        {
                            pn2 = Convert.ToString(dt2.Rows[j][col_EAST_PN]);


                            for (int k = 0; k < dt2.Rows.Count; ++k)
                            {
                                string pn4 = Convert.ToString(dt2.Rows[k][col_pn]);
                                if (pn2 == pn4)
                                {
                                    double x4 = Convert.ToDouble(dt2.Rows[k][col_x]);
                                    double y4 = Convert.ToDouble(dt2.Rows[k][col_y]);


                                    double z4 = Convert.ToDouble(dt2.Rows[k][col_z]);
                                    double Run = Math.Abs(x2 - x4);
                                    double Rise = z3 - z4;
                                    slope_EAST = Math.Round(100 * Rise / Run, 4);
                                    k = dt2.Rows.Count;

                                }

                            }
                            x1 = x2;
                            y1 = y2;
                            z1 = z3;
                            return true;
                        }

                        x1 = x2;
                        y1 = y2;
                        z1 = z3;
                    }
                }
            }
            pn2 = "";
            return false;
        }

        private void button_load_point_style1_Click(object sender, EventArgs e)
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

        private void button_load_point_style3_Click(object sender, EventArgs e)
        {
            comboBox_point_style3.Items.Clear();
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
                            comboBox_point_style3.Items.Add(pst1.Name);
                        }
                        if (comboBox_point_style3.Items.Contains("Anchor") == true)
                        {
                            comboBox_point_style3.SelectedIndex = comboBox_point_style3.Items.IndexOf("Anchor");
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

        private void button_load_point_style4_Click(object sender, EventArgs e)
        {
            comboBox_point_style4.Items.Clear();
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
                            comboBox_point_style4.Items.Add(pst1.Name);
                        }
                        if (comboBox_point_style4.Items.Contains("EG") == true)
                        {
                            comboBox_point_style4.SelectedIndex = comboBox_point_style4.Items.IndexOf("EG");
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

        private void button_load_point_style5_Click(object sender, EventArgs e)
        {
            comboBox_point_style5.Items.Clear();
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
                            comboBox_point_style5.Items.Add(pst1.Name);
                        }
                        if (comboBox_point_style5.Items.Contains("Daylight") == true)
                        {
                            comboBox_point_style5.SelectedIndex = comboBox_point_style5.Items.IndexOf("Daylight");
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

        private void button_load_point_label_style1_Click(object sender, EventArgs e)
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

        private void button_load_point_label_style3_Click(object sender, EventArgs e)
        {
            comboBox_point_label_style3.Items.Clear();

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
                            comboBox_point_label_style3.Items.Add(plst2.Name);
                        }

                        if (comboBox_point_label_style3.Items.Contains("Point#-Elevation-Description [ANCHOR]") == true)
                        {
                            comboBox_point_label_style3.SelectedIndex = comboBox_point_label_style3.Items.IndexOf("Point#-Elevation-Description [ANCHOR]");
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

        private void button_load_point_label_style4_Click(object sender, EventArgs e)
        {
            comboBox_point_label_style4.Items.Clear();

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
                            comboBox_point_label_style4.Items.Add(plst2.Name);
                        }

                        if (comboBox_point_label_style4.Items.Contains("Point#-Elevation-Description [EG]") == true)
                        {
                            comboBox_point_label_style4.SelectedIndex = comboBox_point_label_style4.Items.IndexOf("Point#-Elevation-Description [EG]");
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

        private void button_load_point_label_style5_Click(object sender, EventArgs e)
        {
            comboBox_point_label_style5.Items.Clear();

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
                            comboBox_point_label_style5.Items.Add(plst2.Name);
                        }

                        if (comboBox_point_label_style5.Items.Contains("Point#-Elevation-Description [Daylight]") == true)
                        {
                            comboBox_point_label_style5.SelectedIndex = comboBox_point_label_style5.Items.IndexOf("Point#-Elevation-Description [Daylight]");
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

        private void button_load_layers_to_combobox1_Click(object sender, EventArgs e)
        {
            Functions.Incarca_existing_layers_to_combobox(comboBox_point_layer1);

            if (comboBox_point_layer1.Items.Contains("PTS-Fill") == true)
            {
                comboBox_point_layer1.SelectedIndex = comboBox_point_layer1.Items.IndexOf("PTS-Fill");
            }

        }

        private void button_load_layers_to_combobox2_Click(object sender, EventArgs e)
        {
            Functions.Incarca_existing_layers_to_combobox(comboBox_point_layer2);

            if (comboBox_point_layer2.Items.Contains("PTS-Cut") == true)
            {
                comboBox_point_layer2.SelectedIndex = comboBox_point_layer2.Items.IndexOf("PTS-Cut");
            }

        }

        private void button_load_layers_to_combobox3_Click(object sender, EventArgs e)
        {
            Functions.Incarca_existing_layers_to_combobox(comboBox_point_layer3);

            if (comboBox_point_layer3.Items.Contains("PTS-Anchor") == true)
            {
                comboBox_point_layer3.SelectedIndex = comboBox_point_layer3.Items.IndexOf("PTS-Anchor");
            }

        }

        private void button_load_layers_to_combobox4_Click(object sender, EventArgs e)
        {
            Functions.Incarca_existing_layers_to_combobox(comboBox_point_layer4);

            if (comboBox_point_layer4.Items.Contains("PTS-EG") == true)
            {
                comboBox_point_layer4.SelectedIndex = comboBox_point_layer4.Items.IndexOf("PTS-EG");
            }

        }

        private void button_load_layers_to_combobox5_Click(object sender, EventArgs e)
        {
            Functions.Incarca_existing_layers_to_combobox(comboBox_point_layer5);

            if (comboBox_point_layer5.Items.Contains("PTS-Daylight") == true)
            {
                comboBox_point_layer5.SelectedIndex = comboBox_point_layer5.Items.IndexOf("PTS-Daylight");
            }

        }








        private void button_set_anchor_points_Click(object sender, EventArgs e)
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

                        Polyline poly_anchor = null;

                        bool delete_anchor = false;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat2;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez2 = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez2.MessageForAdding = "\nSelect anchor polyline";
                        Prompt_rez2.SingleOnly = true;

                        Rezultat2 = ThisDrawing.Editor.GetSelection(Prompt_rez2);

                        if (Rezultat2.Status != PromptStatus.OK)
                        {
                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                            Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                            PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify start point:");
                            PP1.AllowNone = false;
                            Point_res1 = Editor1.GetPoint(PP1);

                            if (Point_res1.Status != PromptStatus.OK)
                            {
                                set_enable_true();
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                this.MdiParent.WindowState = FormWindowState.Normal;
                                return;

                            }

                            List<Point3d> lista1 = new List<Point3d>();
                            lista1.Add(Point_res1.Value);
                            bool run1 = true;
                            do
                            {
                                Alignment_mdi.Jig_draw_polyline Jig1 = new Alignment_mdi.Jig_draw_polyline();
                                Autodesk.AutoCAD.EditorInput.PromptPointResult Rez_point_2 = Jig1.StartJig(lista1, "Specify next point:");

                                if (Rez_point_2.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                {
                                    lista1.Add(Rez_point_2.Value);

                                }
                                else
                                {
                                    run1 = false;

                                }

                            } while (run1 == true);

                            poly_anchor = new Polyline();
                            for (int i = 0; i < lista1.Count; i++)
                            {
                                poly_anchor.AddVertexAt(i, new Point2d(lista1[i].X, lista1[i].Y), 0, 0, 0);
                            }

                            BTrecord.AppendEntity(poly_anchor);
                            Trans1.AddNewlyCreatedDBObject(poly_anchor, true);
                            delete_anchor = true;

                        }
                        else
                        {
                            poly_anchor = Trans1.GetObject(Rezultat2.Value[0].ObjectId, OpenMode.ForRead) as Polyline;

                        }

                        if (poly_anchor == null)
                        {
                            set_enable_true();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            this.MdiParent.WindowState = FormWindowState.Normal;
                            return;
                        }

                        Trans1.TransactionManager.QueueForGraphicsFlush();
                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez1 = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez1.MessageForAdding = "\nSelect points";
                        Prompt_rez1.SingleOnly = false;

                        Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez1);

                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            set_enable_true();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            this.MdiParent.WindowState = FormWindowState.Normal;
                            return;
                        }


                        #region CIVIL STYLES AND LAYERS
                        Autodesk.Civil.DatabaseServices.Styles.PointStyleCollection col_point_styles = CivilDrawing.Styles.PointStyles;
                        IEnumerator<ObjectId> enum1 = col_point_styles.GetEnumerator();

                        Autodesk.Civil.DatabaseServices.Styles.PointStyle pst1 = null;
                        Autodesk.Civil.DatabaseServices.Styles.PointStyle pst2 = null;
                        Autodesk.Civil.DatabaseServices.Styles.PointStyle pst3 = null;
                        Autodesk.Civil.DatabaseServices.Styles.PointStyle pst4 = null;
                        Autodesk.Civil.DatabaseServices.Styles.PointStyle pst5 = null;

                        while (enum1.MoveNext())
                        {
                            Autodesk.Civil.DatabaseServices.Styles.PointStyle pst11 = Trans1.GetObject(enum1.Current, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.PointStyle;
                            if (pst11.Name == comboBox_point_style1.Text)
                            {
                                pst1 = pst11;
                            }
                            if (pst11.Name == comboBox_point_style2.Text)
                            {
                                pst2 = pst11;
                            }
                            if (pst11.Name == comboBox_point_style3.Text)
                            {
                                pst3 = pst11;
                            }
                            if (pst11.Name == comboBox_point_style4.Text)
                            {
                                pst4 = pst11;
                            }
                            if (pst11.Name == comboBox_point_style5.Text)
                            {
                                pst5 = pst11;
                            }
                        }

                        Autodesk.Civil.DatabaseServices.Styles.LabelStyleCollection col_point_label_styles = CivilDrawing.Styles.LabelStyles.PointLabelStyles.LabelStyles;

                        Autodesk.Civil.DatabaseServices.Styles.LabelStyle plst1 = null;
                        Autodesk.Civil.DatabaseServices.Styles.LabelStyle plst2 = null;
                        Autodesk.Civil.DatabaseServices.Styles.LabelStyle plst3 = null;
                        Autodesk.Civil.DatabaseServices.Styles.LabelStyle plst4 = null;
                        Autodesk.Civil.DatabaseServices.Styles.LabelStyle plst5 = null;

                        IEnumerator<ObjectId> enum11 = col_point_label_styles.GetEnumerator();

                        while (enum11.MoveNext())
                        {
                            Autodesk.Civil.DatabaseServices.Styles.LabelStyle plst11 = Trans1.GetObject(enum11.Current, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.LabelStyle;


                            if (plst11.Name == comboBox_point_label_style1.Text)
                            {
                                plst1 = plst11;
                            }

                            if (plst11.Name == comboBox_point_label_style2.Text)
                            {
                                plst2 = plst11;
                            }
                            if (plst11.Name == comboBox_point_label_style3.Text)
                            {
                                plst3 = plst11;
                            }

                            if (plst11.Name == comboBox_point_label_style4.Text)
                            {
                                plst4 = plst11;
                            }
                            if (plst11.Name == comboBox_point_label_style5.Text)
                            {
                                plst5 = plst11;
                            }
                        }

                        string layer_pt1 = "0";
                        if (comboBox_point_layer1.Text.Length > 0)
                        {
                            layer_pt1 = comboBox_point_layer1.Text;

                        }


                        string layer_pt2 = "0";
                        if (comboBox_point_layer2.Text.Length > 0)
                        {
                            layer_pt2 = comboBox_point_layer2.Text;

                        }

                        string layer_pt3 = "0";
                        if (comboBox_point_layer3.Text.Length > 0)
                        {
                            layer_pt3 = comboBox_point_layer3.Text;

                        }


                        string layer_pt4 = "0";
                        if (comboBox_point_layer4.Text.Length > 0)
                        {
                            layer_pt4 = comboBox_point_layer4.Text;

                        }

                        string layer_pt5 = "0";
                        if (comboBox_point_layer5.Text.Length > 0)
                        {
                            layer_pt5 = comboBox_point_layer5.Text;

                        }


                        UDPString udp1 = Functions.Find_udp_string(comboBox_udp_field_1.Text);


                        UDPDouble udp2 = Functions.Find_udp_double(comboBox_udp_field_2.Text);


                        #endregion


                        System.Data.DataTable dt1 = new System.Data.DataTable();
                        dt1.Columns.Add(col_pn, typeof(string));
                        dt1.Columns.Add(col_y, typeof(double));
                        dt1.Columns.Add(col_x, typeof(double));

                        dt1.Columns.Add(col_descr, typeof(string));
                        dt1.Columns.Add(col_id, typeof(ObjectId));
                        dt1.Columns.Add(col_col, typeof(int));
                        dt1.Columns.Add(col_row, typeof(int));
                        dt1.Columns.Add(col_dist, typeof(double));
                        dt1.Columns.Add(col_anchor, typeof(bool));


                        col_sel_ids = new ObjectIdCollection();


                        double ymin = -1.234;
                        double ymax = -1.234;

                        double xmin = -1.234;
                        double xmax = -1.234;


                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {
                            CogoPoint cg1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForWrite) as CogoPoint;
                            if (cg1 != null)
                            {

                                col_sel_ids.Add(cg1.ObjectId);

                                dt1.Rows.Add();
                                dt1.Rows[dt1.Rows.Count - 1][col_pn] = cg1.PointNumber;
                                dt1.Rows[dt1.Rows.Count - 1][col_id] = cg1.ObjectId;

                                double x1 = Math.Round(cg1.Easting, 4);
                                double y1 = Math.Round(cg1.Northing, 4);


                                Point3d pt_on_poly = poly_anchor.GetClosestPointTo(new Point3d(x1, y1, poly_anchor.Elevation), Vector3d.ZAxis, false);
                                double x2 = pt_on_poly.X;
                                double y2 = pt_on_poly.Y;

                                double d1 = Math.Pow(Math.Pow(x1 - x2, 2) + Math.Pow(y1 - y2, 2), 0.5);
                                dt1.Rows[dt1.Rows.Count - 1][col_dist] = d1;

                                dt1.Rows[dt1.Rows.Count - 1][col_x] = x1;
                                dt1.Rows[dt1.Rows.Count - 1][col_y] = y1;


                                dt1.Rows[dt1.Rows.Count - 1][col_descr] = cg1.RawDescription;
                                dt1.Rows[dt1.Rows.Count - 1][col_anchor] = false;

                                double orig_elev = Math.Round(cg1.GetUDPValue(udp2), 4);
                                double z1 = Math.Round(cg1.Elevation, 4);

                                if (z1 > orig_elev)
                                {
                                    cg1.Layer = layer_pt1;
                                    if (pst1 != null) cg1.StyleId = pst1.ObjectId;
                                    if (plst1 != null) cg1.LabelStyleId = plst1.ObjectId;
                                    cg1.RawDescription = "FILL";
                                }
                                else if (z1 < orig_elev)
                                {
                                    cg1.Layer = layer_pt2;
                                    if (pst2 != null) cg1.StyleId = pst2.ObjectId;
                                    if (plst2 != null) cg1.LabelStyleId = plst2.ObjectId;
                                    cg1.RawDescription = "CUT";
                                }
                                else
                                {
                                    cg1.Layer = layer_pt4;
                                    if (pst4 != null) cg1.StyleId = pst4.ObjectId;
                                    if (plst4 != null) cg1.LabelStyleId = plst4.ObjectId;
                                    cg1.RawDescription = "EG";
                                }


                                if (xmin == -1.234)
                                {
                                    xmin = x1;
                                }
                                if (xmax == -1.234)
                                {
                                    xmax = x1;
                                }
                                if (ymin == -1.234)
                                {
                                    ymin = y1;
                                }
                                if (ymax == -1.234)
                                {
                                    ymax = y1;
                                }


                                if (x1 < xmin)
                                {
                                    xmin = x1;
                                }
                                if (x1 > xmax)
                                {
                                    xmax = x1;
                                }
                                if (y1 < ymin)
                                {
                                    ymin = y1;
                                }
                                if (y1 > ymax)
                                {
                                    ymax = y1;
                                }

                            }
                        }


                        using (System.Data.DataTable Data_table_temp = dt1.Clone())
                        {
                            DataView dv = new DataView(dt1);
                            dv.Sort = col_x + "," + col_y;
                            for (int i = 0; i < dv.Count; ++i)
                            {
                                System.Data.DataRow Data_row1 = dv[i].Row;
                                Data_table_temp.Rows.Add();
                                for (int j = 0; j < dt1.Columns.Count; ++j)
                                {
                                    Data_table_temp.Rows[Data_table_temp.Rows.Count - 1][j] = Data_row1[j];
                                }
                            }

                            dt1 = Data_table_temp.Copy();
                        }


                        int no_rows = Convert.ToInt32(1 + (ymax - ymin) / gridV);
                        int no_cols = Convert.ToInt32(1 + (xmax - xmin) / gridH);




                        for (int i = 0; i < dt1.Rows.Count; ++i)
                        {
                            double x1 = Convert.ToDouble(dt1.Rows[i][col_x]);
                            double y1 = Convert.ToDouble(dt1.Rows[i][col_y]);

                            int col1 = no_cols - Convert.ToInt32((xmax - x1) / gridH);
                            int row1 = no_rows - Convert.ToInt32((ymax - y1) / gridV);
                            dt1.Rows[i][col_col] = col1;
                            dt1.Rows[i][col_row] = row1;
                        }



                        System.Data.DataTable[] dt_array = null;


                        if (radioButton_EW.Checked == true)
                        {
                            Array.Resize(ref dt_array, no_rows);
                        }

                        if (radioButton_NS.Checked == true)
                        {
                            Array.Resize(ref dt_array, no_cols);
                        }



                        for (int i = 0; i < dt1.Rows.Count; ++i)
                        {
                            int col1 = Convert.ToInt32(dt1.Rows[i][col_col]);
                            int row1 = Convert.ToInt32(dt1.Rows[i][col_row]);

                            if (radioButton_EW.Checked == true)
                            {
                                if (dt_array[row1 - 1] == null)
                                {
                                    dt_array[row1 - 1] = dt1.Clone();
                                }

                                dt_array[row1 - 1].Rows.Add();
                                dt_array[row1 - 1].Rows[dt_array[row1 - 1].Rows.Count - 1].ItemArray = dt1.Rows[i].ItemArray;
                            }

                            if (radioButton_NS.Checked == true)
                            {
                                if (dt_array[col1 - 1] == null)
                                {
                                    dt_array[col1 - 1] = dt1.Clone();
                                }

                                dt_array[col1 - 1].Rows.Add();
                                dt_array[col1 - 1].Rows[dt_array[col1 - 1].Rows.Count - 1].ItemArray = dt1.Rows[i].ItemArray;
                            }

                        }

                        System.Data.DataTable dt2 = dt1.Clone();

                        for (int i = 0; i < dt_array.Length; i++)
                        {
                            System.Data.DataTable dt4 = Functions.Sort_data_table(dt_array[i], col_dist);
                            dt4.Rows[0][col_anchor] = true;

                            for (int j = 0; j < dt4.Rows.Count; j++)
                            {
                                dt2.Rows.Add();
                                dt2.Rows[dt2.Rows.Count - 1].ItemArray = dt4.Rows[j].ItemArray;
                            }

                        }














                        for (int i = 0; i < dt2.Rows.Count; ++i)
                        {
                            if (dt2.Rows[i][col_anchor] != DBNull.Value && dt2.Rows[i][col_id] != DBNull.Value)
                            {
                                if (Convert.ToBoolean(dt2.Rows[i][col_anchor]) == true)
                                {
                                    ObjectId id1 = (ObjectId)dt2.Rows[i][col_id];

                                    CogoPoint cg1 = Trans1.GetObject(id1, OpenMode.ForWrite) as CogoPoint;
                                    if (cg1 != null)
                                    {
                                        cg1.Layer = layer_pt3;

                                        if (pst3 != null) cg1.StyleId = pst3.ObjectId;
                                        if (plst3 != null) cg1.LabelStyleId = plst3.ObjectId;
                                        cg1.RawDescription = "ANCHOR";
                                    }


                                }
                            }
                        }












                        string nume1 = Environment.UserName + " " + DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + " at " + DateTime.Now.Hour + "h" + DateTime.Now.Minute + "m";






                        dt1.Dispose();
                        dt2.Dispose();


                        if (delete_anchor == true)
                        {
                            poly_anchor.UpgradeOpen();
                            poly_anchor.Erase();
                        }

                        Trans1.Commit();






                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                col_sel_ids = null;
            }
            set_enable_true();
            this.MdiParent.WindowState = FormWindowState.Normal;
            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
        }

        private void button_udp1_Click(object sender, EventArgs e)
        {
            comboBox_udp_field_1.Items.Clear();
            foreach (UDP udp1 in Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument.PointUDPs)
            {
             if(udp1!=null)   comboBox_udp_field_1.Items.Add(udp1.Name);
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
                if (udp2 != null) comboBox_udp_field_2.Items.Add(udp2.Name);
            }
            if (comboBox_udp_field_2.Items.Contains("ATT2") == true)
            {
                comboBox_udp_field_2.SelectedIndex = comboBox_udp_field_2.Items.IndexOf("ATT2");
            }
        }

        private void button_reset_to_dt_analise_Click(object sender, EventArgs e)
        {



            if (dt_analize != null && dt_analize.Rows.Count > 0)
            {

                ObjectId[] Empty_array = null;
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.Civil.ApplicationServices.CivilDocument CivilDrawing = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument;

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
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                            this.MdiParent.WindowState = FormWindowState.Minimized;

                            Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_cogopts = null;
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez1 = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rez1.MessageForAdding = "\nSelect points";
                            Prompt_rez1.SingleOnly = false;

                            Rezultat_cogopts = ThisDrawing.Editor.GetSelection(Prompt_rez1);
                            if (Rezultat_cogopts.Status != PromptStatus.OK)
                            {
                                set_enable_true();
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                this.MdiParent.WindowState = FormWindowState.Normal;
                                return;
                            }


                            #region CIVIL STYLES AND LAYERS
                            Autodesk.Civil.DatabaseServices.Styles.PointStyleCollection col_point_styles = CivilDrawing.Styles.PointStyles;
                            IEnumerator<ObjectId> enum1 = col_point_styles.GetEnumerator();

                            Autodesk.Civil.DatabaseServices.Styles.PointStyle pst1 = null;
                            Autodesk.Civil.DatabaseServices.Styles.PointStyle pst2 = null;
                            Autodesk.Civil.DatabaseServices.Styles.PointStyle pst4 = null;

                            while (enum1.MoveNext())
                            {
                                Autodesk.Civil.DatabaseServices.Styles.PointStyle pst11 = Trans1.GetObject(enum1.Current, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.PointStyle;
                                if (pst11.Name == comboBox_point_style1.Text)
                                {
                                    pst1 = pst11;
                                }
                                if (pst11.Name == comboBox_point_style2.Text)
                                {
                                    pst2 = pst11;
                                }
                                if (pst11.Name == comboBox_point_style4.Text)
                                {
                                    pst4 = pst11;
                                }
                            }

                            Autodesk.Civil.DatabaseServices.Styles.LabelStyleCollection col_point_label_styles = CivilDrawing.Styles.LabelStyles.PointLabelStyles.LabelStyles;

                            Autodesk.Civil.DatabaseServices.Styles.LabelStyle plst1 = null;
                            Autodesk.Civil.DatabaseServices.Styles.LabelStyle plst2 = null;
                            Autodesk.Civil.DatabaseServices.Styles.LabelStyle plst4 = null;

                            IEnumerator<ObjectId> enum11 = col_point_label_styles.GetEnumerator();

                            while (enum11.MoveNext())
                            {
                                Autodesk.Civil.DatabaseServices.Styles.LabelStyle plst11 = Trans1.GetObject(enum11.Current, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.LabelStyle;


                                if (plst11.Name == comboBox_point_label_style1.Text)
                                {
                                    plst1 = plst11;
                                }

                                if (plst11.Name == comboBox_point_label_style2.Text)
                                {
                                    plst2 = plst11;
                                }

                                if (plst11.Name == comboBox_point_label_style4.Text)
                                {
                                    plst4 = plst11;
                                }
                            }

                            string layer_pt1 = "0";
                            if (comboBox_point_layer1.Text.Length > 0)
                            {
                                layer_pt1 = comboBox_point_layer1.Text;

                            }


                            string layer_pt2 = "0";
                            if (comboBox_point_layer2.Text.Length > 0)
                            {
                                layer_pt2 = comboBox_point_layer2.Text;

                            }

                            string layer_pt4 = "0";
                            if (comboBox_point_layer4.Text.Length > 0)
                            {
                                layer_pt4 = comboBox_point_layer4.Text;

                            }

                            #endregion



                            for (int i = 0; i < Rezultat_cogopts.Value.Count; ++i)
                            {
                                ObjectId id1 = Rezultat_cogopts.Value[i].ObjectId;
                                CogoPoint cg1 = Trans1.GetObject(id1, OpenMode.ForRead) as CogoPoint;
                                if (cg1 != null)
                                {
                                    string pn1 = Convert.ToString(cg1.PointNumber);
                                    for (int j = 0; j < dt_analize.Rows.Count; ++j)
                                    {
                                        if (dt_analize.Rows[j][col_pn] != DBNull.Value && dt_analize.Rows[j][col_z] != DBNull.Value && dt_analize.Rows[j][col_descr] != DBNull.Value)
                                        {
                                            string pn2 = Convert.ToString(dt_analize.Rows[j][col_pn]);
                                            if (pn1 == pn2)
                                            {
                                                cg1.UpgradeOpen();
                                                cg1.Elevation = Convert.ToDouble(dt_analize.Rows[j][col_new_elev]);
                                                string descr1 = Convert.ToString(dt_analize.Rows[j][col_descr]);
                                                cg1.RawDescription = descr1;

                                                if (descr1 == "FILL")
                                                {
                                                    if (pst1 != null) cg1.StyleId = pst1.ObjectId;
                                                    if (plst1 != null) cg1.LabelStyleId = plst1.ObjectId;
                                                    cg1.Layer = layer_pt1;
                                                }
                                                else if (descr1 == "CUT")
                                                {
                                                    if (pst2 != null) cg1.StyleId = pst2.ObjectId;
                                                    if (plst2 != null) cg1.LabelStyleId = plst2.ObjectId;
                                                    cg1.Layer = layer_pt2;
                                                }
                                                else
                                                {
                                                    if (pst4 != null) cg1.StyleId = pst4.ObjectId;
                                                    if (plst4 != null) cg1.LabelStyleId = plst4.ObjectId;
                                                    cg1.Layer = layer_pt4;
                                                }

                                            }

                                        }
                                    }
                                }
                            }
                            Trans1.Commit();
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                this.MdiParent.WindowState = FormWindowState.Normal;
                Editor1.SetImpliedSelection(Empty_array);
                Editor1.WriteMessage("\nCommand:");
                set_enable_true();

            }
        }

        private void button_load_dt_analize_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application Excel1 = null;
            Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;
            Microsoft.Office.Interop.Excel.Worksheet W1 = null;

            using (OpenFileDialog fbd = new OpenFileDialog())
            {
                fbd.Multiselect = false;
                fbd.Filter = "xls files (*.xlsx)|*.xlsx|csv files (*.csv)|*.csv";
                fbd.FilterIndex = 2;
                fbd.CheckFileExists = true;
                fbd.CheckPathExists = true;

                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {

                    string csv_file = fbd.FileName;

                    try
                    {

                        bool csv_is_opened = false;

                        csv_file = fbd.FileName;
                        try
                        {
                            Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                            foreach (Microsoft.Office.Interop.Excel.Workbook Workbook2 in Excel1.Workbooks)
                            {
                                if (Workbook2.FullName.ToLower() == csv_file.ToLower())
                                {
                                    Workbook1 = Workbook2;
                                    csv_is_opened = true;
                                }

                            }

                        }
                        catch (System.Exception ex)
                        {
                            Excel1 = new Microsoft.Office.Interop.Excel.Application();
                        }

                        if (Functions.is_dan_popescu() == false)
                        {
                            if (Excel1.Workbooks.Count == 0)
                            {
                                Excel1.Visible = false;
                            }
                            else
                            {
                                Excel1.Visible = true;
                            }
                        }
                        else
                        {
                            Excel1.Visible = true;
                        }


                        if (Workbook1 == null)
                        {
                            Workbook1 = Excel1.Workbooks.Open(fbd.FileName);
                        }

                        W1 = Workbook1.Worksheets[1];

                        dt_analize = new System.Data.DataTable();
                        dt_analize.Columns.Add(col_pn, typeof(string));
                        dt_analize.Columns.Add(col_y, typeof(double));
                        dt_analize.Columns.Add(col_x, typeof(double));
                        dt_analize.Columns.Add(col_z, typeof(double));
                        dt_analize.Columns.Add(col_descr, typeof(string));
                        dt_analize.Columns.Add(col_new_description, typeof(string));
                        dt_analize.Columns.Add(col_new_elev, typeof(double));
                        string Col1 = "A";

                        Microsoft.Office.Interop.Excel.Range range2 = W1.Range[Col1 + "1:" + Col1 + "300000"];
                        object[,] values2 = new object[300000, 1];
                        values2 = range2.Value2;

                        bool is_data = false;
                        for (int i = 1; i <= values2.Length; ++i)
                        {
                            object Valoare2 = values2[i, 1];
                            if (Valoare2 != null)
                            {
                                dt_analize.Rows.Add();
                                is_data = true;
                            }
                            else
                            {
                                i = values2.Length + 1;
                            }
                        }

                        if (is_data == false)
                        {
                            dt_analize = null;
                        }

                        int NrR = dt_analize.Rows.Count;

                        Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A1:G" + Convert.ToString(NrR)];
                        object[,] values = new object[NrR, 7];
                        values = range1.Value2;

                        for (int i = 0; i < dt_analize.Rows.Count; ++i)
                        {
                            for (int j = 0; j < dt_analize.Columns.Count; ++j)
                            {
                                object val = values[i + 1, j + 1];
                                if (val == null) val = DBNull.Value;
                                if (j == 1 || j == 2 || j == 3 || j == 6)
                                {
                                    if (val == DBNull.Value || Functions.IsNumeric(Convert.ToString(val)) == false) val = DBNull.Value;
                                }
                                dt_analize.Rows[i][j] = val;
                            }
                        }

                        if (csv_is_opened == false)
                        {
                            Workbook1.Close();
                        }
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    finally
                    {
                        if (W1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W1);
                        if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                    }

                    if (Excel1 != null && Excel1.Workbooks.Count == 0)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                    }
                    //Functions.Transfer_datatable_to_new_excel_spreadsheet(dt_analize);
                }
            }
        }

        private void button_reset_fce_Click(object sender, EventArgs e)
        {

            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.Civil.ApplicationServices.CivilDocument CivilDrawing = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument;

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
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                        this.MdiParent.WindowState = FormWindowState.Minimized;

                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_cogopts = null;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez1 = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez1.MessageForAdding = "\nSelect points";
                        Prompt_rez1.SingleOnly = false;

                        Rezultat_cogopts = ThisDrawing.Editor.GetSelection(Prompt_rez1);
                        if (Rezultat_cogopts.Status != PromptStatus.OK)
                        {
                            set_enable_true();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            this.MdiParent.WindowState = FormWindowState.Normal;
                            return;
                        }


                        #region CIVIL STYLES AND LAYERS
                        Autodesk.Civil.DatabaseServices.Styles.PointStyleCollection col_point_styles = CivilDrawing.Styles.PointStyles;
                        IEnumerator<ObjectId> enum1 = col_point_styles.GetEnumerator();

                        Autodesk.Civil.DatabaseServices.Styles.PointStyle pst1 = null;
                        Autodesk.Civil.DatabaseServices.Styles.PointStyle pst2 = null;
                        Autodesk.Civil.DatabaseServices.Styles.PointStyle pst4 = null;

                        while (enum1.MoveNext())
                        {
                            Autodesk.Civil.DatabaseServices.Styles.PointStyle pst11 = Trans1.GetObject(enum1.Current, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.PointStyle;
                            if (pst11.Name == comboBox_point_style1.Text)
                            {
                                pst1 = pst11;
                            }
                            if (pst11.Name == comboBox_point_style2.Text)
                            {
                                pst2 = pst11;
                            }
                            if (pst11.Name == comboBox_point_style4.Text)
                            {
                                pst4 = pst11;
                            }
                        }

                        Autodesk.Civil.DatabaseServices.Styles.LabelStyleCollection col_point_label_styles = CivilDrawing.Styles.LabelStyles.PointLabelStyles.LabelStyles;

                        Autodesk.Civil.DatabaseServices.Styles.LabelStyle plst1 = null;
                        Autodesk.Civil.DatabaseServices.Styles.LabelStyle plst2 = null;
                        Autodesk.Civil.DatabaseServices.Styles.LabelStyle plst4 = null;

                        IEnumerator<ObjectId> enum11 = col_point_label_styles.GetEnumerator();

                        while (enum11.MoveNext())
                        {
                            Autodesk.Civil.DatabaseServices.Styles.LabelStyle plst11 = Trans1.GetObject(enum11.Current, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.LabelStyle;


                            if (plst11.Name == comboBox_point_label_style1.Text)
                            {
                                plst1 = plst11;
                            }

                            if (plst11.Name == comboBox_point_label_style2.Text)
                            {
                                plst2 = plst11;
                            }

                            if (plst11.Name == comboBox_point_label_style4.Text)
                            {
                                plst4 = plst11;
                            }
                        }

                        string layer_pt1 = "0";
                        if (comboBox_point_layer1.Text.Length > 0)
                        {
                            layer_pt1 = comboBox_point_layer1.Text;

                        }


                        string layer_pt2 = "0";
                        if (comboBox_point_layer2.Text.Length > 0)
                        {
                            layer_pt2 = comboBox_point_layer2.Text;

                        }

                        string layer_pt4 = "0";
                        if (comboBox_point_layer4.Text.Length > 0)
                        {
                            layer_pt4 = comboBox_point_layer4.Text;

                        }


                        UDPString udp1 = Functions.Find_udp_string(comboBox_udp_field_1.Text);


                        UDPDouble udp2 = Functions.Find_udp_double(comboBox_udp_field_2.Text);


                        #endregion



                        for (int i = 0; i < Rezultat_cogopts.Value.Count; ++i)
                        {
                            ObjectId id1 = Rezultat_cogopts.Value[i].ObjectId;
                            CogoPoint cg1 = Trans1.GetObject(id1, OpenMode.ForWrite) as CogoPoint;
                            if (cg1 != null)
                            {
                                double z1 = cg1.Elevation;
                                if (udp2 != null)
                                {
                                    double z0 = cg1.GetUDPValue(udp2);

                                    if (Math.Round(z0, 4) < Math.Round(z1, 4))
                                    {
                                        cg1.RawDescription = "FILL";
                                        if (pst1 != null) cg1.StyleId = pst1.ObjectId;
                                        if (plst1 != null) cg1.LabelStyleId = plst1.ObjectId;
                                        cg1.Layer = layer_pt1;
                                    }
                                    else if (Math.Round(z0, 4) > Math.Round(z1, 4))
                                    {
                                        cg1.RawDescription = "CUT";
                                        if (pst2 != null) cg1.StyleId = pst2.ObjectId;
                                        if (plst2 != null) cg1.LabelStyleId = plst2.ObjectId;
                                        cg1.Layer = layer_pt2;
                                    }
                                    else
                                    {
                                        cg1.RawDescription = "EG";
                                        if (pst4 != null) cg1.StyleId = pst4.ObjectId;
                                        if (plst4 != null) cg1.LabelStyleId = plst4.ObjectId;
                                        cg1.Layer = layer_pt4;
                                    }
                                }

                            }
                        }
                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            this.MdiParent.WindowState = FormWindowState.Normal;
            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            set_enable_true();


        }

        private void button_set_daylight_Click(object sender, EventArgs e)
        {

            string col_processed = "processed";

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


                        #region CIVIL STYLES AND LAYERS
                        Autodesk.Civil.DatabaseServices.Styles.PointStyleCollection col_point_styles = CivilDrawing.Styles.PointStyles;
                        IEnumerator<ObjectId> enum1 = col_point_styles.GetEnumerator();

                        Autodesk.Civil.DatabaseServices.Styles.PointStyle pst1 = null;
                        Autodesk.Civil.DatabaseServices.Styles.PointStyle pst2 = null;
                        Autodesk.Civil.DatabaseServices.Styles.PointStyle pst4 = null;
                        Autodesk.Civil.DatabaseServices.Styles.PointStyle pst5 = null;

                        while (enum1.MoveNext())
                        {
                            Autodesk.Civil.DatabaseServices.Styles.PointStyle pst11 = Trans1.GetObject(enum1.Current, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.PointStyle;
                            if (pst11.Name == comboBox_point_style1.Text)
                            {
                                pst1 = pst11;
                            }
                            if (pst11.Name == comboBox_point_style2.Text)
                            {
                                pst2 = pst11;
                            }
                            if (pst11.Name == comboBox_point_style4.Text)
                            {
                                pst4 = pst11;
                            }
                            if (pst11.Name == comboBox_point_style5.Text)
                            {
                                pst5 = pst11;
                            }
                        }

                        Autodesk.Civil.DatabaseServices.Styles.LabelStyleCollection col_point_label_styles = CivilDrawing.Styles.LabelStyles.PointLabelStyles.LabelStyles;

                        Autodesk.Civil.DatabaseServices.Styles.LabelStyle plst1 = null;
                        Autodesk.Civil.DatabaseServices.Styles.LabelStyle plst2 = null;
                        Autodesk.Civil.DatabaseServices.Styles.LabelStyle plst4 = null;
                        Autodesk.Civil.DatabaseServices.Styles.LabelStyle plst5 = null;

                        IEnumerator<ObjectId> enum11 = col_point_label_styles.GetEnumerator();

                        while (enum11.MoveNext())
                        {
                            Autodesk.Civil.DatabaseServices.Styles.LabelStyle plst11 = Trans1.GetObject(enum11.Current, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.LabelStyle;


                            if (plst11.Name == comboBox_point_label_style1.Text)
                            {
                                plst1 = plst11;
                            }

                            if (plst11.Name == comboBox_point_label_style2.Text)
                            {
                                plst2 = plst11;
                            }

                            if (plst11.Name == comboBox_point_label_style4.Text)
                            {
                                plst4 = plst11;
                            }
                            if (plst11.Name == comboBox_point_label_style5.Text)
                            {
                                plst5 = plst11;
                            }
                        }

                        string layer_pt1 = "0";
                        if (comboBox_point_layer1.Text.Length > 0)
                        {
                            layer_pt1 = comboBox_point_layer1.Text;

                        }


                        string layer_pt2 = "0";
                        if (comboBox_point_layer2.Text.Length > 0)
                        {
                            layer_pt2 = comboBox_point_layer2.Text;

                        }

                        string layer_pt4 = "0";
                        if (comboBox_point_layer4.Text.Length > 0)
                        {
                            layer_pt4 = comboBox_point_layer4.Text;

                        }

                        string layer_pt5 = "0";
                        if (comboBox_point_layer5.Text.Length > 0)
                        {
                            layer_pt5 = comboBox_point_layer5.Text;

                        }


                        UDPString udp1 = Functions.Find_udp_string(comboBox_udp_field_1.Text);


                        UDPDouble udp2 = Functions.Find_udp_double(comboBox_udp_field_2.Text);


                        #endregion





                        System.Data.DataTable dt1 = new System.Data.DataTable();
                        dt1.Columns.Add(col_pn, typeof(string));
                        dt1.Columns.Add(col_y, typeof(double));
                        dt1.Columns.Add(col_x, typeof(double));
                        dt1.Columns.Add(col_z, typeof(double));
                        dt1.Columns.Add(col_descr, typeof(string));
                        dt1.Columns.Add(col_new_elev, typeof(double));
                        dt1.Columns.Add(col_new_description, typeof(string));
                        dt1.Columns.Add(col_col, typeof(int));
                        dt1.Columns.Add(col_row, typeof(int));
                        dt1.Columns.Add(col_id, typeof(ObjectId));
                        dt1.Columns.Add(col_processed, typeof(bool));



                        double ymin = -1.234;
                        double ymax = -1.234;

                        double xmin = -1.234;
                        double xmax = -1.234;

                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {
                            CogoPoint cg1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as CogoPoint;
                            if (cg1 != null)
                            {
                                dt1.Rows.Add();
                                dt1.Rows[dt1.Rows.Count - 1][col_pn] = cg1.PointNumber;
                                dt1.Rows[dt1.Rows.Count - 1][col_id] = cg1.ObjectId;

                                double x1 = Math.Round(cg1.Location.X, 4);
                                double y1 = Math.Round(cg1.Location.Y, 4);

                                dt1.Rows[dt1.Rows.Count - 1][col_x] = x1;
                                dt1.Rows[dt1.Rows.Count - 1][col_y] = y1;
                                dt1.Rows[dt1.Rows.Count - 1][col_z] = Math.Round(cg1.Elevation, 4);
                                dt1.Rows[dt1.Rows.Count - 1][col_descr] = cg1.RawDescription;
                                dt1.Rows[dt1.Rows.Count - 1][col_processed] = false;

                                string udp1_string = "";
                                double udp2_double = -1.234;

                                if (udp1 != null)
                                {
                                    udp1_string = Convert.ToString(cg1.GetUDPValue(udp1));
                                }

                                if (udp2 != null)
                                {
                                    udp2_double = Convert.ToDouble(cg1.GetUDPValue(udp2));
                                }

                                if (udp1_string != "")
                                {
                                    dt1.Rows[dt1.Rows.Count - 1][col_new_description] = udp1_string;
                                }

                                if (udp2_double != -1.234)
                                {
                                    dt1.Rows[dt1.Rows.Count - 1][col_new_elev] = Math.Round(udp2_double, 4);
                                }


                                if (xmin == -1.234)
                                {
                                    xmin = x1;
                                }
                                if (xmax == -1.234)
                                {
                                    xmax = x1;
                                }
                                if (ymin == -1.234)
                                {
                                    ymin = y1;
                                }
                                if (ymax == -1.234)
                                {
                                    ymax = y1;
                                }


                                if (x1 < xmin)
                                {
                                    xmin = x1;
                                }
                                if (x1 > xmax)
                                {
                                    xmax = x1;
                                }
                                if (y1 < ymin)
                                {
                                    ymin = y1;
                                }
                                if (y1 > ymax)
                                {
                                    ymax = y1;
                                }

                            }
                        }



                        int no_rows = Convert.ToInt32(1 + (ymax - ymin) / gridV);
                        int no_cols = Convert.ToInt32(1 + (xmax - xmin) / gridH);

                        for (int i = 0; i < dt1.Rows.Count; ++i)
                        {
                            double x1 = Convert.ToDouble(dt1.Rows[i][col_x]);
                            double y1 = Convert.ToDouble(dt1.Rows[i][col_y]);

                            int col1 = no_cols - Convert.ToInt32((xmax - x1) / gridH);
                            int row1 = no_rows - Convert.ToInt32((ymax - y1) / gridV);
                            dt1.Rows[i][col_col] = col1;
                            dt1.Rows[i][col_row] = row1;
                        }


                        System.Data.DataTable dt2 = null;

                        using (System.Data.DataTable Data_table_temp = dt1.Clone())
                        {
                            DataView dv = new DataView(dt1);
                            dv.Sort = col_row + "," + col_col;
                            for (int i = 0; i < dv.Count; ++i)
                            {
                                System.Data.DataRow Data_row1 = dv[i].Row;
                                Data_table_temp.Rows.Add();
                                for (int j = 0; j < dt1.Columns.Count; ++j)
                                {
                                    Data_table_temp.Rows[Data_table_temp.Rows.Count - 1][j] = Data_row1[j];
                                }
                            }

                            dt2 = Data_table_temp.Copy();

                        }

                        for (int i = 0; i < dt2.Rows.Count - 1; ++i)
                        {
                            if (dt2.Rows[i][col_new_elev] != DBNull.Value &&
                                dt2.Rows[i + 1][col_new_elev] != DBNull.Value &&
                                dt2.Rows[i][col_z] != DBNull.Value &&
                                dt2.Rows[i + 1][col_z] != DBNull.Value &&
                                dt2.Rows[i][col_col] != DBNull.Value &&
                                dt2.Rows[i][col_row] != DBNull.Value &&
                                dt2.Rows[i + 1][col_row] != DBNull.Value &&
                                dt2.Rows[i][col_x] != DBNull.Value &&
                                dt2.Rows[i + 1][col_x] != DBNull.Value &&
                                dt2.Rows[i][col_y] != DBNull.Value &&
                                dt2.Rows[i + 1][col_y] != DBNull.Value &&
                                dt2.Rows[i][col_descr] != DBNull.Value &&
                                dt2.Rows[i + 1][col_descr] != DBNull.Value &&
                                dt2.Rows[i][col_id] != DBNull.Value &&
                                dt2.Rows[i][col_processed] != DBNull.Value)

                            {
                                double z_orig = Convert.ToDouble(dt2.Rows[i][col_new_elev]);
                                double z_current = Convert.ToDouble(dt2.Rows[i][col_z]);
                                int col1 = Convert.ToInt32(dt2.Rows[i][col_col]);
                                int row1 = Convert.ToInt32(dt2.Rows[i][col_row]);
                                int row2 = Convert.ToInt32(dt2.Rows[i + 1][col_row]);
                                string descr1 = Convert.ToString(dt2.Rows[i][col_descr]);
                                bool processed1 = Convert.ToBoolean(dt2.Rows[i][col_processed]);
                                string descr2 = Convert.ToString(dt2.Rows[i + 1][col_descr]);
                                double z_orig2 = Convert.ToDouble(dt2.Rows[i + 1][col_new_elev]);
                                double z_current2 = Convert.ToDouble(dt2.Rows[i + 1][col_z]);
                                string pn1 = Convert.ToString(dt2.Rows[i][col_pn]);
                                string pn2 = Convert.ToString(dt2.Rows[i + 1][col_pn]);
                                double x1 = Convert.ToDouble(dt2.Rows[i][col_x]);
                                double x2 = Convert.ToDouble(dt2.Rows[i + 1][col_x]);
                                double y1 = Convert.ToDouble(dt2.Rows[i][col_y]);
                                double y2 = Convert.ToDouble(dt2.Rows[i + 1][col_y]);




                                if (pn1 == "")
                                {

                                }

                                if (processed1 == false)
                                {
                                    if (row1 == row2)
                                    {
                                        if (z_current == z_orig)
                                        {

                                            ObjectId id1 = (ObjectId)dt2.Rows[i][col_id];
                                            CogoPoint cg1 = Trans1.GetObject(id1, OpenMode.ForWrite) as CogoPoint;

                                            if (cg1 != null)
                                            {
                                                if (z_orig2 != z_current2 && Math.Abs(Math.Round(x1 - x2, 0)) == Math.Round(Convert.ToDouble(textBox_H.Text), 0))
                                                {
                                                    cg1.RawDescription = "DAYLIGHT";
                                                    if (pst5 != null) cg1.StyleId = pst5.ObjectId;
                                                    if (plst5 != null) cg1.LabelStyleId = plst5.ObjectId;
                                                    cg1.Layer = layer_pt5;
                                                    dt2.Rows[i][col_processed] = true;
                                                }
                                                else
                                                {
                                                    if (descr1 != "EG")
                                                    {
                                                        cg1.RawDescription = "EG";
                                                        if (pst4 != null) cg1.StyleId = pst4.ObjectId;
                                                        if (plst4 != null) cg1.LabelStyleId = plst4.ObjectId;
                                                        cg1.Layer = layer_pt4;
                                                    }
                                                }
                                            }
                                        }
                                    }

                                }


                            }
                        }

                        using (System.Data.DataTable Data_table_temp = dt2.Clone())
                        {
                            DataView dv = new DataView(dt2);
                            dv.Sort = col_row + " DESC," + col_col + " DESC";
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

                        }

                        for (int i = 0; i < dt2.Rows.Count - 1; ++i)
                        {
                            if (dt2.Rows[i][col_new_elev] != DBNull.Value &&
                                dt2.Rows[i + 1][col_new_elev] != DBNull.Value &&
                                dt2.Rows[i][col_z] != DBNull.Value &&
                                dt2.Rows[i + 1][col_z] != DBNull.Value &&
                                dt2.Rows[i][col_col] != DBNull.Value &&
                                dt2.Rows[i][col_row] != DBNull.Value &&
                                dt2.Rows[i + 1][col_row] != DBNull.Value &&
                                dt2.Rows[i][col_x] != DBNull.Value &&
                                dt2.Rows[i + 1][col_x] != DBNull.Value &&
                                dt2.Rows[i][col_y] != DBNull.Value &&
                                dt2.Rows[i + 1][col_y] != DBNull.Value &&
                                dt2.Rows[i][col_descr] != DBNull.Value &&
                                dt2.Rows[i + 1][col_descr] != DBNull.Value &&
                                dt2.Rows[i][col_id] != DBNull.Value &&
                                dt2.Rows[i][col_processed] != DBNull.Value)

                            {
                                double z_orig = Convert.ToDouble(dt2.Rows[i][col_new_elev]);
                                double z_current = Convert.ToDouble(dt2.Rows[i][col_z]);
                                int col1 = Convert.ToInt32(dt2.Rows[i][col_col]);
                                int row1 = Convert.ToInt32(dt2.Rows[i][col_row]);
                                string descr1 = Convert.ToString(dt2.Rows[i][col_descr]);
                                bool processed1 = Convert.ToBoolean(dt2.Rows[i][col_processed]);
                                string descr2 = Convert.ToString(dt2.Rows[i + 1][col_descr]);
                                int row2 = Convert.ToInt32(dt2.Rows[i + 1][col_row]);
                                double z_orig2 = Convert.ToDouble(dt2.Rows[i + 1][col_new_elev]);
                                double z_current2 = Convert.ToDouble(dt2.Rows[i + 1][col_z]);
                                string pn1 = Convert.ToString(dt2.Rows[i][col_pn]);
                                string pn2 = Convert.ToString(dt2.Rows[i + 1][col_pn]);
                                double x1 = Convert.ToDouble(dt2.Rows[i][col_x]);
                                double x2 = Convert.ToDouble(dt2.Rows[i + 1][col_x]);
                                double y1 = Convert.ToDouble(dt2.Rows[i][col_y]);
                                double y2 = Convert.ToDouble(dt2.Rows[i + 1][col_y]);
                                if (pn1 == "")
                                {

                                }

                                if (processed1 == false)
                                {
                                    if (row1 == row2)
                                    {
                                        if (z_current == z_orig)
                                        {

                                            ObjectId id1 = (ObjectId)dt2.Rows[i][col_id];
                                            CogoPoint cg1 = Trans1.GetObject(id1, OpenMode.ForWrite) as CogoPoint;

                                            if (cg1 != null)
                                            {
                                                if (z_orig2 != z_current2 && Math.Abs(Math.Round(x1 - x2, 0)) == Math.Round(Convert.ToDouble(textBox_H.Text), 0))
                                                {
                                                    cg1.RawDescription = "DAYLIGHT";
                                                    if (pst5 != null) cg1.StyleId = pst5.ObjectId;
                                                    if (plst5 != null) cg1.LabelStyleId = plst5.ObjectId;
                                                    cg1.Layer = layer_pt5;
                                                    dt2.Rows[i][col_processed] = true;
                                                }
                                                else
                                                {
                                                    if (descr1 != "EG")
                                                    {
                                                        cg1.RawDescription = "EG";
                                                        if (pst4 != null) cg1.StyleId = pst4.ObjectId;
                                                        if (plst4 != null) cg1.LabelStyleId = plst4.ObjectId;
                                                        cg1.Layer = layer_pt4;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }


                        using (System.Data.DataTable Data_table_temp = dt1.Clone())
                        {
                            DataView dv = new DataView(dt1);
                            dv.Sort = col_col + "," + col_row;
                            for (int i = 0; i < dv.Count; ++i)
                            {
                                System.Data.DataRow Data_row1 = dv[i].Row;
                                Data_table_temp.Rows.Add();
                                for (int j = 0; j < dt1.Columns.Count; ++j)
                                {
                                    Data_table_temp.Rows[Data_table_temp.Rows.Count - 1][j] = Data_row1[j];
                                }
                            }
                            dt2 = Data_table_temp.Copy();
                        }

                        for (int i = 0; i < dt2.Rows.Count - 1; ++i)
                        {
                            if (dt2.Rows[i][col_new_elev] != DBNull.Value &&
                                dt2.Rows[i + 1][col_new_elev] != DBNull.Value &&
                                dt2.Rows[i][col_z] != DBNull.Value &&
                                dt2.Rows[i + 1][col_z] != DBNull.Value &&
                                dt2.Rows[i][col_col] != DBNull.Value &&
                                dt2.Rows[i + 1][col_col] != DBNull.Value &&
                                dt2.Rows[i][col_row] != DBNull.Value &&
                                dt2.Rows[i][col_x] != DBNull.Value &&
                                dt2.Rows[i + 1][col_x] != DBNull.Value &&
                                dt2.Rows[i][col_y] != DBNull.Value &&
                                dt2.Rows[i + 1][col_y] != DBNull.Value &&
                                dt2.Rows[i][col_descr] != DBNull.Value &&
                                dt2.Rows[i + 1][col_descr] != DBNull.Value &&
                                dt2.Rows[i][col_id] != DBNull.Value &&
                                dt2.Rows[i][col_processed] != DBNull.Value)

                            {
                                double z_orig = Convert.ToDouble(dt2.Rows[i][col_new_elev]);
                                double z_current = Convert.ToDouble(dt2.Rows[i][col_z]);
                                int col1 = Convert.ToInt32(dt2.Rows[i][col_col]);
                                int col2 = Convert.ToInt32(dt2.Rows[i + 1][col_col]);
                                int row1 = Convert.ToInt32(dt2.Rows[i][col_row]);
                                string descr1 = Convert.ToString(dt2.Rows[i][col_descr]);
                                bool processed1 = Convert.ToBoolean(dt2.Rows[i][col_processed]);
                                string descr2 = Convert.ToString(dt2.Rows[i + 1][col_descr]);
                                double z_orig2 = Convert.ToDouble(dt2.Rows[i + 1][col_new_elev]);
                                double z_current2 = Convert.ToDouble(dt2.Rows[i + 1][col_z]);

                                string pn1 = Convert.ToString(dt2.Rows[i][col_pn]);
                                string pn2 = Convert.ToString(dt2.Rows[i + 1][col_pn]);
                                double x1 = Convert.ToDouble(dt2.Rows[i][col_x]);
                                double x2 = Convert.ToDouble(dt2.Rows[i + 1][col_x]);
                                double y1 = Convert.ToDouble(dt2.Rows[i][col_y]);
                                double y2 = Convert.ToDouble(dt2.Rows[i + 1][col_y]);
                                if (pn1 == "")
                                {

                                }

                                if (processed1 == false)
                                {
                                    if (col1 == col2)
                                    {
                                        if (z_current == z_orig)
                                        {

                                            ObjectId id1 = (ObjectId)dt2.Rows[i][col_id];
                                            CogoPoint cg1 = Trans1.GetObject(id1, OpenMode.ForWrite) as CogoPoint;

                                            if (cg1 != null)
                                            {

                                                if (z_current2 != z_orig2 && Math.Abs(Math.Round(y1 - y2, 0)) == Math.Round(Convert.ToDouble(textBox_V.Text), 0))
                                                {
                                                    cg1.RawDescription = "DAYLIGHT";
                                                    if (pst5 != null) cg1.StyleId = pst5.ObjectId;
                                                    if (plst5 != null) cg1.LabelStyleId = plst5.ObjectId;
                                                    cg1.Layer = layer_pt5;
                                                    dt2.Rows[i][col_processed] = true;
                                                }
                                                else
                                                {
                                                    if (descr1 != "EG")
                                                    {
                                                        cg1.RawDescription = "EG";
                                                        if (pst4 != null) cg1.StyleId = pst4.ObjectId;
                                                        if (plst4 != null) cg1.LabelStyleId = plst4.ObjectId;
                                                        cg1.Layer = layer_pt4;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }

                            }
                        }

                        using (System.Data.DataTable Data_table_temp = dt2.Clone())
                        {
                            DataView dv = new DataView(dt2);
                            dv.Sort = col_col + " DESC," + col_row + " DESC";
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

                        }

                        for (int i = 0; i < dt2.Rows.Count - 1; ++i)
                        {
                            if (dt2.Rows[i][col_new_elev] != DBNull.Value &&
                                dt2.Rows[i + 1][col_new_elev] != DBNull.Value &&
                                dt2.Rows[i][col_z] != DBNull.Value &&
                                dt2.Rows[i + 1][col_z] != DBNull.Value &&
                                dt2.Rows[i][col_col] != DBNull.Value &&
                                dt2.Rows[i + 1][col_col] != DBNull.Value &&
                                dt2.Rows[i][col_row] != DBNull.Value &&
                                dt2.Rows[i][col_x] != DBNull.Value &&
                                dt2.Rows[i + 1][col_x] != DBNull.Value &&
                                dt2.Rows[i][col_y] != DBNull.Value &&
                                dt2.Rows[i + 1][col_y] != DBNull.Value &&
                                dt2.Rows[i][col_descr] != DBNull.Value &&
                                dt2.Rows[i + 1][col_descr] != DBNull.Value &&
                                dt2.Rows[i][col_id] != DBNull.Value &&
                                dt2.Rows[i][col_processed] != DBNull.Value)

                            {
                                double z_orig = Convert.ToDouble(dt2.Rows[i][col_new_elev]);
                                double z_current = Convert.ToDouble(dt2.Rows[i][col_z]);
                                int col1 = Convert.ToInt32(dt2.Rows[i][col_col]);
                                int row1 = Convert.ToInt32(dt2.Rows[i][col_row]);
                                string descr1 = Convert.ToString(dt2.Rows[i][col_descr]);
                                bool processed1 = Convert.ToBoolean(dt2.Rows[i][col_processed]);
                                string descr2 = Convert.ToString(dt2.Rows[i + 1][col_descr]);
                                int col2 = Convert.ToInt32(dt2.Rows[i + 1][col_col]);
                                double z_orig2 = Convert.ToDouble(dt2.Rows[i + 1][col_new_elev]);
                                double z_current2 = Convert.ToDouble(dt2.Rows[i + 1][col_z]);

                                string pn1 = Convert.ToString(dt2.Rows[i][col_pn]);
                                string pn2 = Convert.ToString(dt2.Rows[i + 1][col_pn]);

                                double x1 = Convert.ToDouble(dt2.Rows[i][col_x]);
                                double x2 = Convert.ToDouble(dt2.Rows[i + 1][col_x]);
                                double y1 = Convert.ToDouble(dt2.Rows[i][col_y]);
                                double y2 = Convert.ToDouble(dt2.Rows[i + 1][col_y]);
                                if (pn1 == "")
                                {

                                }

                                if (processed1 == false)
                                {
                                    if (col1 == col2)
                                    {
                                        if (z_current == z_orig)
                                        {

                                            ObjectId id1 = (ObjectId)dt2.Rows[i][col_id];
                                            CogoPoint cg1 = Trans1.GetObject(id1, OpenMode.ForWrite) as CogoPoint;

                                            if (cg1 != null)
                                            {
                                                if (z_orig2 != z_current2 && Math.Abs(Math.Round(y1 - y2, 0)) == Math.Round(Convert.ToDouble(textBox_V.Text), 0))
                                                {
                                                    cg1.RawDescription = "DAYLIGHT";
                                                    if (pst5 != null) cg1.StyleId = pst5.ObjectId;
                                                    if (plst5 != null) cg1.LabelStyleId = plst5.ObjectId;
                                                    cg1.Layer = layer_pt5;
                                                    dt2.Rows[i][col_processed] = true;
                                                }
                                                else
                                                {
                                                    if (descr1 != "EG")
                                                    {
                                                        cg1.RawDescription = "EG";
                                                        if (pst4 != null) cg1.StyleId = pst4.ObjectId;
                                                        if (plst4 != null) cg1.LabelStyleId = plst4.ObjectId;
                                                        cg1.Layer = layer_pt4;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }


                            }
                        }



                        Trans1.Commit();


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

        private void button_global_load_Click(object sender, EventArgs e)
        {


            comboBox_point_style3.Items.Clear();
            comboBox_point_label_style3.Items.Clear();
            comboBox_point_style4.Items.Clear();
            comboBox_point_label_style4.Items.Clear();
            comboBox_point_style5.Items.Clear();
            comboBox_point_label_style5.Items.Clear();

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
                            comboBox_point_style3.Items.Add(pst1.Name);
                            comboBox_point_style4.Items.Add(pst1.Name);
                            comboBox_point_style5.Items.Add(pst1.Name);
                            comboBox_point_style1.Items.Add(pst1.Name);
                            comboBox_point_style2.Items.Add(pst1.Name);
                        }
                        if (comboBox_point_style3.Items.Contains("Anchor") == true)
                        {
                            comboBox_point_style3.SelectedIndex = comboBox_point_style3.Items.IndexOf("Anchor");
                        }
                        if (comboBox_point_style4.Items.Contains("EG") == true)
                        {
                            comboBox_point_style4.SelectedIndex = comboBox_point_style4.Items.IndexOf("EG");
                        }
                        if (comboBox_point_style5.Items.Contains("Daylight") == true)
                        {
                            comboBox_point_style5.SelectedIndex = comboBox_point_style5.Items.IndexOf("Daylight");
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
                            comboBox_point_label_style3.Items.Add(plst2.Name);
                            comboBox_point_label_style4.Items.Add(plst2.Name);
                            comboBox_point_label_style5.Items.Add(plst2.Name);
                            comboBox_point_label_style1.Items.Add(plst2.Name);
                            comboBox_point_label_style2.Items.Add(plst2.Name);
                        }

                        if (comboBox_point_label_style3.Items.Contains("Point#-Elevation-Description [ANCHOR]") == true)
                        {
                            comboBox_point_label_style3.SelectedIndex = comboBox_point_label_style3.Items.IndexOf("Point#-Elevation-Description [ANCHOR]");
                        }
                        if (comboBox_point_label_style4.Items.Contains("Point#-Elevation-Description [EG]") == true)
                        {
                            comboBox_point_label_style4.SelectedIndex = comboBox_point_label_style4.Items.IndexOf("Point#-Elevation-Description [EG]");
                        }

                        if (comboBox_point_label_style5.Items.Contains("Point#-Elevation-Description [Daylight]") == true)
                        {
                            comboBox_point_label_style5.SelectedIndex = comboBox_point_label_style5.Items.IndexOf("Point#-Elevation-Description [Daylight]");
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
                            if (udp1 != null) comboBox_udp_field_1.Items.Add(udp1.Name);
                        }
                        if (comboBox_udp_field_1.Items.Contains("ATT1") == true)
                        {
                            comboBox_udp_field_1.SelectedIndex = comboBox_udp_field_1.Items.IndexOf("ATT1");
                        }

                        comboBox_udp_field_2.Items.Clear();
                        foreach (UDP udp2 in Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument.PointUDPs)
                        {
                            if (udp2 != null) comboBox_udp_field_2.Items.Add(udp2.Name);
                        }
                        if (comboBox_udp_field_2.Items.Contains("ATT2") == true)
                        {
                            comboBox_udp_field_2.SelectedIndex = comboBox_udp_field_2.Items.IndexOf("ATT2");
                        }

                        Functions.Incarca_existing_layers_to_combobox(comboBox_point_layer3);

                        if (comboBox_point_layer3.Items.Contains("PTS-Anchor") == true)
                        {
                            comboBox_point_layer3.SelectedIndex = comboBox_point_layer3.Items.IndexOf("PTS-Anchor");
                        }

                        Functions.Incarca_existing_layers_to_combobox(comboBox_point_layer4);

                        if (comboBox_point_layer4.Items.Contains("PTS-EG") == true)
                        {
                            comboBox_point_layer4.SelectedIndex = comboBox_point_layer4.Items.IndexOf("PTS-EG");
                        }
                        Functions.Incarca_existing_layers_to_combobox(comboBox_point_layer5);

                        if (comboBox_point_layer5.Items.Contains("PTS-Daylight") == true)
                        {
                            comboBox_point_layer5.SelectedIndex = comboBox_point_layer5.Items.IndexOf("PTS-Daylight");
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
