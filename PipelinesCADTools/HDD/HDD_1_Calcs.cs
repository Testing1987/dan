using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Alignment_mdi
{

    public partial class HDD_calcs : Form
    {

        List<object> OD_lista_val = null;
        List<Autodesk.Gis.Map.Constants.DataType> OD_lista_type = null;
        string OD_table_name = "HDD_DESIGN_v1";


        Polyline Poly_Graph;
        Point3d pt_1;
        Point3d pt_2;
        Point3d pt_3;
        Point3d pt_4;
        Point3d pt_5;
        Point3d pt_6;


        double hexag = 1;
        double vexag = 1;
        double known_y1 = -123.1234567;
        double known_el1 = -123.1234567;
        double known_y2 = -123.1234567;
        double known_el2 = -123.1234567;
        double known_x1 = -123.1234567;
        double known_x2 = -123.1234567;
        double known_sta1 = -123.1234567;
        double known_sta2 = -123.1234567;
        System.Data.DataTable dt_steq = null;



        int round1 = 2;
        int lr = 1;
       
        int cid = 7;

        double elev1 = 0;
        double sta_original = 0;
        double sta1 = 0;
        double eq_sta1 = 0;


        double hd1 = 0;
        double L1 = 0;
        double h1 = 0;
        double Angle1 = 0;
        double elev2 = 0;
        double slope2 = 0;
        double hd2 = 0;
        double L2 = 0;
        double h2 = 0;
        double radius1 = 0;
        double arc_len1 = 0;
        double elev3 = 0;
        double slope3 = 0;
        double hd3 = 0;
        double hda = 0;
        double L3 = 0;
        double h4a = 0;
        double elev4 = 0;
        double slope4 = 0;
        double hd4 = 0;
        double L4 = 0;
        double h4 = 0;
        double Angle2 = 0;
        double radius2 = 0;
        double elev5 = 0;
        double slope5 = 0;
        double hd5 = 0;
        double arc_len2 = 0;
        double hd6 = 0;
        double L5 = 0;
        double h5 = 0;
        double elev6 = 0;
        double sta2 = 0;
        double sta3 = 0;
        double sta4 = 0;
        double sta5 = 0;
        double sta6 = 0;
        double slope6 = 0;

        double elevB = 0;
        double staC = 0;
        double difD = 0;
        double difE = 0;

        Point3d point_referinta;

        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_calc_HDD);
            lista_butoane.Add(button_load_ground_profile);
            lista_butoane.Add(button_draw_HDD);
            lista_butoane.Add(button_label_hdd);

            lista_butoane.Add(textBox_angle_end);
            lista_butoane.Add(textBox_angle_start);
            lista_butoane.Add(textBox_arc_length1);
            lista_butoane.Add(textBox_arc_length2);
            lista_butoane.Add(textBox_dev_angle1);
            lista_butoane.Add(textBox_dev_angle1);
            lista_butoane.Add(textBox_dev_angle2);
            lista_butoane.Add(textBox_elev1);
            lista_butoane.Add(textBox_elev2);
            lista_butoane.Add(textBox_elev3);
            lista_butoane.Add(textBox_elev4);
            lista_butoane.Add(textBox_elev5);
            lista_butoane.Add(textBox_elev6);
            lista_butoane.Add(textBox_h1);
            lista_butoane.Add(textBox_h2);
            lista_butoane.Add(textBox_h4);
            lista_butoane.Add(textBox_h4a);
            lista_butoane.Add(textBox_h5);
            lista_butoane.Add(textBox_hd1);
            lista_butoane.Add(textBox_hd2);
            lista_butoane.Add(textBox_hd3);
            lista_butoane.Add(textBox_hd4);
            lista_butoane.Add(textBox_hd5);
            lista_butoane.Add(textBox_hd6);
            lista_butoane.Add(textBox_L1);
            lista_butoane.Add(textBox_L2);
            lista_butoane.Add(textBox_L3);
            lista_butoane.Add(textBox_L4);
            lista_butoane.Add(textBox_L5);
            lista_butoane.Add(textBox_radius1);
            lista_butoane.Add(textBox_radius2);
            lista_butoane.Add(textBox_sta1);
            lista_butoane.Add(textBox_slope2);
            lista_butoane.Add(textBox_slope3);
            lista_butoane.Add(textBox_slope4);
            lista_butoane.Add(textBox_slope5);
            lista_butoane.Add(textBox_sta6);
            lista_butoane.Add(textBox_sta2);
            lista_butoane.Add(textBox_sta3);
            lista_butoane.Add(textBox_sta4);
            lista_butoane.Add(textBox_sta5);


            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {

                bt1.Enabled = false;

            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_calc_HDD);
            lista_butoane.Add(button_load_ground_profile);
            lista_butoane.Add(button_draw_HDD);
            lista_butoane.Add(button_label_hdd);


            lista_butoane.Add(textBox_angle_end);
            lista_butoane.Add(textBox_angle_start);
            lista_butoane.Add(textBox_arc_length1);
            lista_butoane.Add(textBox_arc_length2);
            lista_butoane.Add(textBox_dev_angle1);
            lista_butoane.Add(textBox_dev_angle1);
            lista_butoane.Add(textBox_dev_angle2);
            lista_butoane.Add(textBox_elev1);
            lista_butoane.Add(textBox_elev2);
            lista_butoane.Add(textBox_elev3);
            lista_butoane.Add(textBox_elev4);
            lista_butoane.Add(textBox_elev5);
            lista_butoane.Add(textBox_elev6);
            lista_butoane.Add(textBox_h1);
            lista_butoane.Add(textBox_h2);
            lista_butoane.Add(textBox_h4a);
            lista_butoane.Add(textBox_h4);
            lista_butoane.Add(textBox_h5);
            lista_butoane.Add(textBox_hd1);
            lista_butoane.Add(textBox_hd2);
            lista_butoane.Add(textBox_hd3);
            lista_butoane.Add(textBox_hd4);
            lista_butoane.Add(textBox_hd5);
            lista_butoane.Add(textBox_hd6);
            lista_butoane.Add(textBox_L1);
            lista_butoane.Add(textBox_L2);
            lista_butoane.Add(textBox_L3);
            lista_butoane.Add(textBox_L4);
            lista_butoane.Add(textBox_L5);
            lista_butoane.Add(textBox_radius1);
            lista_butoane.Add(textBox_radius2);
            lista_butoane.Add(textBox_sta1);
            lista_butoane.Add(textBox_slope2);
            lista_butoane.Add(textBox_slope3);
            lista_butoane.Add(textBox_slope4);
            lista_butoane.Add(textBox_slope5);
            lista_butoane.Add(textBox_sta6);
            lista_butoane.Add(textBox_sta2);
            lista_butoane.Add(textBox_sta3);
            lista_butoane.Add(textBox_sta4);
            lista_butoane.Add(textBox_sta5);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }

        public HDD_calcs()
        {
            InitializeComponent();
            if (Functions.is_dan_popescu() == true)
            {
                textBox_angle_start.Text = "16.00";
                textBox_L1.Text = "295.1";
                textBox_radius1.Text = "1600";
                textBox_L3.Text = "693.8369";
                textBox_radius2.Text = "1600";
                textBox_angle_end.Text = "12";
                textBox_dev_angle1.Text = "0.00";
            }
            comboBox_scales.SelectedIndex = 7;
        }

        private void reset_variables()
        {
            pt_2 = new Point3d();
            pt_3 = new Point3d();
            pt_4 = new Point3d();
            pt_5 = new Point3d();
            pt_6 = new Point3d();

            L1 = 0;
            h1 = 0;
            Angle1 = 0;
            elev2 = 0;
            slope2 = 0;
            hd2 = 0;
            L2 = 0;
            h2 = 0;
            radius1 = 0;
            arc_len1 = 0;
            elev3 = 0;
            slope3 = 0;
            hd3 = 0;
            hda = 0;
            L3 = 0;
            h4a = 0;
            elev4 = 0;
            slope4 = 0;
            hd4 = 0;
            L4 = 0;
            h4 = 0;
            Angle2 = 0;
            radius2 = 0;
            elev5 = 0;
            slope5 = 0;
            hd5 = 0;
            arc_len2 = 0;
            hd6 = 0;
            L5 = 0;
            h5 = 0;
            elev6 = 0;
            sta6 = 0;
            slope6 = 0;

            elevB = 0;
            staC = 0;
            difD = 0;
            difE = 0;
        }

        private void reset_form()
        {
            textBox_h1.Text = "0.00";
            textBox_elev2.Text = "0.00";
            textBox_slope2.Text = "0.00";
            textBox_sta2.Text = "0.00";
            textBox_hd2.Text = "0.00";
            textBox_arc_length1.Text = "0.00";
            textBox_L2.Text = "0.00";
            textBox_h2.Text = "0.00";
            textBox_elev3.Text = "0.00";
            textBox_slope3.Text = "0.00";
            textBox_sta3.Text = "0.00";
            textBox_hd3.Text = "0.00";
            textBox_h4a.Text = "0.00";
            textBox_elev4.Text = "0.00";
            textBox_slope4.Text = "0.00";
            textBox_sta4.Text = "0.00";
            textBox_hd4.Text = "0.00";
            textBox_h4.Text = "0.00";
            textBox_L4.Text = "0.00";
            textBox_arc_length2.Text = "0.00";
            textBox_elev5.Text = "0.00";
            textBox_slope5.Text = "0.00";
            textBox_sta5.Text = "0.00";
            textBox_hd5.Text = "0.00";
            textBox_hd6.Text = "0.00";
            textBox_h5.Text = "0.00";
            textBox_L5.Text = "0.00";
            textBox_elev6.Text = "0.00";
            textBox_sta6.Text = "0.00";
            textBox_slope6.Text = "0.00";
            textBoxB.Text = "0.00";
            textBoxC.Text = "0.00";
            textBoxD.Text = "0.00";
            textBoxE.Text = "0.00";
        }
        private void button_load_ground_profile_Click(object sender, EventArgs e)
        {
            this.MdiParent.WindowState = FormWindowState.Minimized;

            set_enable_false();

            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                Matrix3d CurentUCSmatrix = Editor1.CurrentUserCoordinateSystem;
                ObjectId[] Empty_array = null;
                Editor1.SetImpliedSelection(Empty_array);

                dt_steq = new System.Data.DataTable();
                dt_steq.Columns.Add("Start", typeof(double));
                dt_steq.Columns.Add("Back", typeof(double));
                dt_steq.Columns.Add("Ahead", typeof(double));
                dt_steq.Columns.Add("Increase", typeof(bool));


                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                        Polyline Poly_Graph_exag = null;
                        known_x1 = -123.1234567;
                        known_x2 = -123.1234567;
                        known_y1 = -123.1234567;
                        known_sta1 = -123.1234567;
                        known_sta2 = -123.1234567;
                        known_el1 = -123.1234567;
                        hexag = 1;
                        vexag = 1;



                        if (panel_exaggeration.Visible == false)
                        {

                            #region select grid

                            Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_hor1;
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rezh1 = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rezh1.MessageForAdding = "\nSelect first vertical line (STATION) and the label for it:";
                            Prompt_rezh1.SingleOnly = false;
                            Rezultat_hor1 = ThisDrawing.Editor.GetSelection(Prompt_rezh1);

                            if (Rezultat_hor1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                            {
                                this.MdiParent.WindowState = FormWindowState.Normal;
                                set_enable_true();
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }



                            if (Rezultat_hor1.Value.Count != 2)
                            {
                                MessageBox.Show("the selection has to contain 2 objects. You have selected only" + Rezultat_hor1.Value.Count.ToString() + " object(s).\r\nOperation aborted");
                                this.MdiParent.WindowState = FormWindowState.Normal;
                                set_enable_true();
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }


                            Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_hor2;
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rezh2 = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rezh2.MessageForAdding = "\nSelect the second vertical line (STATION) and the label for it:";
                            Prompt_rezh2.SingleOnly = false;
                            Rezultat_hor2 = ThisDrawing.Editor.GetSelection(Prompt_rezh2);

                            if (Rezultat_hor2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                            {
                                this.MdiParent.WindowState = FormWindowState.Normal;
                                set_enable_true();
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }



                            if (Rezultat_hor2.Value.Count != 2)
                            {
                                MessageBox.Show("the selection has to contain 2 objects. You have selected only" + Rezultat_hor2.Value.Count.ToString() + " object(s).\r\nOperation aborted");
                                this.MdiParent.WindowState = FormWindowState.Normal;
                                set_enable_true();
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }


                            Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_ver1;
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rezv1 = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rezv1.MessageForAdding = "\nSelect first horizontal line (ELEVATION) and the label for it:";
                            Prompt_rezv1.SingleOnly = false;
                            Rezultat_ver1 = ThisDrawing.Editor.GetSelection(Prompt_rezv1);

                            if (Rezultat_ver1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                            {
                                this.MdiParent.WindowState = FormWindowState.Normal;
                                set_enable_true();
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }

                            if (Rezultat_ver1.Value.Count != 2)
                            {
                                MessageBox.Show("the selection has to contain 2 objects. You have selected only" + Rezultat_ver1.Value.Count.ToString() + " object(s).\r\nOperation aborted");
                                this.MdiParent.WindowState = FormWindowState.Normal;
                                set_enable_true();
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }


                            Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_ver2;
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rezv2 = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rezv2.MessageForAdding = "\nSelect second horizontal line (ELEVATION) and the label for it:";
                            Prompt_rezv2.SingleOnly = false;
                            Rezultat_ver2 = ThisDrawing.Editor.GetSelection(Prompt_rezv2);

                            if (Rezultat_ver2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                            {
                                this.MdiParent.WindowState = FormWindowState.Normal;
                                set_enable_true();
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }

                            if (Rezultat_ver2.Value.Count != 2)
                            {
                                MessageBox.Show("the selection has to contain 2 objects. You have selected only" + Rezultat_ver2.Value.Count.ToString() + " object(s).\r\nOperation aborted");
                                this.MdiParent.WindowState = FormWindowState.Normal;
                                set_enable_true();
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }


                            #endregion

                            #region get_parameters for grid
                            known_x2 = -123.1234567;
                            known_y2 = -123.1234567;
                            known_sta2 = -123.1234567;
                            known_el2 = -123.1234567;




                            Entity Ent1 = Trans1.GetObject(Rezultat_hor1.Value[0].ObjectId, OpenMode.ForRead) as Entity;
                            Entity Ent2 = Trans1.GetObject(Rezultat_hor1.Value[1].ObjectId, OpenMode.ForRead) as Entity;

                            Entity Ent3 = Trans1.GetObject(Rezultat_ver1.Value[0].ObjectId, OpenMode.ForRead) as Entity;
                            Entity Ent4 = Trans1.GetObject(Rezultat_ver1.Value[1].ObjectId, OpenMode.ForRead) as Entity;

                            Entity Ent11 = Trans1.GetObject(Rezultat_hor2.Value[0].ObjectId, OpenMode.ForRead) as Entity;
                            Entity Ent12 = Trans1.GetObject(Rezultat_hor2.Value[1].ObjectId, OpenMode.ForRead) as Entity;

                            Entity Ent13 = Trans1.GetObject(Rezultat_ver2.Value[0].ObjectId, OpenMode.ForRead) as Entity;
                            Entity Ent14 = Trans1.GetObject(Rezultat_ver2.Value[1].ObjectId, OpenMode.ForRead) as Entity;



                            if (((Ent1 is Polyline || Ent1 is Line) && (Ent2 is MText || Ent2 is DBText)) || ((Ent2 is Polyline || Ent2 is Line) && (Ent1 is MText || Ent1 is DBText)) ||
                                ((Ent11 is Polyline || Ent11 is Line) && (Ent12 is MText || Ent12 is DBText)) || ((Ent12 is Polyline || Ent12 is Line) && (Ent11 is MText || Ent11 is DBText)))
                            {
                                #region ent1


                                if (Ent1 is Polyline)
                                {
                                    Polyline P1 = Ent1 as Polyline;
                                    if (P1 != null)
                                    {
                                        double x1 = P1.StartPoint.X;
                                        double x2 = P1.EndPoint.X;
                                        if (Math.Round(x1, 2) == Math.Round(x2, 2))
                                        {
                                            known_x1 = x1;

                                        }


                                    }

                                }

                                if (Ent1 is Line)
                                {
                                    Line L1 = Ent1 as Line;
                                    if (L1 != null)
                                    {
                                        double x1 = L1.StartPoint.X;
                                        double x2 = L1.EndPoint.X;
                                        if (Math.Round(x1, 2) == Math.Round(x2, 2))
                                        {
                                            known_x1 = x1;

                                        }


                                    }

                                }

                                if (Ent1 is MText)
                                {
                                    MText M1 = Ent1 as MText;
                                    if (M1 != null)
                                    {
                                        string Continut = M1.Text.Replace("+", "");
                                        if (Functions.IsNumeric(Continut) == true)
                                        {
                                            known_sta1 = Convert.ToDouble(Continut);

                                        }


                                    }

                                }

                                if (Ent1 is DBText)
                                {
                                    DBText T1 = Ent1 as DBText;
                                    if (T1 != null)
                                    {
                                        string Continut = T1.TextString.Replace("+", "");
                                        if (Functions.IsNumeric(Continut) == true)
                                        {
                                            known_sta1 = Convert.ToDouble(Continut);

                                        }


                                    }

                                }


                                if (Ent2 is Polyline)
                                {
                                    Polyline P1 = Ent2 as Polyline;
                                    if (P1 != null)
                                    {
                                        double x1 = P1.StartPoint.X;
                                        double x2 = P1.EndPoint.X;
                                        if (Math.Round(x1, 2) == Math.Round(x2, 2))
                                        {
                                            known_x1 = x1;

                                        }


                                    }

                                }

                                if (Ent2 is Line)
                                {
                                    Line L1 = Ent2 as Line;
                                    if (L1 != null)
                                    {
                                        double x1 = L1.StartPoint.X;
                                        double x2 = L1.EndPoint.X;
                                        if (Math.Round(x1, 2) == Math.Round(x2, 2))
                                        {
                                            known_x1 = x1;

                                        }


                                    }

                                }

                                if (Ent2 is MText)
                                {
                                    MText M1 = Ent2 as MText;
                                    if (M1 != null)
                                    {
                                        string Continut = M1.Text.Replace("+", "");
                                        if (Functions.IsNumeric(Continut) == true)
                                        {
                                            known_sta1 = Convert.ToDouble(Continut);

                                        }


                                    }

                                }

                                if (Ent2 is DBText)
                                {
                                    DBText T1 = Ent2 as DBText;
                                    if (T1 != null)
                                    {
                                        string Continut = T1.TextString.Replace("+", "");
                                        if (Functions.IsNumeric(Continut) == true)
                                        {
                                            known_sta1 = Convert.ToDouble(Continut);

                                        }


                                    }

                                }
                                #endregion

                                #region ent2


                                if (Ent11 is Polyline)
                                {
                                    Polyline P11 = Ent11 as Polyline;
                                    if (P11 != null)
                                    {
                                        double x11 = P11.StartPoint.X;
                                        double x12 = P11.EndPoint.X;
                                        if (Math.Round(x11, 2) == Math.Round(x12, 2))
                                        {
                                            known_x2 = x11;
                                        }
                                    }
                                }

                                if (Ent11 is Line)
                                {
                                    Line L11 = Ent11 as Line;
                                    if (L11 != null)
                                    {
                                        double x11 = L11.StartPoint.X;
                                        double x12 = L11.EndPoint.X;
                                        if (Math.Round(x11, 2) == Math.Round(x12, 2))
                                        {
                                            known_x2 = x11;

                                        }


                                    }

                                }

                                if (Ent11 is MText)
                                {
                                    MText M11 = Ent11 as MText;
                                    if (M11 != null)
                                    {
                                        string Continut = M11.Text.Replace("+", "");
                                        if (Functions.IsNumeric(Continut) == true)
                                        {
                                            known_sta2 = Convert.ToDouble(Continut);

                                        }


                                    }

                                }

                                if (Ent11 is DBText)
                                {
                                    DBText T11 = Ent11 as DBText;
                                    if (T11 != null)
                                    {
                                        string Continut = T11.TextString.Replace("+", "");
                                        if (Functions.IsNumeric(Continut) == true)
                                        {
                                            known_sta2 = Convert.ToDouble(Continut);

                                        }


                                    }

                                }


                                if (Ent12 is Polyline)
                                {
                                    Polyline P12 = Ent12 as Polyline;
                                    if (P12 != null)
                                    {
                                        double x12 = P12.StartPoint.X;
                                        double x22 = P12.EndPoint.X;
                                        if (Math.Round(x12, 2) == Math.Round(x22, 2))
                                        {
                                            known_x2 = x12;

                                        }


                                    }

                                }

                                if (Ent12 is Line)
                                {
                                    Line L12 = Ent12 as Line;
                                    if (L12 != null)
                                    {
                                        double x12 = L12.StartPoint.X;
                                        double x22 = L12.EndPoint.X;
                                        if (Math.Round(x12, 2) == Math.Round(x22, 2))
                                        {
                                            known_x2 = x12;

                                        }


                                    }

                                }

                                if (Ent12 is MText)
                                {
                                    MText M12 = Ent12 as MText;
                                    if (M12 != null)
                                    {
                                        string Continut = M12.Text.Replace("+", "");
                                        if (Functions.IsNumeric(Continut) == true)
                                        {
                                            known_sta2 = Convert.ToDouble(Continut);

                                        }


                                    }

                                }

                                if (Ent12 is DBText)
                                {
                                    DBText T12 = Ent12 as DBText;
                                    if (T12 != null)
                                    {
                                        string Continut = T12.TextString.Replace("+", "");
                                        if (Functions.IsNumeric(Continut) == true)
                                        {
                                            known_sta2 = Convert.ToDouble(Continut);

                                        }


                                    }

                                }
                                #endregion

                            }

                            if (((Ent3 is Polyline || Ent3 is Line) & (Ent4 is MText || Ent4 is DBText)) || ((Ent4 is Polyline || Ent4 is Line) & (Ent3 is MText || Ent3 is DBText)) ||
                                ((Ent13 is Polyline || Ent13 is Line) && (Ent14 is MText || Ent14 is DBText)) || ((Ent14 is Polyline || Ent14 is Line) && (Ent13 is MText || Ent13 is DBText)))
                            {
                                #region ent3

                                if (Ent3 is Polyline)
                                {
                                    Polyline P1 = Ent3 as Polyline;
                                    if (P1 != null)
                                    {
                                        double y1 = P1.StartPoint.Y;
                                        double y2 = P1.EndPoint.Y;
                                        if (Math.Round(y1, 2) == Math.Round(y2, 2))
                                        {
                                            known_y1 = y1;

                                        }


                                    }

                                }

                                if (Ent3 is Line)
                                {
                                    Line L1 = Ent3 as Line;
                                    if (L1 != null)
                                    {
                                        double y1 = L1.StartPoint.Y;
                                        double y2 = L1.EndPoint.Y;
                                        if (Math.Round(y1, 2) == Math.Round(y2, 2))
                                        {
                                            known_y1 = y1;

                                        }


                                    }

                                }

                                if (Ent3 is MText)
                                {
                                    MText M1 = Ent3 as MText;
                                    if (M1 != null)
                                    {
                                        string Continut = M1.Text.Replace("'", "");
                                        if (Functions.IsNumeric(Continut) == true)
                                        {
                                            known_el1 = Convert.ToDouble(Continut);

                                        }


                                    }

                                }

                                if (Ent3 is DBText)
                                {
                                    DBText T1 = Ent3 as DBText;
                                    if (T1 != null)
                                    {
                                        string Continut = T1.TextString.Replace("'", "");
                                        if (Functions.IsNumeric(Continut) == true)
                                        {
                                            known_el1 = Convert.ToDouble(Continut);

                                        }


                                    }

                                }


                                if (Ent4 is Polyline)
                                {
                                    Polyline P1 = Ent4 as Polyline;
                                    if (P1 != null)
                                    {
                                        double y1 = P1.StartPoint.Y;
                                        double y2 = P1.EndPoint.Y;
                                        if (Math.Round(y1, 2) == Math.Round(y2, 2))
                                        {
                                            known_y1 = y1;

                                        }


                                    }

                                }

                                if (Ent4 is Line)
                                {
                                    Line L1 = Ent4 as Line;
                                    if (L1 != null)
                                    {
                                        double y1 = L1.StartPoint.Y;
                                        double y2 = L1.EndPoint.Y;
                                        if (Math.Round(y1, 2) == Math.Round(y2, 2))
                                        {
                                            known_y1 = y1;

                                        }


                                    }

                                }

                                if (Ent4 is MText)
                                {
                                    MText M1 = Ent4 as MText;
                                    if (M1 != null)
                                    {
                                        string Continut = M1.Text.Replace("'", "");
                                        if (Functions.IsNumeric(Continut) == true)
                                        {
                                            known_el1 = Convert.ToDouble(Continut);

                                        }


                                    }

                                }

                                if (Ent4 is DBText)
                                {
                                    DBText T1 = Ent4 as DBText;
                                    if (T1 != null)
                                    {
                                        string Continut = T1.TextString.Replace("'", "");
                                        if (Functions.IsNumeric(Continut) == true)
                                        {
                                            known_el1 = Convert.ToDouble(Continut);

                                        }


                                    }

                                }
                                #endregion

                                #region ent4

                                if (Ent13 is Polyline)
                                {
                                    Polyline P13 = Ent13 as Polyline;
                                    if (P13 != null)
                                    {
                                        double y13 = P13.StartPoint.Y;
                                        double y23 = P13.EndPoint.Y;
                                        if (Math.Round(y13, 2) == Math.Round(y23, 2))
                                        {
                                            known_y2 = y13;

                                        }


                                    }

                                }

                                if (Ent13 is Line)
                                {
                                    Line L13 = Ent13 as Line;
                                    if (L13 != null)
                                    {
                                        double y13 = L13.StartPoint.Y;
                                        double y23 = L13.EndPoint.Y;
                                        if (Math.Round(y13, 2) == Math.Round(y23, 2))
                                        {
                                            known_y2 = y13;

                                        }


                                    }

                                }

                                if (Ent13 is MText)
                                {
                                    MText M13 = Ent13 as MText;
                                    if (M13 != null)
                                    {
                                        string Continut = M13.Text.Replace("'", "");
                                        if (Functions.IsNumeric(Continut) == true)
                                        {
                                            known_el2 = Convert.ToDouble(Continut);

                                        }


                                    }

                                }

                                if (Ent13 is DBText)
                                {
                                    DBText T1 = Ent13 as DBText;
                                    if (T1 != null)
                                    {
                                        string Continut = T1.TextString.Replace("'", "");
                                        if (Functions.IsNumeric(Continut) == true)
                                        {
                                            known_el2 = Convert.ToDouble(Continut);

                                        }


                                    }

                                }


                                if (Ent14 is Polyline)
                                {
                                    Polyline P1 = Ent14 as Polyline;
                                    if (P1 != null)
                                    {
                                        double y1 = P1.StartPoint.Y;
                                        double y2 = P1.EndPoint.Y;
                                        if (Math.Round(y1, 2) == Math.Round(y2, 2))
                                        {
                                            known_y2 = y1;

                                        }


                                    }

                                }

                                if (Ent14 is Line)
                                {
                                    Line L1 = Ent14 as Line;
                                    if (L1 != null)
                                    {
                                        double y1 = L1.StartPoint.Y;
                                        double y2 = L1.EndPoint.Y;
                                        if (Math.Round(y1, 2) == Math.Round(y2, 2))
                                        {
                                            known_y2 = y1;

                                        }


                                    }

                                }

                                if (Ent14 is MText)
                                {
                                    MText M1 = Ent14 as MText;
                                    if (M1 != null)
                                    {
                                        string Continut = M1.Text.Replace("'", "");
                                        if (Functions.IsNumeric(Continut) == true)
                                        {
                                            known_el2 = Convert.ToDouble(Continut);

                                        }


                                    }

                                }

                                if (Ent14 is DBText)
                                {
                                    DBText T1 = Ent14 as DBText;
                                    if (T1 != null)
                                    {
                                        string Continut = T1.TextString.Replace("'", "");
                                        if (Functions.IsNumeric(Continut) == true)
                                        {
                                            known_el2 = Convert.ToDouble(Continut);

                                        }


                                    }

                                }

                                hexag = (known_x2 - known_x1) / (known_sta2 - known_sta1);

                                vexag = (known_y2 - known_y1) / (known_el2 - known_el1);


                                #endregion

                            }


                            #endregion

                        }
                        else
                        {
                            if (Functions.IsNumeric(textBox_hex.Text) == true)
                            {
                                hexag = Convert.ToDouble(textBox_hex.Text);
                            }

                            if (Functions.IsNumeric(textBox_vex.Text) == true)
                            {
                                vexag = Convert.ToDouble(textBox_vex.Text);
                            }


                            #region select grid

                            Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_hor1;
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rezh1 = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rezh1.MessageForAdding = "\nSelect a vertical line (STATION) and the label for it:";
                            Prompt_rezh1.SingleOnly = false;
                            Rezultat_hor1 = ThisDrawing.Editor.GetSelection(Prompt_rezh1);

                            if (Rezultat_hor1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                            {
                                this.MdiParent.WindowState = FormWindowState.Normal;
                                set_enable_true();
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }

                            if (Rezultat_hor1.Value.Count != 2)
                            {
                                MessageBox.Show("the selection has to contain 2 objects. You have selected only" + Rezultat_hor1.Value.Count.ToString() + " object(s).\r\nOperation aborted");
                                this.MdiParent.WindowState = FormWindowState.Normal;
                                set_enable_true();
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }

                            Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_ver1;
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rezv1 = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rezv1.MessageForAdding = "\nSelect a horizontal line (ELEVATION) and the label for it:";
                            Prompt_rezv1.SingleOnly = false;
                            Rezultat_ver1 = ThisDrawing.Editor.GetSelection(Prompt_rezv1);

                            if (Rezultat_ver1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                            {
                                this.MdiParent.WindowState = FormWindowState.Normal;
                                set_enable_true();
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }

                            if (Rezultat_ver1.Value.Count != 2)
                            {
                                MessageBox.Show("the selection has to contain 2 objects. You have selected only" + Rezultat_ver1.Value.Count.ToString() + " object(s).\r\nOperation aborted");
                                this.MdiParent.WindowState = FormWindowState.Normal;
                                set_enable_true();
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }

                            #endregion

                            #region get_parameters for grid



                            Entity Ent1 = Trans1.GetObject(Rezultat_hor1.Value[0].ObjectId, OpenMode.ForRead) as Entity;
                            Entity Ent2 = Trans1.GetObject(Rezultat_hor1.Value[1].ObjectId, OpenMode.ForRead) as Entity;

                            Entity Ent3 = Trans1.GetObject(Rezultat_ver1.Value[0].ObjectId, OpenMode.ForRead) as Entity;
                            Entity Ent4 = Trans1.GetObject(Rezultat_ver1.Value[1].ObjectId, OpenMode.ForRead) as Entity;


                            if (((Ent1 is Polyline || Ent1 is Line) && (Ent2 is MText || Ent2 is DBText)) || ((Ent2 is Polyline || Ent2 is Line) && (Ent1 is MText || Ent1 is DBText)))
                            {
                                #region ent1
                                if (Ent1 is Polyline)
                                {
                                    Polyline P1 = Ent1 as Polyline;
                                    if (P1 != null)
                                    {
                                        double x1 = P1.StartPoint.X;
                                        double x2 = P1.EndPoint.X;
                                        if (Math.Round(x1, 2) == Math.Round(x2, 2))
                                        {
                                            known_x1 = x1;

                                        }


                                    }

                                }

                                if (Ent1 is Line)
                                {
                                    Line L1 = Ent1 as Line;
                                    if (L1 != null)
                                    {
                                        double x1 = L1.StartPoint.X;
                                        double x2 = L1.EndPoint.X;
                                        if (Math.Round(x1, 2) == Math.Round(x2, 2))
                                        {
                                            known_x1 = x1;

                                        }


                                    }

                                }

                                if (Ent1 is MText)
                                {
                                    MText M1 = Ent1 as MText;
                                    if (M1 != null)
                                    {
                                        string Continut = M1.Text.Replace("+", "");
                                        if (Functions.IsNumeric(Continut) == true)
                                        {
                                            known_sta1 = Convert.ToDouble(Continut);

                                        }


                                    }

                                }

                                if (Ent1 is DBText)
                                {
                                    DBText T1 = Ent1 as DBText;
                                    if (T1 != null)
                                    {
                                        string Continut = T1.TextString.Replace("+", "");
                                        if (Functions.IsNumeric(Continut) == true)
                                        {
                                            known_sta1 = Convert.ToDouble(Continut);

                                        }


                                    }

                                }


                                if (Ent2 is Polyline)
                                {
                                    Polyline P1 = Ent2 as Polyline;
                                    if (P1 != null)
                                    {
                                        double x1 = P1.StartPoint.X;
                                        double x2 = P1.EndPoint.X;
                                        if (Math.Round(x1, 2) == Math.Round(x2, 2))
                                        {
                                            known_x1 = x1;

                                        }


                                    }

                                }

                                if (Ent2 is Line)
                                {
                                    Line L1 = Ent2 as Line;
                                    if (L1 != null)
                                    {
                                        double x1 = L1.StartPoint.X;
                                        double x2 = L1.EndPoint.X;
                                        if (Math.Round(x1, 2) == Math.Round(x2, 2))
                                        {
                                            known_x1 = x1;

                                        }


                                    }

                                }

                                if (Ent2 is MText)
                                {
                                    MText M1 = Ent2 as MText;
                                    if (M1 != null)
                                    {
                                        string Continut = M1.Text.Replace("+", "");
                                        if (Functions.IsNumeric(Continut) == true)
                                        {
                                            known_sta1 = Convert.ToDouble(Continut);

                                        }


                                    }

                                }

                                if (Ent2 is DBText)
                                {
                                    DBText T1 = Ent2 as DBText;
                                    if (T1 != null)
                                    {
                                        string Continut = T1.TextString.Replace("+", "");
                                        if (Functions.IsNumeric(Continut) == true)
                                        {
                                            known_sta1 = Convert.ToDouble(Continut);

                                        }


                                    }

                                }
                                #endregion

                            }

                            if (((Ent3 is Polyline || Ent3 is Line) & (Ent4 is MText || Ent4 is DBText)) || ((Ent4 is Polyline || Ent4 is Line) & (Ent3 is MText || Ent3 is DBText)))
                            {
                                #region ent3

                                if (Ent3 is Polyline)
                                {
                                    Polyline P1 = Ent3 as Polyline;
                                    if (P1 != null)
                                    {
                                        double y1 = P1.StartPoint.Y;
                                        double y2 = P1.EndPoint.Y;
                                        if (Math.Round(y1, 2) == Math.Round(y2, 2))
                                        {
                                            known_y1 = y1;

                                        }


                                    }

                                }

                                if (Ent3 is Line)
                                {
                                    Line L1 = Ent3 as Line;
                                    if (L1 != null)
                                    {
                                        double y1 = L1.StartPoint.Y;
                                        double y2 = L1.EndPoint.Y;
                                        if (Math.Round(y1, 2) == Math.Round(y2, 2))
                                        {
                                            known_y1 = y1;

                                        }


                                    }

                                }

                                if (Ent3 is MText)
                                {
                                    MText M1 = Ent3 as MText;
                                    if (M1 != null)
                                    {
                                        string Continut = M1.Text.Replace("'", "");
                                        if (Functions.IsNumeric(Continut) == true)
                                        {
                                            known_el1 = Convert.ToDouble(Continut);

                                        }


                                    }

                                }

                                if (Ent3 is DBText)
                                {
                                    DBText T1 = Ent3 as DBText;
                                    if (T1 != null)
                                    {
                                        string Continut = T1.TextString.Replace("'", "");
                                        if (Functions.IsNumeric(Continut) == true)
                                        {
                                            known_el1 = Convert.ToDouble(Continut);

                                        }


                                    }

                                }


                                if (Ent4 is Polyline)
                                {
                                    Polyline P1 = Ent4 as Polyline;
                                    if (P1 != null)
                                    {
                                        double y1 = P1.StartPoint.Y;
                                        double y2 = P1.EndPoint.Y;
                                        if (Math.Round(y1, 2) == Math.Round(y2, 2))
                                        {
                                            known_y1 = y1;

                                        }


                                    }

                                }

                                if (Ent4 is Line)
                                {
                                    Line L1 = Ent4 as Line;
                                    if (L1 != null)
                                    {
                                        double y1 = L1.StartPoint.Y;
                                        double y2 = L1.EndPoint.Y;
                                        if (Math.Round(y1, 2) == Math.Round(y2, 2))
                                        {
                                            known_y1 = y1;

                                        }


                                    }

                                }

                                if (Ent4 is MText)
                                {
                                    MText M1 = Ent4 as MText;
                                    if (M1 != null)
                                    {
                                        string Continut = M1.Text.Replace("'", "");
                                        if (Functions.IsNumeric(Continut) == true)
                                        {
                                            known_el1 = Convert.ToDouble(Continut);

                                        }


                                    }

                                }

                                if (Ent4 is DBText)
                                {
                                    DBText T1 = Ent4 as DBText;
                                    if (T1 != null)
                                    {
                                        string Continut = T1.TextString.Replace("'", "");
                                        if (Functions.IsNumeric(Continut) == true)
                                        {
                                            known_el1 = Convert.ToDouble(Continut);

                                        }


                                    }

                                }
                                #endregion
                            }


                            #endregion


                        }

                        #region select ground polyline

                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_poly;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_poly = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_poly.MessageForAdding = "\nselect profile polyline:";
                        Prompt_poly.SingleOnly = true;
                        Rezultat_poly = ThisDrawing.Editor.GetSelection(Prompt_poly);

                        if (Rezultat_poly.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            this.MdiParent.WindowState = FormWindowState.Normal;
                            set_enable_true();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        #endregion

                        #region get polyline of profile
                        Entity Ent0 = Trans1.GetObject(Rezultat_poly.Value[0].ObjectId, OpenMode.ForRead) as Entity;
                        if ((Ent0 is Polyline) == false)
                        {
                            MessageBox.Show("the polyline profile is not a polyline\r\n" + Ent0.GetType().ToString() + "\r\nOperation aborted");
                            this.MdiParent.WindowState = FormWindowState.Normal;
                            set_enable_true();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }
                        Poly_Graph_exag = Ent0 as Polyline;
                        #endregion




                        if (Poly_Graph_exag != null && known_x1 != -123.1234567 && known_y1 != -123.1234567 && known_sta1 != -123.1234567 && known_el1 != -123.1234567)
                        {


                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                            Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                            PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease Specify the HDD Entry Point");
                            PP1.AllowNone = false;
                            Point_res1 = Editor1.GetPoint(PP1);

                            if (Point_res1.Status != PromptStatus.OK)
                            {
                                set_enable_true();
                                this.MdiParent.WindowState = FormWindowState.Normal;

                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }

                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                            Autodesk.AutoCAD.EditorInput.PromptPointOptions PP2;
                            PP2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease Specify a point towards the HDD Exit Point");
                            PP2.AllowNone = false;
                            PP2.UseBasePoint = true;
                            PP2.BasePoint = Point_res1.Value;

                            Point_res2 = Editor1.GetPoint(PP2);

                            if (Point_res2.Status != PromptStatus.OK)
                            {
                                set_enable_true();
                                this.MdiParent.WindowState = FormWindowState.Normal;

                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }


                            pt_1 = Poly_Graph_exag.GetClosestPointTo(Point_res1.Value.TransformBy(CurentUCSmatrix), Vector3d.ZAxis, false);

                            point_referinta = pt_1;

                            Point3d pt_dir = Poly_Graph_exag.GetClosestPointTo(Point_res2.Value.TransformBy(CurentUCSmatrix), Vector3d.ZAxis, false);



                            if (pt_1.X > pt_dir.X)
                            {
                                lr = -1;
                              
                            }
                            else
                            {
                                lr = 1;
                                
                            }







                            Poly_Graph = new Polyline();
                            for (int i = 0; i < Poly_Graph_exag.NumberOfVertices; ++i)
                            {
                                Poly_Graph.AddVertexAt(i, new Point2d(pt_1.X + (Poly_Graph_exag.GetPoint2dAt(i).X - pt_1.X) / hexag, pt_1.Y + (Poly_Graph_exag.GetPoint2dAt(i).Y - pt_1.Y) / vexag), 0, 0, 0);
                            }
                            Point2d pt_last = Poly_Graph.GetPoint2dAt(Poly_Graph_exag.NumberOfVertices - 1);

                            if (lr == 1) Poly_Graph.AddVertexAt(Poly_Graph.NumberOfVertices, new Point2d(pt_last.X + 1000, pt_last.Y), 0, 0, 0);

                            sta_original = known_sta1 + (pt_1.X - known_x1) / hexag;

                            elev1 = known_el1 + (pt_1.Y - known_y1) / vexag;

                            eq_sta1 = Station_equation_of(sta_original, dt_steq);

                            textBox_sta1.Text = Functions.Get_chainage_from_double(eq_sta1, "f", 4);
                            textBox_elev1.Text = Functions.Get_String_Rounded(elev1, 4);
                            textBoxA.Text = Functions.Get_String_Rounded(elev1, 4);
                            reset_variables();
                            reset_form();
                        }
                        else
                        {
                            MessageBox.Show("you did not selected the proper entities");
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
        }
        private double Station_equation_of(double Station_measured, System.Data.DataTable Data_table_station_equation)
        {
            double Total_length = 0;
            double Station_ahead_p = 0;
            bool increase_p = true;
            double Total_length_p = 0;

            if (Data_table_station_equation != null)
            {
                if (Data_table_station_equation.Rows.Count > 0)
                {
                    for (int i = 0; i < Data_table_station_equation.Rows.Count; ++i)
                    {
                        if (Data_table_station_equation.Rows[i]["Back"] != DBNull.Value && Data_table_station_equation.Rows[i]["Start"] != DBNull.Value && Data_table_station_equation.Rows[i]["Ahead"] != DBNull.Value && Data_table_station_equation.Rows[i]["Increase"] != DBNull.Value)
                        {
                            double Station_back = Convert.ToDouble(Data_table_station_equation.Rows[i]["Back"]);
                            double Start1 = Convert.ToDouble(Data_table_station_equation.Rows[i]["Start"]);
                            double Station_ahead = Convert.ToDouble(Data_table_station_equation.Rows[i]["Ahead"]);
                            bool increase1 = (bool)Data_table_station_equation.Rows[i]["Increase"];
                            if (i == 0)
                            {
                                if (increase1 == true)
                                {
                                    Total_length = Station_back - Start1;
                                }
                                else
                                {
                                    Total_length = Start1 - Station_back;
                                }

                                Station_ahead_p = Station_ahead;
                            }
                            else
                            {
                                if (increase_p == true)
                                {
                                    Total_length = Total_length + (Station_back - Station_ahead_p);
                                }
                                else
                                {
                                    Total_length = Total_length - (Station_back - Station_ahead_p);
                                }
                            }


                            if (Station_measured - Start1 < Total_length)
                            {
                                if (increase_p == true)
                                {
                                    return Station_ahead_p + (Station_measured - Start1 - Total_length_p);
                                }
                                else
                                {
                                    return Station_ahead_p - (Station_measured - Start1 - Total_length_p);
                                }
                            }

                            if (Station_measured - Start1 == Total_length)
                            {
                                return Station_ahead;
                            }

                            if (Station_measured - Start1 > Total_length)
                            {
                                if (i == Data_table_station_equation.Rows.Count - 1)
                                {
                                    if (increase1 == true)
                                    {
                                        return Station_ahead + (Station_measured - Start1 - Total_length);
                                    }
                                    else
                                    {
                                        return Station_ahead - (Station_measured - Start1 - Total_length);
                                    }
                                }
                            }
                            increase_p = increase1;
                            Total_length_p = Total_length;
                        }
                    }
                }
            }
            return Station_measured;
        }

        private void textBox_pozitive_KeyPress(object sender, KeyPressEventArgs e)
        {
            Functions.textbox_input_only_pozitive_doubles_at_keypress(sender, e);
        }
        private void textBox_negative_KeyPress(object sender, KeyPressEventArgs e)
        {
            Functions.textbox_input_only_pozitive_and_negative_doubles_at_keypress(sender, e);
        }
        private void textBox_dev_angle1_TextChanged(object sender, EventArgs e)
        {
            textBox_dev_angle2.Text = textBox_dev_angle1.Text;
        }
        private void textBox_elev1_TextChanged(object sender, EventArgs e)
        {
            textBoxA.Text = textBox_elev1.Text;
        }


        private void Create_HDD_object_data()
        {

            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {

                        List<string> List1 = new List<string>();
                        List<string> List2 = new List<string>();
                        List<Autodesk.Gis.Map.Constants.DataType> List3 = new List<Autodesk.Gis.Map.Constants.DataType>();

                        List1.Add("Elevation1");
                        List2.Add("Elevation");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                        List1.Add("Sta1");
                        List2.Add("Station");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                        List1.Add("Angle_start");
                        List2.Add("Start angle");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                        List1.Add("L1");
                        List2.Add("L1");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                        List1.Add("Radius1");
                        List2.Add("Radius1");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                        List1.Add("L3");
                        List2.Add("L3");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                        List1.Add("H_Angle");
                        List2.Add("Horizontal angle");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                        List1.Add("Radius2");
                        List2.Add("Radius2");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                        List1.Add("Angle_end");
                        List2.Add("End angle");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                        List1.Add("Notes");
                        List2.Add("Notes");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        Functions.Get_object_data_table(OD_table_name, "Generated by HDD tool", List1, List2, List3);

                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button_HDD_Calculate_parameters_Click(object sender, EventArgs e)
        {
            set_enable_false();
            reset_variables();

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            try
            {



                if (dt_steq != null && dt_steq.Rows.Count > 0)
                {
                    if (Functions.IsNumeric(textBox_sta1.Text.Replace("+", "")) == true)
                    {
                        sta1 = sta_original + Convert.ToDouble(textBox_sta1.Text.Replace("+", "")) - eq_sta1;
                    }
                }

                else
                {
                    if (Functions.IsNumeric(textBox_sta1.Text.Replace("+", "")) == true)
                    {
                        sta1 = Convert.ToDouble(textBox_sta1.Text.Replace("+", ""));
                    }
                }

                if (Functions.IsNumeric(textBox_hd1.Text) == true)
                {
                    hd1 = Convert.ToDouble(textBox_hd1.Text);
                }
                if (Functions.IsNumeric(textBox_angle_start.Text) == true)
                {
                    Angle1 = Convert.ToDouble(textBox_angle_start.Text);
                }
                if (Functions.IsNumeric(textBox_L1.Text) == true)
                {
                    L1 = Convert.ToDouble(textBox_L1.Text);
                }
                if (Functions.IsNumeric(textBox_radius1.Text) == true)
                {
                    radius1 = Convert.ToDouble(textBox_radius1.Text);
                }
                if (Functions.IsNumeric(textBox_dev_angle1.Text) == true)
                {
                    hda = Convert.ToDouble(textBox_dev_angle1.Text);
                }
                if (Functions.IsNumeric(textBox_L3.Text) == true)
                {
                    L3 = Convert.ToDouble(textBox_L3.Text);
                }

                if (Functions.IsNumeric(textBox_angle_end.Text) == true)
                {
                    Angle2 = Convert.ToDouble(textBox_angle_end.Text);
                }

                if (Functions.IsNumeric(textBox_radius2.Text) == true)
                {
                    radius2 = Convert.ToDouble(textBox_radius2.Text);
                }

                if (Poly_Graph != null)
                {
                    Polyline Poly_Graph_exag = new Polyline();
                    for (int i = 0; i < Poly_Graph.NumberOfVertices; ++i)
                    {
                        Poly_Graph_exag.AddVertexAt(i, new Point2d(pt_1.X + (Poly_Graph.GetPoint2dAt(i).X - pt_1.X) * hexag, pt_1.Y + (Poly_Graph.GetPoint2dAt(i).Y - pt_1.Y) * vexag), 0, 0, 0);
                    }

                    Poly_Graph_exag.Elevation = Poly_Graph.Elevation;


                    Xline xline1 = new Xline();
                    xline1.BasePoint = new Point3d(known_x1 +  (sta1 - known_sta1) * hexag, 0, Poly_Graph.Elevation);   // MINUS IS TAKEN FROM STA1-KNOWN STA
                    xline1.SecondPoint = new Point3d(known_x1 +  (sta1 - known_sta1) * hexag, 10, Poly_Graph.Elevation);
                    Point3dCollection col1 = Functions.Intersect_on_both_operands(Poly_Graph, xline1);
                    Point3dCollection col2 = Functions.Intersect_on_both_operands(Poly_Graph_exag, xline1);

                   


                    if (col1.Count > 0 && col2.Count > 0)
                    {

                        using (DocumentLock lock1 = ThisDrawing.LockDocument())
                        {
                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                            {
                                BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                                BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                                Poly_Graph.TransformBy(Matrix3d.Displacement(col1[0].GetVectorTo(col2[0])));

                                //xline1.ColorIndex = 1;
                                //BTrecord.AppendEntity(xline1);
                                //Trans1.AddNewlyCreatedDBObject(xline1, true);

                                //Poly_Graph_exag.ColorIndex = 1;
                                //BTrecord.AppendEntity(Poly_Graph_exag);
                                //Trans1.AddNewlyCreatedDBObject(Poly_Graph_exag, true);

                                double yk = col2[0].Y;
                                elev1 = known_el1 + (yk - known_y1) / vexag;
                                pt_1 = col2[0]; //ok-display-dwg

                                if (L1 > 0)
                                {
                                    h1 = L1 * Math.Tan(Angle1 * Math.PI / 180);
                                    elev2 = elev1 - h1;
                                    slope2 = L1 / Math.Cos(Angle1 * Math.PI / 180);
                                    hd2 = L1;
                                    sta2 = lr * hd2 + sta1;

                                    Line linie1 = new Line(pt_1, new Point3d(pt_1.X + lr * L1, pt_1.Y - h1, Poly_Graph.Elevation));
                                    pt_2 = linie1.EndPoint;

                                    if (radius1 > 0)
                                    {
                                        arc_len1 = radius1 * (Angle1 + hda) * Math.PI / 180;//ok-display
                                        L2 = radius1 * Math.Sin(Angle1 * Math.PI / 180);//ok-display
                                        h2 = radius1 - radius1 * Math.Cos(Angle1 * Math.PI / 180);//ok-display
                                        linie1 = new Line(pt_2, new Point3d(pt_2.X - lr * radius1, pt_2.Y, 0));//no exag pt ca aici intervine hor and ver exag
                                        linie1.TransformBy(Matrix3d.Rotation(-lr * (Angle1 + 90) * Math.PI / 180, Vector3d.ZAxis, linie1.StartPoint));
                                        linie1.TransformBy(Matrix3d.Rotation(lr * (Angle1 + hda) * Math.PI / 180, Vector3d.ZAxis, linie1.EndPoint));
                                        pt_3 = linie1.StartPoint;
                                        elev3 = elev2 - (pt_2.Y - pt_3.Y);
                                        slope3 = slope2 + arc_len1;
                                        hd3 = lr * (pt_3.X - pt_1.X);
                                        sta3 = lr * hd3 + sta1;

                                        linie1 = new Line(pt_3, new Point3d(pt_3.X + lr * L3, pt_3.Y, 0));
                                        linie1.TransformBy(Matrix3d.Rotation(lr * hda * Math.PI / 180, Vector3d.ZAxis, linie1.StartPoint));
                                        pt_4 = linie1.EndPoint;
                                        h4a = pt_4.Y - linie1.StartPoint.Y;
                                        elev4 = elev1 - (pt_1.Y - pt_4.Y);
                                        slope4 = slope3 + L3;
                                        hd4 = lr * (pt_4.X - pt_1.X);
                                        sta4 = lr * hd4 + sta1;

                                        if (radius2 > 0)
                                        {
                                            linie1 = new Line(pt_4, new Point3d(pt_4.X - lr * radius2, pt_4.Y, 0));

                                            linie1.TransformBy(Matrix3d.Rotation(-lr * (90 - hda) * Math.PI / 180, Vector3d.ZAxis, linie1.StartPoint));
                                            linie1.TransformBy(Matrix3d.Rotation(lr * Angle2 * Math.PI / 180, Vector3d.ZAxis, linie1.EndPoint));
                                            pt_5 = linie1.StartPoint;
                                            h4 = pt_5.Y - pt_4.Y;
                                            L4 = lr * (pt_5.X - pt_4.X);
                                            arc_len2 = radius2 * Angle2 * Math.PI / 180;
                                            elev5 = elev4 + h4;
                                            slope5 = slope4 + arc_len2;

                                            sta5 = lr * hd5 + sta1;

                                            hd5 = lr * (pt_5.X - pt_1.X);
                                            linie1 = new Line(new Point3d(pt_5.X, pt_5.Y, Poly_Graph.Elevation), new Point3d(pt_5.X, pt_5.Y + 1000000000, Poly_Graph.Elevation));
                                            col1 = new Point3dCollection();
                                            linie1.IntersectWith(Poly_Graph, Intersect.OnBothOperands, col1, IntPtr.Zero, IntPtr.Zero);
                                            if (col1.Count > 0)
                                            {
                                                Point3d pt_5G = col1[0];
                                                double ground_el_at_5 = elev1 - (pt_1.Y - pt_5G.Y);

                                                if (ground_el_at_5 > elev5)
                                                {
                                                    linie1 = new Line(new Point3d(pt_5.X, pt_5.Y, Poly_Graph.Elevation), new Point3d(pt_5.X + lr * 10, pt_5.Y, Poly_Graph.Elevation));
                                                    linie1.TransformBy(Matrix3d.Rotation(lr * Angle2 * Math.PI / 180, Vector3d.ZAxis, linie1.StartPoint));

                                                    col1 = new Point3dCollection();
                                                    linie1.IntersectWith(Poly_Graph, Intersect.ExtendThis, col1, IntPtr.Zero, IntPtr.Zero);
                                                    if (col1.Count > 0)
                                                    {
                                                        pt_6 = col1[0];
                                                        hd6 = lr * (pt_6.X - pt_1.X);
                                                        h5 = (pt_6.Y - pt_5.Y);
                                                        L5 = lr * (pt_6.X - pt_5.X);
                                                        elev6 = elev5 + h5;
                                                        sta6 = sta1 + lr * hd6;
                                                        slope6 = slope5 + Math.Pow(Math.Pow(L5, 2) + Math.Pow(h5, 2), 0.5);

                                                        Point3d ppoly1 = Poly_Graph.GetClosestPointTo(pt_1, Vector3d.ZAxis, false);
                                                        double param1 = Poly_Graph.GetParameterAtPoint(ppoly1);
                                                        Point3d ppoly6 = Poly_Graph.GetClosestPointTo(pt_6, Vector3d.ZAxis, false);
                                                        double param6 = Poly_Graph.GetParameterAtPoint(ppoly6);

                                                        if (param1 > param6)
                                                        {
                                                            double t = param1;
                                                            param1 = param6;
                                                            param6 = t;
                                                        }

                                                        if (param1 <= param6)
                                                        {
                                                            double y_min = ppoly1.Y;
                                                            double x_low = ppoly1.X;
                                                            for (int i = Convert.ToInt32(Math.Ceiling(param1)); i < param6; ++i)
                                                            {
                                                                Point3d vertex1 = Poly_Graph.GetPointAtParameter(i);
                                                                if (vertex1.Y < y_min)
                                                                {
                                                                    y_min = vertex1.Y;
                                                                    x_low = vertex1.X;
                                                                }
                                                            }

                                                            elevB = elev1 - (pt_1.Y - y_min);
                                                            staC = sta1 - lr * (pt_1.X - x_low);
                                                            difD = elev1 - elevB;

                                                            Polyline poly_hdd = new Polyline();
                                                            poly_hdd.AddVertexAt(0, new Point2d(pt_1.X, pt_1.Y), 0, 0, 0);
                                                            poly_hdd.AddVertexAt(1, new Point2d(pt_2.X, pt_2.Y), lr * Math.Tan(0.25 * (Angle1 + hda) * Math.PI / 180), 0, 0);
                                                            poly_hdd.AddVertexAt(2, new Point2d(pt_3.X, pt_3.Y), 0, 0, 0);
                                                            poly_hdd.AddVertexAt(3, new Point2d(pt_4.X, pt_4.Y), lr * Math.Tan(0.25 * Angle2 * Math.PI / 180), 0, 0);
                                                            poly_hdd.AddVertexAt(4, new Point2d(pt_5.X, pt_5.Y), 0, 0, 0);
                                                            poly_hdd.AddVertexAt(5, new Point2d(pt_6.X, pt_6.Y), 0, 0, 0);

                                                            linie1 = new Line(new Point3d(x_low, -1, Poly_Graph.Elevation), new Point3d(x_low, 1, Poly_Graph.Elevation));
                                                            col1 = new Point3dCollection();
                                                            linie1.IntersectWith(poly_hdd, Intersect.ExtendThis, col1, IntPtr.Zero, IntPtr.Zero);
                                                            if (col1.Count > 0)
                                                            {
                                                                difE = y_min - col1[0].Y;
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
                    else
                    {
                        MessageBox.Show("the station you specified is not on the graph polyline");
                    }
                }

                textBox_elev1.Text = Functions.Get_String_Rounded(elev1, round1);
                textBox_h1.Text = Functions.Get_String_Rounded(h1, round1);
                textBox_arc_length1.Text = Functions.Get_String_Rounded(arc_len1, round1);

                textBox_elev2.Text = Functions.Get_String_Rounded(elev2, round1);
                textBox_slope2.Text = Functions.Get_String_Rounded(slope2, round1);


                double eq_sta2 = Station_equation_of(sta2, dt_steq);
                textBox_sta2.Text = Functions.Get_chainage_from_double(eq_sta2, "f", round1);
                textBox_hd2.Text = Functions.Get_String_Rounded(hd2, round1);
                textBox_L2.Text = Functions.Get_String_Rounded(L2, round1);
                textBox_h2.Text = Functions.Get_String_Rounded(h2, round1);
                textBox_arc_length2.Text = Functions.Get_String_Rounded(arc_len2, round1);

                textBox_elev3.Text = Functions.Get_String_Rounded(elev3, round1);
                textBox_slope3.Text = Functions.Get_String_Rounded(slope3, round1);

                double eq_sta3 = Station_equation_of(sta3, dt_steq);
                textBox_sta3.Text = Functions.Get_chainage_from_double(eq_sta3, "f", round1);
                textBox_hd3.Text = Functions.Get_String_Rounded(hd3, round1);

                textBox_elev4.Text = Functions.Get_String_Rounded(elev4, round1);
                textBox_h4.Text = Functions.Get_String_Rounded(h4, round1);
                textBox_L4.Text = Functions.Get_String_Rounded(L4, round1);
                textBox_slope4.Text = Functions.Get_String_Rounded(slope4, round1);

                double eq_sta4 = Station_equation_of(sta4, dt_steq);
                textBox_sta4.Text = Functions.Get_chainage_from_double(eq_sta4, "f", round1);
                textBox_hd4.Text = Functions.Get_String_Rounded(hd4, round1);
                textBox_h4a.Text = Functions.Get_String_Rounded(h4a, round1);

                textBox_elev5.Text = Functions.Get_String_Rounded(elev5, round1);
                textBox_slope5.Text = Functions.Get_String_Rounded(slope5, round1);

                double eq_sta5 = Station_equation_of(sta5, dt_steq);
                textBox_sta5.Text = Functions.Get_chainage_from_double(eq_sta5, "f", round1);
                textBox_hd5.Text = Functions.Get_String_Rounded(hd5, round1);
                textBox_h5.Text = Functions.Get_String_Rounded(h5, round1);
                textBox_L5.Text = Functions.Get_String_Rounded(L5, round1);

                textBox_hd6.Text = Functions.Get_String_Rounded(hd6, round1);
                textBox_elev6.Text = Functions.Get_String_Rounded(elev6, round1);
                double eq_sta6 = Station_equation_of(sta6, dt_steq);
                textBox_sta6.Text = Functions.Get_chainage_from_double(eq_sta6, "f", 4);
                textBox_slope6.Text = Functions.Get_String_Rounded(slope6, round1);



                textBoxB.Text = Functions.Get_String_Rounded(elevB, round1);
                textBoxC.Text = Functions.Get_chainage_from_double(staC, "f", round1);
                textBoxD.Text = Functions.Get_String_Rounded(difD, round1);
                textBoxE.Text = Functions.Get_String_Rounded(difE, round1);

                OD_lista_val = new List<object>();
                OD_lista_type = new List<Autodesk.Gis.Map.Constants.DataType>();

                OD_lista_val.Add(elev1);
                OD_lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                OD_lista_val.Add(Convert.ToDouble(textBox_sta1.Text.Replace("+", "")));
                OD_lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                OD_lista_val.Add(Angle1);
                OD_lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                OD_lista_val.Add(L1);
                OD_lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                OD_lista_val.Add(radius1);
                OD_lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                OD_lista_val.Add(L3);
                OD_lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                OD_lista_val.Add(hda);
                OD_lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                OD_lista_val.Add(radius2);
                OD_lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                OD_lista_val.Add(Angle2);
                OD_lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                OD_lista_val.Add(Environment.UserName.ToUpper() + " at:" + DateTime.Now.Month + "/" + DateTime.Now.Day + "/" + DateTime.Now.Year);
                OD_lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            set_enable_true();
        }


        private void button_draw_HDD_Click(object sender, EventArgs e)
        {
            if (pt_5.X == pt_6.X && pt_5.Y == pt_6.Y) return;

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
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

                        Create_HDD_object_data();

                        string hdd_design = "HDD_Design";
                        Functions.Creaza_layer(hdd_design, 1, true);






                        Polyline poly_hdd = new Polyline();
                        poly_hdd.AddVertexAt(0, new Point2d(pt_1.X, pt_1.Y), 0, 0, 0);
                        poly_hdd.AddVertexAt(1, new Point2d(pt_2.X, pt_2.Y), lr * Math.Tan(0.25 * (Angle1 + hda) * Math.PI / 180), 0, 0);
                        poly_hdd.AddVertexAt(2, new Point2d(pt_3.X, pt_3.Y), 0, 0, 0);
                        poly_hdd.AddVertexAt(3, new Point2d(pt_4.X, pt_4.Y), lr * Math.Tan(0.25 * Angle2 * Math.PI / 180), 0, 0);
                        poly_hdd.AddVertexAt(4, new Point2d(pt_5.X, pt_5.Y), 0, 0, 0);
                        poly_hdd.AddVertexAt(5, new Point2d(pt_6.X, pt_6.Y), 0, 0, 0);


                        Point3d ptm1 = poly_hdd.GetPointAtDist((slope2 + slope3) / 2);
                        Point3d ptm2 = poly_hdd.GetPointAtDist((slope4 + slope5) / 2);



                        double x1 = pt_1.X;
                        double x2 = pt_1.X - (pt_1.X - pt_2.X) * hexag;
                        double x3 = pt_1.X - (pt_1.X - pt_3.X) * hexag;
                        double x4 = pt_1.X - (pt_1.X - pt_4.X) * hexag;
                        double x5 = pt_1.X - (pt_1.X - pt_5.X) * hexag;
                        double x6 = pt_1.X - (pt_1.X - pt_6.X) * hexag;

                        double x7 = pt_1.X - (pt_1.X - ptm1.X) * hexag;
                        double x8 = pt_1.X - (pt_1.X - ptm2.X) * hexag;


                        double y1 = pt_1.Y;
                        double y2 = pt_1.Y - (pt_1.Y - pt_2.Y) * vexag;
                        double y3 = pt_1.Y - (pt_1.Y - pt_3.Y) * vexag;
                        double y4 = pt_1.Y - (pt_1.Y - pt_4.Y) * vexag;
                        double y5 = pt_1.Y - (pt_1.Y - pt_5.Y) * vexag;
                        double y6 = pt_1.Y - (pt_1.Y - pt_6.Y) * vexag;

                        double y7 = pt_1.Y - (pt_1.Y - ptm1.Y) * vexag;
                        double y8 = pt_1.Y - (pt_1.Y - ptm2.Y) * vexag;


                        if (hexag == 1 && vexag == 1)
                        {
                            poly_hdd.ColorIndex = cid;
                            poly_hdd.Layer = hdd_design;
                            BTrecord.AppendEntity(poly_hdd);
                            Trans1.AddNewlyCreatedDBObject(poly_hdd, true);

                            Functions.Populate_object_data_table_from_objectid(Tables1, poly_hdd.ObjectId, OD_table_name, OD_lista_val, OD_lista_type);

                        }

                        else
                        {


                            BlockTable1.UpgradeOpen();

                            int idx = 1;
                            bool exista = true;
                            do
                            {
                                if (BlockTable1.Has("HDD" + idx.ToString()) == false)
                                {
                                    using (BlockTableRecord bltrec1 = new BlockTableRecord())
                                    {

                                        bltrec1.Name = "HDD" + idx.ToString();
                                        poly_hdd.Layer = "0";
                                        poly_hdd.ColorIndex = 0;


                                        poly_hdd.TransformBy(Matrix3d.Displacement(pt_1.GetVectorTo(new Point3d(0, 0, 0))));

                                        bltrec1.AppendEntity(poly_hdd);
                                        BlockTable1.Add(bltrec1);
                                        Trans1.AddNewlyCreatedDBObject(bltrec1, true);
                                    }
                                    exista = false;
                                }
                                else
                                {
                                    ++idx;
                                }
                            } while (exista == true);

                            BlockReference block1 = Functions.InsertBlock_with_2scales(ThisDrawing.Database, BTrecord, "HDD" + idx.ToString(), pt_1, hexag, vexag, 0, hdd_design);
                            block1.ColorIndex = cid;
                            Functions.Populate_object_data_table_from_objectid(Tables1, block1.ObjectId, OD_table_name, OD_lista_val, OD_lista_type);

                            if (cid == 7)
                            {
                                cid = 1;
                            }
                            else
                            {
                                ++cid;
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

            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            set_enable_true();


        }

        private void Button_label_hdd_Click(object sender, EventArgs e)
        {
            if (pt_5.X == pt_6.X && pt_5.Y == pt_6.Y) return;

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
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

                        Create_HDD_object_data();

                        string hdd_design = "HDD_Design";
                        Functions.Creaza_layer(hdd_design, 1, true);

                        Polyline poly_hdd = new Polyline();
                        poly_hdd.AddVertexAt(0, new Point2d(pt_1.X, pt_1.Y), 0, 0, 0);
                        poly_hdd.AddVertexAt(1, new Point2d(pt_2.X, pt_2.Y), lr * Math.Tan(0.25 * (Angle1 + hda) * Math.PI / 180), 0, 0);
                        poly_hdd.AddVertexAt(2, new Point2d(pt_3.X, pt_3.Y), 0, 0, 0);
                        poly_hdd.AddVertexAt(3, new Point2d(pt_4.X, pt_4.Y), lr * Math.Tan(0.25 * Angle2 * Math.PI / 180), 0, 0);
                        poly_hdd.AddVertexAt(4, new Point2d(pt_5.X, pt_5.Y), 0, 0, 0);
                        poly_hdd.AddVertexAt(5, new Point2d(pt_6.X, pt_6.Y), 0, 0, 0);

                        Point3d ptm1 = poly_hdd.GetPointAtDist((slope2 + slope3) / 2);
                        Point3d ptm2 = poly_hdd.GetPointAtDist((slope4 + slope5) / 2);

                        double x1 = pt_1.X;
                        double x2 = pt_1.X - (pt_1.X - pt_2.X) * hexag;
                        double x3 = pt_1.X - (pt_1.X - pt_3.X) * hexag;
                        double x4 = pt_1.X - (pt_1.X - pt_4.X) * hexag;
                        double x5 = pt_1.X - (pt_1.X - pt_5.X) * hexag;
                        double x6 = pt_1.X - (pt_1.X - pt_6.X) * hexag;

                        double x7 = pt_1.X - (pt_1.X - ptm1.X) * hexag;
                        double x8 = pt_1.X - (pt_1.X - ptm2.X) * hexag;

                        double y1 = pt_1.Y;
                        double y2 = pt_1.Y - (pt_1.Y - pt_2.Y) * vexag;
                        double y3 = pt_1.Y - (pt_1.Y - pt_3.Y) * vexag;
                        double y4 = pt_1.Y - (pt_1.Y - pt_4.Y) * vexag;
                        double y5 = pt_1.Y - (pt_1.Y - pt_5.Y) * vexag;
                        double y6 = pt_1.Y - (pt_1.Y - pt_6.Y) * vexag;

                        double y7 = pt_1.Y - (pt_1.Y - ptm1.Y) * vexag;
                        double y8 = pt_1.Y - (pt_1.Y - ptm2.Y) * vexag;

                        double textH = 0.08;

                        if (comboBox_scales.SelectedIndex > 0)
                        {
                            if (comboBox_scales.Text != "")
                            {
                                string txt = comboBox_scales.Text.Replace("1:", "");
                                if (Functions.IsNumeric(txt) == true)
                                {
                                    textH = textH * Convert.ToDouble(txt);
                                }
                            }
                        }

                        DBDictionary leader_style_table = Trans1.GetObject(ThisDrawing.Database.MLeaderStyleDictionaryId, OpenMode.ForRead) as DBDictionary;
                        DimStyleTable dim_style_table = Trans1.GetObject(ThisDrawing.Database.DimStyleTableId, OpenMode.ForWrite) as DimStyleTable;

                        string mleaderstyle_name = "HDD1";
                        string textstyle_name = "HDD1";
                        string dimstyle_name = "HDD1";
                        TextStyleTableRecord HDD_textstyle = null;

                        if (leader_style_table != null)
                        {
                            MLeaderStyle HDD_mleader = new MLeaderStyle();


                            TextStyleTable Text_style_table = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                            foreach (ObjectId TextStyle_id in Text_style_table)
                            {
                                TextStyleTableRecord TextStyle1 = Trans1.GetObject(TextStyle_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite) as TextStyleTableRecord;
                                if (TextStyle1.Name.ToLower() == textstyle_name.ToLower())
                                {
                                    HDD_textstyle = TextStyle1;
                                    HDD_textstyle.TextSize = textH;
                                }
                            }

                            if (HDD_textstyle == null)
                            {
                                Text_style_table.UpgradeOpen();
                                HDD_textstyle = new TextStyleTableRecord();
                                HDD_textstyle.Name = textstyle_name;
                                HDD_textstyle.TextSize = textH;
                                HDD_textstyle.ObliquingAngle = 0;
                                HDD_textstyle.FileName = "romans.shx"; //"arial.ttf" 
                                HDD_textstyle.XScale = 1.0;
                                Text_style_table.Add(HDD_textstyle);
                                Trans1.AddNewlyCreatedDBObject(HDD_textstyle, true);
                            }

                            ObjectId Arrowid = ObjectId.Null;
                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("DIMBLK", "_Dot");
                            if (BlockTable1.Has("_Dot") == true)
                            {
                                Arrowid = BlockTable1["_Dot"];
                            }


                            if (leader_style_table.Contains(mleaderstyle_name) == true)
                            {
                                ObjectId ID1 = leader_style_table.GetAt(mleaderstyle_name);
                                HDD_mleader = Trans1.GetObject(ID1, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite) as MLeaderStyle;
                                HDD_mleader.ArrowSize = textH;
                                HDD_mleader.BreakSize = textH;
                                HDD_mleader.DoglegLength = textH;
                                HDD_mleader.LeaderLineColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 256);
                                HDD_mleader.TextColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 256);
                                HDD_mleader.TextHeight = textH;
                                HDD_mleader.TextStyleId = Text_style_table.ObjectId;
                                HDD_mleader.ArrowSymbolId = Arrowid;
                                HDD_mleader.BlockColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 0);
                                HDD_mleader.BlockRotation = 0;
                                HDD_mleader.BlockScale = new Autodesk.AutoCAD.Geometry.Scale3d(1, 1, 1);
                                HDD_mleader.ContentType = ContentType.MTextContent;
                                HDD_mleader.DrawLeaderOrderType = DrawLeaderOrderType.DrawLeaderHeadFirst;
                                HDD_mleader.DrawMLeaderOrderType = DrawMLeaderOrderType.DrawLeaderFirst;
                                HDD_mleader.EnableBlockRotation = true;
                                HDD_mleader.EnableBlockScale = true;
                                HDD_mleader.EnableDogleg = true;
                                HDD_mleader.EnableFrameText = false;
                                HDD_mleader.EnableLanding = true;
                                HDD_mleader.ExtendLeaderToText = false;
                                HDD_mleader.TextAlignAlwaysLeft = true;
                                HDD_mleader.LandingGap = 0.8 * textH;
                                HDD_mleader.LeaderLineType = LeaderType.StraightLeader;
                                HDD_mleader.LeaderLineWeight = LineWeight.ByBlock;
                                HDD_mleader.MaxLeaderSegmentsPoints = 2;
                                HDD_mleader.Scale = 1;
                                HDD_mleader.TextAlignAlwaysLeft = false;
                                HDD_mleader.TextAlignmentType = TextAlignmentType.LeftAlignment;
                                HDD_mleader.TextAngleType = TextAngleType.HorizontalAngle;
                            }
                            else
                            {
                                leader_style_table.UpgradeOpen();
                                leader_style_table.SetAt(mleaderstyle_name, HDD_mleader);
                                HDD_mleader.ArrowSize = textH;
                                HDD_mleader.BreakSize = textH;
                                HDD_mleader.DoglegLength = textH;
                                HDD_mleader.LeaderLineColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 256);
                                HDD_mleader.TextColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 256);
                                HDD_mleader.TextHeight = textH;
                                HDD_mleader.TextStyleId = Text_style_table.ObjectId;
                                HDD_mleader.ArrowSymbolId = Arrowid;
                                HDD_mleader.BlockColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 0);
                                HDD_mleader.BlockRotation = 0;
                                HDD_mleader.BlockScale = new Autodesk.AutoCAD.Geometry.Scale3d(1, 1, 1);
                                HDD_mleader.ContentType = ContentType.MTextContent;
                                HDD_mleader.DrawLeaderOrderType = DrawLeaderOrderType.DrawLeaderHeadFirst;
                                HDD_mleader.DrawMLeaderOrderType = DrawMLeaderOrderType.DrawLeaderFirst;
                                HDD_mleader.EnableBlockRotation = true;
                                HDD_mleader.EnableBlockScale = true;
                                HDD_mleader.EnableDogleg = true;
                                HDD_mleader.EnableFrameText = false;
                                HDD_mleader.EnableLanding = true;
                                HDD_mleader.ExtendLeaderToText = false;
                                HDD_mleader.TextAlignAlwaysLeft = true;
                                HDD_mleader.LandingGap = 0.8 * textH;
                                HDD_mleader.LeaderLineType = LeaderType.StraightLeader;
                                HDD_mleader.LeaderLineWeight = LineWeight.ByBlock;
                                HDD_mleader.MaxLeaderSegmentsPoints = 2;
                                HDD_mleader.Scale = 1;
                                HDD_mleader.TextAlignAlwaysLeft = false;
                                HDD_mleader.TextAlignmentType = TextAlignmentType.LeftAlignment;
                                HDD_mleader.TextAngleType = TextAngleType.HorizontalAngle;
                                Trans1.AddNewlyCreatedDBObject(HDD_mleader, true);
                            }

                            string content1 = "ENTRY POINT @ " + Functions.Get_String_Rounded(Angle1, 0) + "%%D\r\nSTA: " + Functions.Get_chainage_from_double(Station_equation_of(sta1, dt_steq), "f", 0) + "\r\nELEV: " + Functions.Get_String_Rounded(elev1, 2) + "'";
                            MLeader ML1 = Functions.creaza_mleader_with_style(new Point3d(x1, y1, 0), content1, textH, mleaderstyle_name, textstyle_name, 3 * textH, 4 * textH, hdd_design);
                            ML1.SetTextAttachmentType(TextAttachmentType.AttachmentMiddleOfTop, LeaderDirectionType.RightLeader);
                            ML1.SetTextAttachmentType(TextAttachmentType.AttachmentMiddleOfTop, LeaderDirectionType.LeftLeader);
                            Functions.Populate_object_data_table_from_objectid(Tables1, ML1.ObjectId, OD_table_name, OD_lista_val, OD_lista_type);

                            string content2 = "PVC\r\nSTA: " + Functions.Get_chainage_from_double(Station_equation_of(lr * hd2 + sta1, dt_steq), "f", 0) + "\r\nELEV: " + Functions.Get_String_Rounded(elev2, 2) + "'";
                            MLeader ML2 = Functions.creaza_mleader_with_style(new Point3d(x2, y2, 0), content2, textH, mleaderstyle_name, textstyle_name, 3 * textH, -4 * textH, hdd_design);
                            ML2.SetTextAttachmentType(TextAttachmentType.AttachmentMiddleOfTop, LeaderDirectionType.RightLeader);
                            ML2.SetTextAttachmentType(TextAttachmentType.AttachmentMiddleOfTop, LeaderDirectionType.LeftLeader);
                            Functions.Populate_object_data_table_from_objectid(Tables1, ML2.ObjectId, OD_table_name, OD_lista_val, OD_lista_type);

                            string content3 = "PVT\r\nSTA: " + Functions.Get_chainage_from_double(Station_equation_of(lr * hd3 + sta1, dt_steq), "f", 0) + "\r\nELEV: " + Functions.Get_String_Rounded(elev3, 2) + "'";
                            MLeader ML3 = Functions.creaza_mleader_with_style(new Point3d(x3, y3, 0), content3, textH, mleaderstyle_name, textstyle_name, 3 * textH, -4 * textH, hdd_design);
                            ML3.SetTextAttachmentType(TextAttachmentType.AttachmentMiddleOfTop, LeaderDirectionType.RightLeader);
                            ML3.SetTextAttachmentType(TextAttachmentType.AttachmentMiddleOfTop, LeaderDirectionType.LeftLeader);
                            Functions.Populate_object_data_table_from_objectid(Tables1, ML3.ObjectId, OD_table_name, OD_lista_val, OD_lista_type);

                            string content4 = "PVC\r\nSTA: " + Functions.Get_chainage_from_double(Station_equation_of(lr * hd4 + sta1, dt_steq), "f", 0) + "\r\nELEV: " + Functions.Get_String_Rounded(elev4, 2) + "'";
                            MLeader ML4 = Functions.creaza_mleader_with_style(new Point3d(x4, y4, 0), content4, textH, mleaderstyle_name, textstyle_name, 3 * textH, -4 * textH, hdd_design);
                            ML4.SetTextAttachmentType(TextAttachmentType.AttachmentMiddleOfTop, LeaderDirectionType.RightLeader);
                            ML4.SetTextAttachmentType(TextAttachmentType.AttachmentMiddleOfTop, LeaderDirectionType.LeftLeader);
                            Functions.Populate_object_data_table_from_objectid(Tables1, ML4.ObjectId, OD_table_name, OD_lista_val, OD_lista_type);

                            string content5 = "PVT\r\nSTA: " + Functions.Get_chainage_from_double(Station_equation_of(lr * hd5 + sta1, dt_steq), "f", 0) + "\r\nELEV: " + Functions.Get_String_Rounded(elev5, 2) + "'";
                            MLeader ML5 = Functions.creaza_mleader_with_style(new Point3d(x5, y5, 0), content5, textH, mleaderstyle_name, textstyle_name, 3 * textH, -4 * textH, hdd_design);
                            ML5.SetTextAttachmentType(TextAttachmentType.AttachmentMiddleOfTop, LeaderDirectionType.RightLeader);
                            ML5.SetTextAttachmentType(TextAttachmentType.AttachmentMiddleOfTop, LeaderDirectionType.LeftLeader);
                            Functions.Populate_object_data_table_from_objectid(Tables1, ML5.ObjectId, OD_table_name, OD_lista_val, OD_lista_type);

                            string content6 = "ENTRY POINT @ " + Functions.Get_String_Rounded(Angle2, 0) + "%%D\r\nSTA: " + Functions.Get_chainage_from_double(Station_equation_of(sta1 + lr * hd6, dt_steq), "f", 0) + "\r\nELEV: " + Functions.Get_String_Rounded(elev6, 2) + "'";
                            MLeader ML6 = Functions.creaza_mleader_with_style(new Point3d(x6, y6, 0), content6, textH, mleaderstyle_name, textstyle_name, 3 * textH, 4 * textH, hdd_design);
                            ML6.SetTextAttachmentType(TextAttachmentType.AttachmentMiddleOfTop, LeaderDirectionType.RightLeader);
                            ML6.SetTextAttachmentType(TextAttachmentType.AttachmentMiddleOfTop, LeaderDirectionType.LeftLeader);
                            Functions.Populate_object_data_table_from_objectid(Tables1, ML6.ObjectId, OD_table_name, OD_lista_val, OD_lista_type);

                            string content7 = Functions.Get_String_Rounded(radius1, 0) + "' RADIUS";
                            MLeader ML7 = Functions.creaza_mleader_with_style(new Point3d(x7, y7, 0), content7, textH, mleaderstyle_name, textstyle_name, 3 * textH, -4 * textH, hdd_design);
                            ML7.SetTextAttachmentType(TextAttachmentType.AttachmentMiddleOfTop, LeaderDirectionType.RightLeader);
                            ML7.SetTextAttachmentType(TextAttachmentType.AttachmentMiddleOfTop, LeaderDirectionType.LeftLeader);
                            ML7.ArrowSymbolId = ObjectId.Null;
                            Functions.Populate_object_data_table_from_objectid(Tables1, ML7.ObjectId, OD_table_name, OD_lista_val, OD_lista_type);

                            string content8 = Functions.Get_String_Rounded(radius2, 0) + "' RADIUS";
                            MLeader ML8 = Functions.creaza_mleader_with_style(new Point3d(x8, y8, 0), content8, textH, mleaderstyle_name, textstyle_name, 3 * textH, -4 * textH, hdd_design);
                            ML8.SetTextAttachmentType(TextAttachmentType.AttachmentMiddleOfTop, LeaderDirectionType.RightLeader);
                            ML8.SetTextAttachmentType(TextAttachmentType.AttachmentMiddleOfTop, LeaderDirectionType.LeftLeader);
                            ML8.ArrowSymbolId = ObjectId.Null;
                            Functions.Populate_object_data_table_from_objectid(Tables1, ML8.ObjectId, OD_table_name, OD_lista_val, OD_lista_type);
                        }

                        if (dim_style_table != null)
                        {
                            ObjectId dimId = ObjectId.Null;
                            if (dim_style_table.Has(dimstyle_name) == true)
                            {
                                dimId = dim_style_table[dimstyle_name];
                                int idx = 1;
                                do
                                {
                                    ++idx;
                                } while (dim_style_table.Has("HDD" + idx.ToString()) == true);
                                DimStyleTableRecord dim2 = Trans1.GetObject(dimId, OpenMode.ForWrite) as DimStyleTableRecord;
                                dim2.Name = "HDD" + idx.ToString();
                            }


                            DimStyleTableRecord dim1 = new DimStyleTableRecord();
                            dim1.Name = dimstyle_name;
                            dim1.Dimadec = 0;
                            dim1.Dimasz = textH;
                            //Controls the size of dimension line and leader line arrowheads. Also controls the size of hook lines
                            //Multiples of the arrowhead size determine whether dimension lines and text should fit between the extension lines. DIMASZ is also used to scale arrowhead blocks if set by DIMBLK. DIMASZ has no effect when DIMTSZ is other than zero
                            dim1.Dimdec = 0;
                            //Sets the number of decimal places displayed for the primary units of a dimension
                            //The precision is based on the units or angle format you have selected. 
                            dim1.Dimtxt = textH;
                            //Specifies the height of dimension text, unless the current text style has a fixed height
                            dim1.Dimtxsty = HDD_textstyle.ObjectId;
                            dim1.Dimtxtdirection = false;
                            //Specifies the reading direction of the dimension text. 
                            //0 - Displays dimension text in a Left-to-Right reading style 
                            //1 - Displays dimension text in a Right-to-Left reading style  
                            dim1.Dimtofl = false;
                            //Initial value: Off (imperial) or On (metric)  
                            //Controls whether a dimension line is drawn between the extension lines even when the text is placed outside. 
                            //For radius and diameter dimensions (when DIMTIX is off), draws a dimension line inside the circle or arc and places the text, arrowheads, and leader outside. 
                            // Off -  Does not draw dimension lines between the measured points when arrowheads are placed outside the measured points 
                            // On -  Draws dimension lines between the measured points even when arrowheads are placed outside the measured points 
                            dim1.Dimtoh = false;
                            //Controls the position of dimension text outside the extension lines. 
                            // Off -  Aligns text with the dimension line
                            // On -  Draws text horizontally
                            dim1.Dimtih = false;
                            //Initial value: On (imperial) or Off (metric)  
                            //Controls the position of dimension text inside the extension lines for all dimension types except Ordinate. 
                            //Off - Aligns text with the dimension line
                            //On -  Draws text horizontally
                            dim1.Dimtad = 0;
                            //Controls the vertical position of text in relation to the dimension line. 
                            //0 - Centers the dimension text between the extension lines. 
                            //1 - Places the dimension text above the dimension line except when the dimension line is not horizontal and text inside the extension lines is forced horizontal ( DIMTIH = 1). 
                            //    The distance from the dimension line to the baseline of the lowest line of text is the current DIMGAP value. 
                            //2 - Places the dimension text on the side of the dimension line farthest away from the defining points. 
                            //3 - Places the dimension text to conform to Japanese Industrial Standards (JIS). 
                            //4 - Places the dimension text below the dimension line. 
                            dim1.Dimtvp = 0;
                            //Controls the vertical position of dimension text above or below the dimension line. 
                            //The DIMTVP value is used when DIMTAD is off. The magnitude of the vertical offset of text is the product of the text height and DIMTVP. 
                            //Setting DIMTVP to 1.0 is equivalent to setting DIMTAD to on. The dimension line splits to accommodate the text only if the absolute value of DIMTVP is less than 0.7. 
                            dim1.Dimsd1 = false;
                            //Controls suppression of the first dimension line and arrowhead. 
                            //When turned on, suppresses the display of the dimension line and arrowhead between the first extension line and the text. 
                            dim1.Dimsd2 = false;
                            //Controls suppression of the second dimension line and arrowhead. 
                            //When turned on, suppresses the display of the dimension line and arrowhead between the second extension line and the text. 
                            dim1.Dimse1 = false;
                            //Suppresses display of the first extension line. 
                            dim1.Dimse2 = false;
                            //Suppresses display of the second extension line
                            dim1.Dimrnd = 0;
                            //Rounds all dimensioning distances to the specified value. 
                            //For instance, if DIMRND is set to 0.25, all distances round to the nearest 0.25 unit. 
                            //If you set DIMRND to 1.0, all distances round to the nearest integer. 
                            //Note that the number of digits edited after the decimal point depends on the precision set by DIMDEC. DIMRND does not apply to angular dimensions. 
                            dim1.Dimpost = "'";
                            //Specifies a text prefix or suffix (or both) to the dimension measurement. 
                            //For example, to establish a suffix for millimeters, set DIMPOST to mm; a distance of 19.2 units would be displayed as 19.2 mm. 
                            //If tolerances are turned on, the suffix is applied to the tolerances as well as to the main dimension. 
                            //Use <> to indicate placement of the text in relation to the dimension value. 
                            //For example, enter <>mm to display a 5.0 millimeter radial dimension as "5.0mm." 
                            //If you entered mm <>, the dimension would be displayed as "mm 5.0." 
                            //Use the <> mechanism for angular dimensions. 
                            dim1.Dimjust = 0;
                            //Controls the horizontal positioning of dimension text. 
                            //0 -  Positions the text above the dimension line and center-justifies it between the extension lines 
                            //1 -  Positions the text next to the first extension line 
                            //2 -  Positions the text next to the second extension line 
                            //3 -  Positions the text above and aligned with the first extension line 
                            //4 -  Positions the text above and aligned with the second extension line 
                            dim1.Dimadec = 0;
                            //Controls the number of precision places displayed in angular dimensions. (0-8)
                            dim1.Dimalt = false;
                            //Controls the display of alternate units in dimensions. Off - Disables alternate units
                            dim1.Dimaltd = 2;
                            //Controls the number of decimal places in alternate units. If DIMALT is turned on, DIMALTD sets the number of digits displayed to the right of the decimal point in the alternate measurement
                            dim1.Dimaltf = 25.4;
                            //Controls the multiplier for alternate units. If DIMALT is turned on, DIMALTF multiplies linear dimensions by a factor to produce a value in an alternate system of measurement. The initial value represents the number of millimeters in an inch.
                            dim1.Dimaltrnd = 0;
                            //Rounds off the alternate dimension units. 
                            dim1.Dimalttd = 2;
                            //Sets the number of decimal places for the tolerance values in the alternate units of a dimension. 
                            dim1.Dimalttz = 0;
                            //Controls suppression of zeros in tolerance values. 
                            dim1.Dimaltu = 2;
                            //Sets the units format for alternate units of all dimension substyles except Angular. (2 - Decimal)
                            dim1.Dimaltz = 0;
                            //Controls the suppression of zeros for alternate unit dimension values. 
                            dim1.Dimapost = "";
                            //Specifies a text prefix or suffix (or both) to the alternate dimension measurement for all types of dimensions except angular. 
                            //For instance, if the current units are Architectural, DIMALT is on, DIMALTF is 25.4 (the number of millimeters per inch), DIMALTD is 2, and DIMPOST is set to "mm," a distance of 10 units would be displayed as 10"[254.00mm]. 
                            //To turn off an established prefix or suffix (or both), set it to a single period (.). 
                            dim1.Dimarcsym = 0;
                            //Controls display of the arc symbol in an arc length dimension. (0- Places arc length symbols before the dimension text )
                            //1 - Places arc length symbols above the dimension text 
                            //2 -  Suppresses the display of arc length symbols 
                            dim1.Dimatfit = 3;
                            //Determines how dimension text and arrows are arranged when space is not sufficient to place both within the extension lines. 
                            //0 -  Places both text and arrows outside extension lines 
                            //1 -  Moves arrows first, then text
                            //2 -  Moves text first, then arrows
                            //3 -  Moves either text or arrows, whichever fits best 
                            //A leader is added to moved dimension text when DIMTMOVE is set to 1. 
                            dim1.Dimaunit = 0;
                            //Sets the units format for angular dimensions. (0 - Decimal degrees)
                            dim1.Dimazin = 0;
                            //Suppresses zeros for angular dimensions. 
                            dim1.Dimsah = false;
                            //Controls the display of dimension line arrowhead blocks. 
                            //Off - Use arrowhead blocks set by DIMBLK
                            //On - Use arrowhead blocks set by DIMBLK1 and DIMBLK2
                            ObjectId Arrowid = ObjectId.Null;
                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("DIMBLK", ".");
                            //if (BlockTable1.Has("_Dot") == true)
                            //{
                            //    Arrowid = BlockTable1["_Dot"];
                            //}
                            dim1.Dimblk = Arrowid;
                            //Sets the arrowhead block displayed at the ends of dimension lines or leader lines. 
                            //To return to the default, closed-filled arrowhead display, enter a single period (.). Arrowhead block entries and the names used to select them in the New, Modify, and Override Dimension Style dialog boxes are shown below. You can also enter the names of user-defined arrowhead blocks. 
                            //Note Annotative blocks cannot be used as custom arrowheads for dimensions or leaders. 
                            //"" - Closed(filled)
                            //"_DOT" - dot
                            //"_DOTSMALL" - dot small
                            //"_DOTBLANK" - dot blank
                            //"_ORIGIN" - origin indicator
                            //"_ORIGIN2" - origin indicator 2
                            //"_OPEN" - open
                            //"_OPEN90" - Right(angle)
                            //"_OPEN30" - open 30
                            //"_CLOSED" - Closed
                            //"_SMALL" - dot small blank
                            //"_NONE" - none
                            //"_OBLIQUE" - oblique
                            //"_BOXFILLED" - box filled
                            //"_BOXBLANK" - box
                            //"_CLOSEDBLANK" - Closed(blank)
                            //"_DATUMFILLED" - datum triangle filled
                            //"_DATUMBLANK" - datum triangle
                            //"_INTEGRAL" - integral
                            //"_ARCHTICK" - architectural tick
                            dim1.Dimblk1 = Arrowid;
                            //Sets the arrowhead for the first end of the dimension line when DIMSAH is on. 
                            //To return to the default, closed-filled arrowhead display, enter a single period (.). For a list of arrowheads, see DIMBLK. 
                            //Note Annotative blocks cannot be used as custom arrowheads for dimensions or leaders. 
                            dim1.Dimblk2 = Arrowid;
                            //Sets the arrowhead for the second end of the dimension line when DIMSAH is on. 
                            //To return to the default, closed-filled arrowhead display, enter a single period (.). For a list of arrowhead entries, see DIMBLK. 
                            //Note Annotative blocks cannot be used as custom arrowheads for dimensions or leaders. 
                            dim1.Dimldrblk = Arrowid;
                            // Specifies the arrow type for leaders. 
                            dim1.Dimcen = 0.09;
                            //Controls drawing of circle or arc center marks and centerlines by the DIMCENTER, DIMDIAMETER, and DIMRADIUS commands. 
                            dim1.Dimclrd = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256);
                            // Assigns colors to dimension lines, arrowheads, and dimension leader lines
                            dim1.Dimclre = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256);
                            //Assigns colors to dimension extension lines.
                            dim1.Dimclrt = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256);
                            //Assigns colors to dimension text
                            dim1.Dimdle = 0;
                            //Sets the distance the dimension line extends beyond the extension line when oblique strokes are drawn instead of arrowheads. 
                            dim1.Dimdli = 0.38;
                            //Controls the spacing of the dimension lines in baseline dimensions. 
                            //Each dimension line is offset from the previous one by this amount, if necessary, to avoid drawing over it. Changes made with DIMDLI are not applied to existing dimensions
                            dim1.Dimdsep = (char)46;
                            //Specifies a single-character decimal separator to use when creating dimensions whose unit format is decimal
                            //When prompted, enter a single character at the Command prompt. If dimension units is set to Decimal, the DIMDSEP character is used instead of the default decimal point.
                            //If DIMDSEP is set to NULL (default value, reset by entering a period), the decimal point is used as the dimension separator
                            dim1.Dimexe = textH / 2;
                            //Specifies how far to extend the extension line beyond the dimension line. 
                            dim1.Dimexo = textH;
                            //Specifies how far extension lines are offset from origin points. 
                            //With fixed-length extension lines, this value determines the minimum offset. 
                            dim1.Dimfrac = 0;
                            //Sets the fraction format when DIMLUNIT is set to 4 (Architectural) or 5 (Fractional).
                            //0 - Horizontal stacking
                            //1 - Diagonal stacking
                            //2 - Not stacked (for example, 1/2)
                            dim1.Dimfxlen = 1;
                            dim1.DimfxlenOn = false;
                            dim1.Dimgap = textH;
                            //Sets the distance around the dimension text when the dimension line breaks to accommodate dimension text.
                            dim1.Dimjogang = 0.785398163397448;
                            //Determines the angle of the transverse segment of the dimension line in a jogged radius dimension. 
                            dim1.Dimlfac = 1;
                            //Sets a scale factor for linear dimension measurements. 
                            //All linear dimension distances, including radii, diameters, and coordinates, are multiplied by DIMLFAC before being converted to dimension text. Positive values of DIMLFAC are applied to dimensions in both model space and paper space; negative values are applied to paper space only. 
                            //DIMLFAC applies primarily to nonassociative dimensions (DIMASSOC set 0 or 1). For nonassociative dimensions in paper space, DIMLFAC must be set individually for each layout viewport to accommodate viewport scaling. 
                            //DIMLFAC has no effect on angular dimensions, and is not applied to the values held in DIMRND, DIMTM, or DIMTP. 
                            dim1.Dimltex1 = ThisDrawing.Database.ByBlockLinetype;
                            //Sets the linetype of the first extension line. 
                            dim1.Dimltex2 = ThisDrawing.Database.ByBlockLinetype;
                            //Sets the linetype of the second extension line. 
                            dim1.Dimltype = ThisDrawing.Database.ByBlockLinetype;
                            //Sets the linetype of the dimension line.
                            dim1.Dimlunit = 2;
                            //Sets units for all dimension types except Angular. 
                            //1 Scientific
                            //2 Decimal
                            //3 Engineering
                            //4 Architectural (always displayed stacked)
                            //5 Fractional (always displayed stacked)
                            //6 Microsoft Windows Desktop (decimal format using Control Panel settings for decimal separator and number grouping symbols) 
                            dim1.Dimlwd = LineWeight.ByBlock;
                            //Assigns lineweight to dimension lines. 
                            //-3 Default (the LWDEFAULT value) 
                            //-2 BYBLOCK
                            //-1 BYLAYER
                            dim1.Dimlwe = LineWeight.ByBlock;
                            //Assigns lineweight to extension  lines. 
                            //-3 Default (the LWDEFAULT value) 
                            //-2 BYBLOCK
                            //-1 BYLAYER
                            dim1.Dimscale = 1;
                            //Sets the overall scale factor applied to dimensioning variables that specify sizes, distances, or offsets. 
                            //Also affects the leader objects with the LEADER command. 
                            //Use MLEADERSCALE to scale multileader objects created with the MLEADER command. 
                            //0.0 - A reasonable default value is computed based on the scaling between the current model space viewport and paper space. 
                            //If you are in paper space or model space and not using the paper space feature, the scale factor is 1.0. 
                            //>0 - A scale factor is computed that leads text sizes, arrowhead sizes, and other scaled distances to plot at their face values. 
                            // DIMSCALE does not affect measured lengths, coordinates, or angles. 
                            //Use DIMSCALE to control the overall scale of dimensions. However, if the current dimension style is annotative, 
                            //DIMSCALE is automatically set to zero and the dimension scale is controlled by the CANNOSCALE system variable. DIMSCALE cannot be set to a non-zero value when using annotative dimensions. 
                            dim1.Dimtdec = 0;
                            //Sets the number of decimal places to display in tolerance values for the primary units in a dimension. 
                            //This system variable has no effect unless DIMTOL is set to On. The default for DIMTOL is Off. 
                            dim1.Dimtfac = 1;
                            //Specifies a scale factor for the text height of fractions and tolerance values relative to the dimension text height, as set by DIMTXT. 
                            //For example, if DIMTFAC is set to 1.0, the text height of fractions and tolerances is the same height as the dimension text. 
                            //If DIMTFAC is set to 0.7500, the text height of fractions and tolerances is three-quarters the size of dimension text. 
                            dim1.Dimtfill = 1;
                            //Controls the background of dimension text. 
                            //0 -  No Background
                            //1 -  The background color of the drawing 
                            //2 -  The background specified by DIMTFILLCLR
                            dim1.Dimtfillclr = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByBlock, 0);
                            dim1.Dimtix = false;
                            //Draws text between extension lines. 
                            //Off -  Varies with the type of dimension. 
                            //        For linear and angular dimensions, text is placed inside the extension lines if there is sufficient room. 
                            //         For radius and diameter dimensions that dont fit inside the circle or arc, DIMTIX has no effect and always forces the text outside the circle or arc. 
                            //On -  Draws dimension text between the extension lines even if it would ordinarily be placed outside those lines 
                            dim1.Dimsoxd = false;
                            //Suppresses arrowheads if not enough space is available inside the extension lines. 
                            //Off -  Arrowheads are not suppressed
                            //On -  Arrowheads are suppressed
                            //If not enough space is available inside the extension lines and DIMTIX is on, setting DIMSOXD to On suppresses the arrowheads. If DIMTIX is off, DIMSOXD has no effect. 
                            dim1.Dimtm = 0;
                            //Sets the minimum (or lower) tolerance limit for dimension text when DIMTOL or DIMLIM is on. 
                            //DIMTM accepts signed values. If DIMTOL is on and DIMTP and DIMTM are set to the same value, a tolerance value is drawn. 
                            //If DIMTM and DIMTP values differ, the upper tolerance is drawn above the lower, and a plus sign is added to the DIMTP value if it is positive. 
                            //For DIMTM, the program uses the negative of the value you enter (adding a minus sign if you specify a positive number and a plus sign if you specify a negative number). 
                            dim1.Dimtmove = 0;
                            //Sets dimension text movement rules. 
                            //0 -  Moves the dimension line with dimension text
                            //1 -  Adds a leader when dimension text is moved
                            //2 -  Allows text to be moved freely without a leader
                            dim1.Dimtp = 0;
                            //Sets the maximum (or upper) tolerance limit for dimension text when DIMTOL or DIMLIM is on. 
                            //DIMTP accepts signed values. If DIMTOL is on and DIMTP and DIMTM are set to the same value, a tolerance value is drawn. 
                            //If DIMTM and DIMTP values differ, the upper tolerance is drawn above the lower and a plus sign is added to the DIMTP value if it is positive. 
                            dim1.Dimlim = false;
                            //Generates dimension limits as the default text. 
                            //Setting DIMLIM to On turns DIMTOL off. 
                            //Off -  Dimension limits are not generated as default text 
                            //On -  Dimension limits are generated as default text
                            dim1.Dimtol = false;
                            //Appends tolerances to dimension text. 
                            //Setting DIMTOL to on turns DIMLIM off. 
                            dim1.Dimtolj = 1;
                            //Sets the vertical justification for tolerance values relative to the nominal dimension text. 
                            dim1.Dimtsz = 0;
                            //Specifies the size of oblique strokes drawn instead of arrowheads for linear, radius, and diameter dimensioning. 
                            //0 -  Draws arrowheads.
                            //>0 -  Draws oblique strokes instead of arrowheads. The size of the oblique strokes is determined by this value multiplied by the DIMSCALE value 
                            dim1.Dimtzin = 0;
                            //Controls the suppression of zeros in tolerance values. 
                            dim1.Dimupt = false;
                            //Controls options for user-positioned text. 
                            //Off -  Cursor controls only the dimension line location
                            //On -  Cursor controls both the text position and the dimension line location 
                            dim1.Dimzin = 0;
                            //Controls the suppression of zeros in the primary unit value. 
                            //Values 0-3 affect feet-and-inch dimensions only: 
                            //0 -  Suppresses zero feet and precisely zero inches
                            //1 -  Includes zero feet and precisely zero inches
                            // 2 -  Includes zero feet and suppresses zero inches
                            //3 -  Includes zero inches and suppresses zero feet
                            //4 -  Suppresses leading zeros in decimal dimensions (for example, 0.5000 becomes .5000) 
                            //8 -  Suppresses trailing zeros in decimal dimensions (for example, 12.5000 becomes 12.5) 
                            //12 -  Suppresses both leading and trailing zeros (for example, 0.5000 becomes .5) 
                            dimId = dim_style_table.Add(dim1);
                            Trans1.AddNewlyCreatedDBObject(dim1, true);

                            DimStyleTableRecord new_dim_style = (DimStyleTableRecord)Trans1.GetObject(dimId, OpenMode.ForRead);
                            if (new_dim_style.ObjectId != ThisDrawing.Database.Dimstyle)
                            {
                                ThisDrawing.Database.Dimstyle = new_dim_style.ObjectId;
                                ThisDrawing.Database.SetDimstyleData(new_dim_style);
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

            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            set_enable_true();
        }

        private void Button_load_HDD_design_Click(object sender, EventArgs e)
        {
            this.MdiParent.WindowState = FormWindowState.Minimized;
            set_enable_false();
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                Matrix3d CurentUCSmatrix = Editor1.CurrentUserCoordinateSystem;
                ObjectId[] Empty_array = null;
                Editor1.SetImpliedSelection(Empty_array);
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;
                        Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat6;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt6;
                        Prompt6 = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect an existing Hdd Design:");
                        Prompt6.SetRejectMessage("\nSelect a polyline or a block reference");
                        Prompt6.AllowNone = true;
                        Prompt6.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Prompt6.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.BlockReference), false);
                        Rezultat6 = ThisDrawing.Editor.GetEntity(Prompt6);

                        if (Rezultat6.Status == PromptStatus.OK)
                        {
                            Entity ent1 = Trans1.GetObject(Rezultat6.ObjectId, OpenMode.ForRead) as Entity;

                            if (ent1 != null)
                            {
                                using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tables1.GetObjectRecords(Convert.ToUInt32(0), ent1.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, false))
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
                                                    if (Nume_field == "Sta1")
                                                    {
                                                        if (valoare1 != null) textBox_sta1.Text = Convert.ToString(valoare1);
                                                    }
                                                    if (Nume_field == "Angle_start")
                                                    {
                                                        if (valoare1 != null) textBox_angle_start.Text = Convert.ToString(valoare1);
                                                    }
                                                    if (Nume_field == "L1")
                                                    {
                                                        if (valoare1 != null) textBox_L1.Text = Convert.ToString(valoare1);
                                                    }
                                                    if (Nume_field == "Radius1")
                                                    {
                                                        if (valoare1 != null) textBox_radius1.Text = Convert.ToString(valoare1);
                                                    }
                                                    if (Nume_field == "L3")
                                                    {
                                                        if (valoare1 != null) textBox_L3.Text = Convert.ToString(valoare1);
                                                    }
                                                    if (Nume_field == "H_Angle")
                                                    {
                                                        if (valoare1 != null)
                                                        {
                                                            textBox_dev_angle1.Text = Convert.ToString(valoare1);
                                                            textBox_dev_angle2.Text = Convert.ToString(valoare1);
                                                        }
                                                    }
                                                    if (Nume_field == "Radius2")
                                                    {
                                                        if (valoare1 != null) textBox_radius2.Text = Convert.ToString(valoare1);
                                                    }
                                                    if (Nume_field == "Angle_end")
                                                    {
                                                        if (valoare1 != null) textBox_angle_end.Text = Convert.ToString(valoare1);
                                                    }
                                                }
                                            }
                                            button_HDD_Calculate_parameters_Click(sender, e);
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
        }

        private void checkBox_exaggeration_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_exaggeration.Checked == true)
            {
                panel_exaggeration.Visible = true;
            }
            else
            {
                panel_exaggeration.Visible = false;
            }
        }
    }




}
