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
    public partial class cogo_points_form : Form
    {




        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_add_cogo_points_EG);
            lista_butoane.Add(button_load_layers_to_combobox1);
            lista_butoane.Add(button_load_layers_to_combobox2);
            lista_butoane.Add(button_load_surface);

            lista_butoane.Add(button_udp1);
            lista_butoane.Add(button_udp2);
            lista_butoane.Add(button_load_point_style);
            lista_butoane.Add(button_load_point_label_style);
            lista_butoane.Add(button_load_layers_to_combobox_point_layer);


            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = false;
            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_add_cogo_points_EG);
            lista_butoane.Add(button_load_layers_to_combobox1);
            lista_butoane.Add(button_load_layers_to_combobox2);

            lista_butoane.Add(button_load_surface);

            lista_butoane.Add(button_udp1);
            lista_butoane.Add(button_udp2);
            lista_butoane.Add(button_load_point_style);
            lista_butoane.Add(button_load_point_label_style);
            lista_butoane.Add(button_load_layers_to_combobox_point_layer);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }

        public cogo_points_form()
        {
            InitializeComponent();

            if (Functions.is_dan_popescu() == true)
            {

            }

            var toolStripMenuItem2 = new ToolStripMenuItem { Text = "tool strip menu" };
            //toolStripMenuItem2.Click += go_to_excel_point;
            //ContextMenuStrip_go_to_error = new ContextMenuStrip();
            //ContextMenuStrip_go_to_error.Items.AddRange(new ToolStripItem[] { toolStripMenuItem2 });
        }

        private void button_load_layers_to_combobox1_Click(object sender, EventArgs e)
        {
            Functions.Incarca_existing_layers_to_combobox(comboBox_layer1);
        }

        private void button_load_layers_to_combobox2_Click(object sender, EventArgs e)
        {
            Functions.Incarca_existing_layers_to_combobox(comboBox_layer2);
        }

        private void button_convert_poly_to_alignment_Click(object sender, EventArgs e)
        {
            if (Functions.IsNumeric(textBox_X_spacing.Text) == false || Functions.IsNumeric(textBox_Y_spacing.Text) == false) return;

            double DeltaX = Convert.ToDouble(textBox_X_spacing.Text);
            double DeltaY = Convert.ToDouble(textBox_Y_spacing.Text);

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
                        LayerTable LayerTable1 = Trans1.GetObject(ThisDrawing.Database.LayerTableId, OpenMode.ForRead) as LayerTable;
                        if (LayerTable1.Has(comboBox_layer1.Text) == true)
                        {
                            string layer1 = comboBox_layer1.Text;
                            string layer2 = comboBox_layer2.Text;

                            ObjectIdCollection lista1 = new ObjectIdCollection();
                            ObjectIdCollection lista2 = new ObjectIdCollection();
                            foreach (ObjectId id1 in BTrecord)
                            {
                                Polyline poly1 = Trans1.GetObject(id1, OpenMode.ForRead) as Polyline;
                                if (poly1 != null && poly1.Closed == true && (poly1.Layer == layer1 || poly1.Layer == layer2))
                                {
                                    if (poly1.Layer == layer1)
                                    {
                                        lista1.Add(id1);
                                    }
                                    else
                                    {
                                        lista2.Add(id1);
                                    }
                                }
                            }

                            Autodesk.Civil.DatabaseServices.Surface surf1 = null;
                            ObjectIdCollection col_surf = CivilDrawing.GetSurfaceIds();
                            for (int j = 0; j < col_surf.Count; ++j)
                            {
                                Autodesk.Civil.DatabaseServices.Surface surf2 = Trans1.GetObject(col_surf[j], OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Surface;
                                if (surf2 != null)
                                {
                                    if (surf2.Name == comboBox_surfaces.Text)
                                    {
                                        surf1 = surf2;
                                        j = col_surf.Count;
                                    }
                                }
                            }


                            if (lista1.Count > 0)
                            {
                                double x0 = -1.1234;
                                double y0 = -1.1234;

                                double x1 = -1.1234;
                                double y1 = -1.1234;

                                for (int i = 0; i < lista1.Count; i++)
                                {
                                    Polyline poly1 = Trans1.GetObject(lista1[i], OpenMode.ForRead) as Polyline;
                                    if (poly1 != null)
                                    {
                                        for (int j = 0; j < poly1.NumberOfVertices; j++)
                                        {
                                            Point3d pt1 = poly1.GetPointAtParameter(j);
                                            double x = pt1.X;
                                            double y = pt1.Y;

                                            if (x0 == -1.1234) x0 = x;
                                            if (x1 == -1.1234) x1 = x;
                                            if (x < x0)
                                            {
                                                x0 = x;
                                            }
                                            if (x > x1)
                                            {
                                                x1 = x;
                                            }
                                            if (y0 == -1.1234) y0 = y;
                                            if (y1 == -1.1234) y1 = y;
                                            if (y < y0)
                                            {
                                                y0 = y;
                                            }

                                            if (y > y1)
                                            {
                                                y1 = y;
                                            }
                                        }
                                    }
                                }
                                double no_x = Math.Ceiling((x1 - x0) / DeltaX);
                                if (no_x == 0) no_x = 1;

                                double x_max = x0 + no_x * DeltaX;

                                double no_y = Math.Ceiling((y1 - y0) / DeltaY);
                                if (no_y == 0) no_y = 1;

                                double y_max = y0 + no_y * DeltaY;

                                CogoPointCollection cogoPoints = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument.CogoPoints;
                                List<Autodesk.AutoCAD.DatabaseServices.Region> lista_reg1 = new List<Autodesk.AutoCAD.DatabaseServices.Region>();
                                List<Autodesk.AutoCAD.DatabaseServices.Region> lista_reg2 = new List<Autodesk.AutoCAD.DatabaseServices.Region>();

                                for (int k = 0; k < lista1.Count; ++k)
                                {
                                    Polyline poly1 = Trans1.GetObject(lista1[k], OpenMode.ForRead) as Polyline;
                                    if (poly1 != null)
                                    {
                                        DBObjectCollection Poly_Colection = new DBObjectCollection();
                                        Poly_Colection.Add(poly1);
                                        DBObjectCollection col_reg = new DBObjectCollection();
                                        col_reg = Autodesk.AutoCAD.DatabaseServices.Region.CreateFromCurves(Poly_Colection);
                                        Autodesk.AutoCAD.DatabaseServices.Region reg1 = col_reg[0] as Autodesk.AutoCAD.DatabaseServices.Region;
                                        lista_reg1.Add(reg1);
                                    }
                                }

                                if (lista2.Count > 0)
                                    for (int k = 0; k < lista2.Count; ++k)
                                    {
                                        Polyline poly2 = Trans1.GetObject(lista2[k], OpenMode.ForRead) as Polyline;
                                        if (poly2 != null)
                                        {
                                            DBObjectCollection Poly_Colection = new DBObjectCollection();
                                            Poly_Colection.Add(poly2);
                                            DBObjectCollection col_reg = new DBObjectCollection();
                                            col_reg = Autodesk.AutoCAD.DatabaseServices.Region.CreateFromCurves(Poly_Colection);
                                            Autodesk.AutoCAD.DatabaseServices.Region reg2 = col_reg[0] as Autodesk.AutoCAD.DatabaseServices.Region;
                                            lista_reg2.Add(reg2);
                                        }
                                    }



                                for (int i = 0; i <= no_x; i++)
                                {
                                    for (int j = 0; j <= no_y; j++)
                                    {
                                        bool insert_cogo = false;
                                        Point3d inspt = new Point3d(x0 + i * DeltaX, y0 + j * DeltaY, 0);
                                        try
                                        {
                                            for (int k = 0; k < lista_reg1.Count; ++k)
                                            {
                                                Autodesk.AutoCAD.DatabaseServices.Region reg1 = lista_reg1[k];

                                                if (reg1 != null)
                                                {
                                                    Autodesk.AutoCAD.BoundaryRepresentation.PointContainment pc1 = Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Outside;

                                                    using (Autodesk.AutoCAD.BoundaryRepresentation.Brep Brep_obj = new Autodesk.AutoCAD.BoundaryRepresentation.Brep(reg1))
                                                    {
                                                        if (Brep_obj != null)
                                                        {
                                                            using (Autodesk.AutoCAD.BoundaryRepresentation.BrepEntity ent1 = Brep_obj.GetPointContainment(inspt, out pc1))
                                                            {
                                                                if (ent1 is Autodesk.AutoCAD.BoundaryRepresentation.Face || ent1 is Autodesk.AutoCAD.BoundaryRepresentation.Edge)
                                                                {
                                                                    insert_cogo = true;
                                                                    k = lista_reg1.Count;
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        catch (Autodesk.AutoCAD.Runtime.Exception)
                                        {
                                            insert_cogo = false;
                                        }

                                        if (lista_reg2.Count > 0)
                                        {
                                            if (insert_cogo == true)
                                            {
                                                try
                                                {
                                                    for (int k = 0; k < lista_reg2.Count; ++k)
                                                    {
                                                        Autodesk.AutoCAD.DatabaseServices.Region reg2 = lista_reg2[k];

                                                        if (reg2 != null)
                                                        {
                                                            Autodesk.AutoCAD.BoundaryRepresentation.PointContainment pc1 = Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Outside;

                                                            using (Autodesk.AutoCAD.BoundaryRepresentation.Brep Brep_obj = new Autodesk.AutoCAD.BoundaryRepresentation.Brep(reg2))
                                                            {
                                                                if (Brep_obj != null)
                                                                {
                                                                    using (Autodesk.AutoCAD.BoundaryRepresentation.BrepEntity ent1 = Brep_obj.GetPointContainment(inspt, out pc1))
                                                                    {
                                                                        if (ent1 is Autodesk.AutoCAD.BoundaryRepresentation.Face)
                                                                        {
                                                                            insert_cogo = false;
                                                                            k = lista_reg2.Count;
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                catch (Autodesk.AutoCAD.Runtime.Exception)
                                                {
                                                    insert_cogo = false;
                                                }
                                            }
                                        }

                                        if (insert_cogo == true)
                                        {
                                            double x = inspt.X;
                                            double y = inspt.Y;
                                            double elev1 = -1.12345;
                                            if (surf1 != null)
                                            {
                                                try
                                                {
                                                    elev1 = surf1.FindElevationAtXY(x, y);
                                                    inspt = new Point3d(x, y, elev1);
                                                }
                                                catch (System.Exception)
                                                {
                                                }
                                            }

                                            if (elev1 != -1.12345)
                                            {
                                                ObjectId cogoID = cogoPoints.Add(inspt, "Survey Point", true);
                                                CogoPoint cg1 = Trans1.GetObject(cogoID, OpenMode.ForWrite) as CogoPoint;//cogoID.GetObject(OpenMode.ForWrite) as CogoPoint;
                                                cg1.RawDescription = "EG";
                                                UDPString udp1 = Functions.Find_udp_string(comboBox_udp_field_1.Text);
                                                if (udp1 != null)
                                                {
                                                    cg1.SetUDPValue(udp1, "EG");
                                                }
                                                UDPDouble udp2 = Functions.Find_udp_double(comboBox_udp_field_2.Text);
                                                if (udp2 != null)
                                                {
                                                    cg1.SetUDPValue(udp2, elev1);
                                                }
                                                Autodesk.Civil.DatabaseServices.Styles.PointStyleCollection col_point_styles = CivilDrawing.Styles.PointStyles;
                                                IEnumerator<ObjectId> enum1 = col_point_styles.GetEnumerator();

                                                while (enum1.MoveNext())
                                                {
                                                    Autodesk.Civil.DatabaseServices.Styles.PointStyle pst1 = Trans1.GetObject(enum1.Current, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.PointStyle;
                                                    if (pst1.Name == comboBox_point_style.Text)
                                                    {
                                                        cg1.StyleId = pst1.ObjectId;
                                                    }
                                                }

                                                Autodesk.Civil.DatabaseServices.Styles.LabelStyleCollection col_point_label_styles = CivilDrawing.Styles.LabelStyles.PointLabelStyles.LabelStyles;

                                                IEnumerator<ObjectId> enum2 = col_point_label_styles.GetEnumerator();

                                                while (enum2.MoveNext())
                                                {
                                                    Autodesk.Civil.DatabaseServices.Styles.LabelStyle plst2 = Trans1.GetObject(enum2.Current, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.LabelStyle;
                                                    if (plst2.Name == comboBox_point_label_style.Text)
                                                    {
                                                        cg1.LabelStyleId = plst2.ObjectId;
                                                    }
                                                }
                                                string layer_pt = "0";
                                                if (comboBox_point_layer.Text.Length > 0)
                                                {
                                                    layer_pt = comboBox_point_layer.Text;
                                                }
                                                cg1.Layer = layer_pt;
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
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            set_enable_true();
        }













        private void button_load_surface_Click(object sender, EventArgs e)
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
                load_surfaces_to_combobox(comboBox_surfaces);

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            this.MdiParent.WindowState = FormWindowState.Normal;

            set_enable_true();
        }



        private void label_conv_poly_to_align_Click(object sender, EventArgs e)
        {

        }


        private void load_surfaces_to_combobox(ComboBox combo1)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.Civil.ApplicationServices.CivilDocument CivilDrawing = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument;
            using (DocumentLock lock1 = ThisDrawing.LockDocument())
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;
                    ObjectIdCollection col_surf = CivilDrawing.GetSurfaceIds();
                    combo1.Items.Clear();
                    for (int j = 0; j < col_surf.Count; ++j)
                    {
                        Autodesk.Civil.DatabaseServices.Surface surf1 = Trans1.GetObject(col_surf[j], OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Surface;
                        if (surf1 != null)
                        {
                            combo1.Items.Add(surf1.Name);
                        }
                    }
                    Trans1.Commit();
                }
            }
        }







        private void button_udp1_Click(object sender, EventArgs e)
        {
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

        private void button_load_point_style_Click(object sender, EventArgs e)
        {
            comboBox_point_style.Items.Clear();

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
                            comboBox_point_style.Items.Add(pst1.Name);
                        }

                        if (comboBox_point_style.Items.Contains("EG") == true)
                        {
                            comboBox_point_style.SelectedIndex = comboBox_point_style.Items.IndexOf("EG");
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
            comboBox_point_label_style.Items.Clear();

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

                            comboBox_point_label_style.Items.Add(plst2.Name);

                        }

                        if (comboBox_point_label_style.Items.Contains("Point#-Elevation-Description [EG]") == true)
                        {
                            comboBox_point_label_style.SelectedIndex = comboBox_point_label_style.Items.IndexOf("Point#-Elevation-Description [EG]");
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
            Functions.Incarca_existing_layers_to_combobox(comboBox_point_layer);

            if (comboBox_point_layer.Items.Contains("PTS-EG") == true)
            {
                comboBox_point_layer.SelectedIndex = comboBox_point_layer.Items.IndexOf("PTS-EG");
            }

        }
    }
}
