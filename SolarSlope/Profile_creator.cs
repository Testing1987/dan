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
    public partial class profiler_form : Form
    {




        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();

            lista_butoane.Add(button_convert_poly_to_alignment);
            lista_butoane.Add(button_create_profiles);
            lista_butoane.Add(button_load_align_layers);
            lista_butoane.Add(button_load_polyline_layers_to_combobox);



            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {

                bt1.Enabled = false;

            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_convert_poly_to_alignment);
            lista_butoane.Add(button_create_profiles);
            lista_butoane.Add(button_load_align_layers);
            lista_butoane.Add(button_load_polyline_layers_to_combobox);



            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }

        public profiler_form()
        {
            InitializeComponent();



            var toolStripMenuItem2 = new ToolStripMenuItem { Text = "tool strip menu" };
            //toolStripMenuItem2.Click += go_to_excel_point;


            //ContextMenuStrip_go_to_error = new ContextMenuStrip();
            //ContextMenuStrip_go_to_error.Items.AddRange(new ToolStripItem[] { toolStripMenuItem2 });


        }

        private void button_load_layers_to_combobox_Click(object sender, EventArgs e)
        {
            Functions.Incarca_existing_layers_to_combobox(comboBox_layer_poly);
            Functions.Incarca_existing_layers_to_combobox(comboBox_layer_alignment);
        }

        private void button_convert_poly_to_alignment_Click(object sender, EventArgs e)
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
                        LayerTable LayerTable1 = Trans1.GetObject(ThisDrawing.Database.LayerTableId, OpenMode.ForRead) as LayerTable;
                        if (LayerTable1.Has(comboBox_layer_poly.Text) == true)
                        {
                            string layer1 = comboBox_layer_poly.Text;

                            int count1 = 1;

                            ObjectIdCollection col1 = CivilDrawing.GetSiteIds();
                            List<string> lista1 = new List<string>();

                            for (int j = 0; j < col1.Count; ++j)
                            {
                                Autodesk.Civil.DatabaseServices.Site site1 = Trans1.GetObject(col1[j], OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Site;
                                if (site1 != null)
                                {
                                    lista1.Add(site1.Name);
                                }
                            }


                            string site_name = "Site_" + count1.ToString();

                            do
                            {
                                if (lista1.Contains(site_name) == true)
                                {
                                    ++count1;
                                    site_name = "Site_" + count1.ToString();
                                }
                            } while (lista1.Contains(site_name) == true);

                            ObjectId siteID = Autodesk.Civil.DatabaseServices.Site.Create(CivilDrawing, site_name);



                            string layer2 = "C-Align";

                            if (comboBox_layer_alignment.Text != "")
                            {
                                layer2 = comboBox_layer_alignment.Text;
                            }

                            Functions.Creaza_layer(layer2, 1, true);

                            Autodesk.Civil.DatabaseServices.Styles.AlignmentStyleCollection col_align_styles = CivilDrawing.Styles.AlignmentStyles;
                            List<string> lista2 = new List<string>();
                            for (int j = 0; j < col_align_styles.Count; ++j)
                            {
                                Autodesk.Civil.DatabaseServices.Styles.AlignmentStyle alstyle1 = Trans1.GetObject(col_align_styles[j], OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.AlignmentStyle;
                                if (alstyle1 != null)
                                {
                                    lista2.Add(alstyle1.Name);
                                }
                            }


                            Autodesk.Civil.DatabaseServices.Styles.AlignmentLabelSetStyleCollection col_label_styles = CivilDrawing.Styles.LabelSetStyles.AlignmentLabelSetStyles;
                            List<string> lista3 = new List<string>();
                            for (int j = 0; j < col_label_styles.Count; ++j)
                            {
                                Autodesk.Civil.DatabaseServices.Styles.AlignmentLabelSetStyle lstyle1 = Trans1.GetObject(col_label_styles[j], OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.AlignmentLabelSetStyle;
                                if (lstyle1 != null)
                                {
                                    lista3.Add(lstyle1.Name);
                                }
                            }


                            string alignment_style_name = "Proposed";
                            string alignment_label_set_style_name = "Major and Minor only";

                            if (lista2.Contains(alignment_style_name) == false)
                            {
                                alignment_style_name = lista2[0];
                            }


                            if (lista3.Contains(alignment_label_set_style_name) == false)
                            {
                                alignment_label_set_style_name = lista3[0];
                            }

                            ObjectIdCollection col4 = CivilDrawing.GetAlignmentIds();
                            List<string> lista4 = new List<string>();
                            for (int j = 0; j < col4.Count; ++j)
                            {
                                Autodesk.Civil.DatabaseServices.Alignment al1 = Trans1.GetObject(col4[j], OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Alignment;
                                if (al1 != null)
                                {
                                    lista4.Add(al1.Name);
                                }
                            }

                            count1 = 1;
                            string new_alignment_name = "AL_" + count1.ToString();

                            do
                            {
                                if (lista4.Contains(new_alignment_name) == true)
                                {
                                    ++count1;
                                    new_alignment_name = "AL_" + count1.ToString();
                                }
                            } while (lista4.Contains(new_alignment_name) == true);



                            foreach (ObjectId id1 in BTrecord)
                            {
                                Polyline poly1 = Trans1.GetObject(id1, OpenMode.ForRead) as Polyline;
                                if (poly1 != null && poly1.Layer == layer1)
                                {
                                    poly1.UpgradeOpen();
                                    PolylineOptions p_opt1 = new PolylineOptions();
                                    p_opt1.AddCurvesBetweenTangents = false;
                                    p_opt1.EraseExistingEntities = true;
                                    p_opt1.PlineId = id1;

                                    ObjectId AlignmentID = Alignment.Create(CivilDrawing, p_opt1, new_alignment_name, site_name, layer2, alignment_style_name, alignment_label_set_style_name);
                                    lista4.Add(new_alignment_name);
                                    ++count1;


                                    new_alignment_name = "AL_" + count1.ToString();
                                    do
                                    {
                                        if (lista4.Contains(new_alignment_name) == true)
                                        {
                                            ++count1;
                                            new_alignment_name = "AL_" + count1.ToString();
                                        }
                                    } while (lista4.Contains(new_alignment_name) == true);

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

        private void button_load_align_layers_Click(object sender, EventArgs e)
        {
            Functions.Incarca_existing_layers_to_combobox(comboBox_align_layers);
        }

        private void button_create_profiles_Click(object sender, EventArgs e)
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
                        LayerTable LayerTable1 = Trans1.GetObject(ThisDrawing.Database.LayerTableId, OpenMode.ForRead) as LayerTable;
                        if (LayerTable1.Has(comboBox_align_layers.Text) == true)
                        {

                            ObjectIdCollection col1 = CivilDrawing.GetAlignmentIds();
                            List<string> lista4 = new List<string>();
                            for (int i = col1.Count - 1; i >= 0; --i)
                            {
                                Autodesk.Civil.DatabaseServices.Alignment al1 = Trans1.GetObject(col1[i], OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Alignment;
                                if (al1 != null)
                                {
                                    if (al1.Layer != comboBox_align_layers.Text)
                                    {
                                        col1.RemoveAt(i);
                                    }
                                }
                            }


                            ThisDrawing.Editor.WriteMessage("\nprofile view band set styles:");

                            Autodesk.Civil.DatabaseServices.Styles.ProfileViewBandSetStyleCollection col2 = CivilDrawing.Styles.ProfileViewBandSetStyles;
                            ObjectId profviewsbandtyle = col2[0];
                            IEnumerator<ObjectId> enum2 = col2.GetEnumerator();

                            while (enum2.MoveNext())
                            {
                                Autodesk.Civil.DatabaseServices.Styles.ProfileViewBandSetStyle tst2 = Trans1.GetObject(enum2.Current, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.ProfileViewBandSetStyle;

                                if (tst2.Name == comboBox_profview_band_style.Text)
                                {
                                    profviewsbandtyle = tst2.ObjectId;
                                }
                                ThisDrawing.Editor.WriteMessage("\n" + tst2.Name);
                            }


                            ThisDrawing.Editor.WriteMessage("\nprofile view styles:");

                            Autodesk.Civil.DatabaseServices.Styles.ProfileViewStyleCollection col3 = CivilDrawing.Styles.ProfileViewStyles;
                            ObjectId profviewstyle = CivilDrawing.Styles.ProfileViewStyles[0];
                            IEnumerator<ObjectId> enum3 = col3.GetEnumerator();

                            double vexag = 1;

                            while (enum3.MoveNext())
                            {
                                Autodesk.Civil.DatabaseServices.Styles.ProfileViewStyle tst3 = Trans1.GetObject(enum3.Current, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.ProfileViewStyle;
                                ThisDrawing.Editor.WriteMessage("\n" + tst3.Name);
                                if (tst3.Name == comboBox_profview_style.Text)
                                {
                                    profviewstyle = tst3.ObjectId;
                                    Autodesk.Civil.DatabaseServices.Styles.GraphStyle graphstyle3 = tst3.GraphStyle;
                                    vexag = graphstyle3.VerticalExaggeration;
                                }
                            }

                            if (col1.Count > 0)
                            {


                                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify the bottom point");
                                PP1.AllowNone = false;
                                this.MdiParent.WindowState = FormWindowState.Minimized;

                                Point_res1 = Editor1.GetPoint(PP1);

                                if (Point_res1.Status != PromptStatus.OK)
                                {
                                    Editor1.SetImpliedSelection(Empty_array);
                                    Editor1.WriteMessage("\nCommand:");
                                    set_enable_true();
                                    this.MdiParent.WindowState = FormWindowState.Normal;
                                    return;
                                }

                                Point3d ptins = Point_res1.Value;


                                string layer1 = "0";
                                string layer2 = "0";

                                if (comboBox_prof1_layer.Text != "")
                                {
                                    layer1 = comboBox_prof1_layer.Text;
                                }


                                if (comboBox_prof2_layer.Text != "")
                                {
                                    layer2 = comboBox_prof2_layer.Text;
                                }

                                Autodesk.Civil.DatabaseServices.Styles.ProfileStyleCollection col_prof_styles = CivilDrawing.Styles.ProfileStyles;

                                ObjectId styleid1 = col_prof_styles[0];
                                ObjectId styleid2 = col_prof_styles[0];



                                IEnumerator<ObjectId> enum5 = col_prof_styles.GetEnumerator();

                                while (enum5.MoveNext())
                                {
                                    Autodesk.Civil.DatabaseServices.Styles.ProfileStyle tst5 = Trans1.GetObject(enum5.Current, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.ProfileStyle;
                                    if (tst5.Name == comboBox_prof_style1.Text)
                                    {
                                        styleid1 = tst5.ObjectId;
                                    }

                                    if (tst5.Name == comboBox_prof_style2.Text)
                                    {
                                        styleid2 = tst5.ObjectId;
                                    }
                                }


                                ObjectId labelsetid = CivilDrawing.Styles.LabelSetStyles.ProfileLabelSetStyles[0];


                                Autodesk.Civil.DatabaseServices.Styles.ProfileLabelSetStyleCollection col4 = CivilDrawing.Styles.LabelSetStyles.ProfileLabelSetStyles;
                                IEnumerator<ObjectId> enum4 = col4.GetEnumerator();

                                while (enum4.MoveNext())
                                {
                                    Autodesk.Civil.DatabaseServices.Styles.ProfileLabelSetStyle tst4 = Trans1.GetObject(enum4.Current, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.ProfileLabelSetStyle;
                                    if (tst4.Name == comboBox_prof_labelset_style.Text)
                                    {
                                        labelsetid = tst4.ObjectId;
                                    }
                                }




                                ObjectIdCollection col_surf = CivilDrawing.GetSurfaceIds();

                                ObjectId surfaceid1 = ObjectId.Null;
                                ObjectId surfaceid2 = ObjectId.Null;
                                string surf1_name = "";
                                string surf2_name = "";

                                for (int j = 0; j < col_surf.Count; ++j)
                                {
                                    Autodesk.Civil.DatabaseServices.Surface surf1 = Trans1.GetObject(col_surf[j], OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Surface;
                                    if (surf1 != null)
                                    {
                                        if (surf1.Name == comboBox_surf1.Text)
                                        {
                                            surfaceid2 = col_surf[j];
                                            surf1_name = surf1.Name.Replace(" ", "_");
                                        }
                                        if (surf1.Name == comboBox_surf2.Text)
                                        {
                                            surfaceid1 = col_surf[j];
                                            surf2_name = surf1.Name.Replace(" ", "_");
                                        }
                                    }
                                }

                                ObjectId layerid1 = LayerTable1[layer1];
                                ObjectId layerid2 = LayerTable1[layer2];

                                for (int i = 0; i < col1.Count; i++)
                                {
                                    Alignment align1 = Trans1.GetObject(col1[i], OpenMode.ForRead) as Alignment;



                                    if (align1 != null)
                                    {

                                        if (surfaceid1 != ObjectId.Null)
                                        {
                                            ObjectId profid1 = Profile.CreateFromSurface("prof_surf_" + surf1_name + "_align_" + align1.Name, align1.ObjectId, surfaceid1, layerid1, styleid1, labelsetid);
                                        }

                                        if (surfaceid2 != ObjectId.Null)
                                        {
                                            ObjectId profid2 = Profile.CreateFromSurface("prof_surf_" + surf2_name + "_align_" + align1.Name, align1.ObjectId, surfaceid2, layerid2, styleid2, labelsetid);
                                        }

                                        ObjectId profview1 = ProfileView.Create(align1.ObjectId, ptins, align1.Name, profviewsbandtyle, profviewstyle);

                                        ProfileView pview1 = Trans1.GetObject(profview1, OpenMode.ForRead) as ProfileView;


                                        if (pview1 != null)
                                        {
                                            //Point3d pt1 = pview1.StartPoint;
                                            Extents3d ext1 = pview1.GeometricExtents;
                                            double height1 = ext1.MaxPoint.Y - ext1.MinPoint.Y;
                                            ptins = new Point3d(ptins.X, ptins.Y + height1 + 100, 0);
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


        private void button_load_surf1_Click(object sender, EventArgs e)
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
                load_surfaces_to_combobox(comboBox_surf1);

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

        private void button_load_surf2_Click(object sender, EventArgs e)
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
                load_surfaces_to_combobox(comboBox_surf2);

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

        private void button_load_prof_view_style_Click(object sender, EventArgs e)
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
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;
                        comboBox_profview_style.Items.Clear();
                        Autodesk.Civil.DatabaseServices.Styles.ProfileViewStyleCollection col3 = CivilDrawing.Styles.ProfileViewStyles;
                        IEnumerator<ObjectId> enum3 = col3.GetEnumerator();

                        while (enum3.MoveNext())
                        {
                            Autodesk.Civil.DatabaseServices.Styles.ProfileViewStyle tst3 = Trans1.GetObject(enum3.Current, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.ProfileViewStyle;
                            comboBox_profview_style.Items.Add(tst3.Name);
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
            this.MdiParent.WindowState = FormWindowState.Normal;

            set_enable_true();
        }

        private void button_load_profileview_band_style_Click(object sender, EventArgs e)
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
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;
                        comboBox_profview_band_style.Items.Clear();
                        Autodesk.Civil.DatabaseServices.Styles.ProfileViewBandSetStyleCollection col2 = CivilDrawing.Styles.ProfileViewBandSetStyles;
                        IEnumerator<ObjectId> enum2 = col2.GetEnumerator();

                        while (enum2.MoveNext())
                        {
                            Autodesk.Civil.DatabaseServices.Styles.ProfileViewBandSetStyle tst2 = Trans1.GetObject(enum2.Current, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.ProfileViewBandSetStyle;
                            comboBox_profview_band_style.Items.Add(tst2.Name);
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
            this.MdiParent.WindowState = FormWindowState.Normal;

            set_enable_true();
            ObjectId labelsetid = CivilDrawing.Styles.LabelSetStyles.ProfileLabelSetStyles[0];
        }

        private void button_refresh_labelset_style_Click(object sender, EventArgs e)
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
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;
                        comboBox_prof_labelset_style.Items.Clear();
                        Autodesk.Civil.DatabaseServices.Styles.ProfileLabelSetStyleCollection col4 = CivilDrawing.Styles.LabelSetStyles.ProfileLabelSetStyles;
                        IEnumerator<ObjectId> enum4 = col4.GetEnumerator();

                        while (enum4.MoveNext())
                        {
                            Autodesk.Civil.DatabaseServices.Styles.ProfileLabelSetStyle tst4 = Trans1.GetObject(enum4.Current, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.ProfileLabelSetStyle;
                            comboBox_prof_labelset_style.Items.Add(tst4.Name);
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
            this.MdiParent.WindowState = FormWindowState.Normal;

            set_enable_true();

        }



        private void load_profile_styles_to_combobox(ComboBox combo1)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.Civil.ApplicationServices.CivilDocument CivilDrawing = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument;
            using (DocumentLock lock1 = ThisDrawing.LockDocument())
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;
                    combo1.Items.Clear();
                    Autodesk.Civil.DatabaseServices.Styles.ProfileStyleCollection col_prof_styles = CivilDrawing.Styles.ProfileStyles;
                    IEnumerator<ObjectId> enum5 = col_prof_styles.GetEnumerator();

                    while (enum5.MoveNext())
                    {
                        Autodesk.Civil.DatabaseServices.Styles.ProfileStyle tst5 = Trans1.GetObject(enum5.Current, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.ProfileStyle;
                        combo1.Items.Add(tst5.Name);
                    }
                    Trans1.Commit();

                }
            }
        }

        private void button_load_prof_style1_Click(object sender, EventArgs e)
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
                load_profile_styles_to_combobox(comboBox_prof_style1);
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

        private void button_load_prof_style2_Click(object sender, EventArgs e)
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
                load_profile_styles_to_combobox(comboBox_prof_style2);
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

        private void button_load_profile1_layers_Click(object sender, EventArgs e)
        {
            Functions.Incarca_existing_layers_to_combobox(comboBox_prof1_layer);
        }

        private void button_load_profile2_layers_Click(object sender, EventArgs e)
        {
            Functions.Incarca_existing_layers_to_combobox(comboBox_prof2_layer);
        }


    }
}
