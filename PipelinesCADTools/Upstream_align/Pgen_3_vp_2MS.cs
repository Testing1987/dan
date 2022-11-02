using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Alignment_mdi
{
    public partial class pgen_vp2ms : Form
    {
        List<string> scales;


        System.Data.DataTable dt_vp;
        int extra1 = 6;


        private void set_enable_false(object sender)
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(Button_pick_vp);
            lista_butoane.Add(button_draw_poly);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                if (sender as System.Windows.Forms.Control != bt1)
                {
                    bt1.Enabled = false;
                }
            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(Button_pick_vp);
            lista_butoane.Add(button_draw_poly);
            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }

        public pgen_vp2ms()
        {
            InitializeComponent();
            dt_vp = new System.Data.DataTable();
            dt_vp.Columns.Add("nw", typeof(Point3d));
            dt_vp.Columns.Add("ne", typeof(Point3d));
            dt_vp.Columns.Add("se", typeof(Point3d));
            dt_vp.Columns.Add("sw", typeof(Point3d));
            dt_vp.Columns.Add("dwg", typeof(string));

        }

        private void make_first_line_invisible()
        {


            textBox1.Visible = false;


            for (int i = panel_err.Controls.Count - 1; i >= 0; --i)
            {
                Control ctrl1 = panel_err.Controls[i] as Control;
                if (ctrl1.Location.Y > textBox1.Location.Y + extra1)
                {
                    panel_err.Controls.Remove(ctrl1);
                    ctrl1.Dispose();
                }
            }
            //textBox_PM_no_rows.Text = "";
            //textBox_PM_no_duplicates.Text = "";
            //textBox_PM_no_null.Text = "";
            //textBox_DJ_no_rows.Text = "";
            //textBox_DJ_no_duplicates.Text = "";
            //textBox_DJ_no_null.Text = "";
        }

        private void Button_pick_vp_Click(object sender, EventArgs e)
        {

            make_first_line_invisible();

            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {
                set_enable_false(sender);
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord_MS = Trans1.GetObject(BlockTable1[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;
                        BlockTableRecord BTrecord_PS = Trans1.GetObject(BlockTable1[BlockTableRecord.PaperSpace], OpenMode.ForRead) as BlockTableRecord;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;


                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_Viewport;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_viewport;
                        Prompt_viewport = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the viewport:");
                        Prompt_viewport.SetRejectMessage("\nSelect a viewport!");
                        Prompt_viewport.AllowNone = true;
                        Prompt_viewport.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Viewport), false);
                        Prompt_viewport.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);

                        Rezultat_Viewport = ThisDrawing.Editor.GetEntity(Prompt_viewport);

                        if (Rezultat_Viewport.Status != PromptStatus.OK)
                        {

                            set_enable_true();

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }





                        Viewport Ent_vp = Trans1.GetObject(Rezultat_Viewport.ObjectId, OpenMode.ForRead) as Viewport;
                        if (Ent_vp == null)
                        {
                            Polyline Ent_poly = Trans1.GetObject(Rezultat_Viewport.ObjectId, OpenMode.ForRead) as Polyline;
                            if (Ent_poly != null)
                            {
                                ObjectId vpId = LayoutManager.Current.GetNonRectangularViewportIdFromClipId(Rezultat_Viewport.ObjectId);
                                if (Trans1.GetObject(vpId, OpenMode.ForRead) is Viewport)
                                {
                                    Ent_vp = Trans1.GetObject(vpId, OpenMode.ForRead) as Viewport;
                                }

                            }
                        }

                        if (Ent_vp != null)
                        {
                            double h1 = Ent_vp.Height;
                            double l1 = Ent_vp.Width;
                            Point3d cen1 = Ent_vp.CenterPoint;

                            dt_vp.Rows.Add();

                            Point3d nw_ps = new Point3d(cen1.X - l1 / 2, cen1.Y + h1 / 2, 0);
                            dt_vp.Rows[dt_vp.Rows.Count - 1][0] = nw_ps.TransformBy(Functions.PaperToModel(Ent_vp));

                            Point3d ne_ps = new Point3d(cen1.X + l1 / 2, cen1.Y + h1 / 2, 0);
                            dt_vp.Rows[dt_vp.Rows.Count - 1][1] = ne_ps.TransformBy(Functions.PaperToModel(Ent_vp));

                            Point3d se_ps = new Point3d(cen1.X + l1 / 2, cen1.Y - h1 / 2, 0);
                            dt_vp.Rows[dt_vp.Rows.Count - 1][2] = se_ps.TransformBy(Functions.PaperToModel(Ent_vp));

                            Point3d sw_ps = new Point3d(cen1.X - l1 / 2, cen1.Y - h1 / 2, 0);
                            dt_vp.Rows[dt_vp.Rows.Count - 1][3] = sw_ps.TransformBy(Functions.PaperToModel(Ent_vp));

                            dt_vp.Rows[dt_vp.Rows.Count - 1][4] = System.IO.Path.GetFileNameWithoutExtension(ThisDrawing.Database.OriginalFileName);

                        }


                        transfer_dwg_to_panel(dt_vp);

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

        private void transfer_dwg_to_panel(System.Data.DataTable dt1)
        {
            try
            {
                if (dt1.Rows.Count > 0)
                {
                    textBox1.Visible = true;
                   

                    for (int i = panel_err.Controls.Count - 1; i >= 0; --i)
                    {
                        Control ctrl1 = panel_err.Controls[i] as Control;
                        if (ctrl1.Location.Y > textBox1.Location.Y + extra1)
                        {
                            panel_err.Controls.Remove(ctrl1);
                            ctrl1.Dispose();
                        }
                    }

                    string text1 = "";

                    if (dt1.Rows[0][4] != DBNull.Value)
                    {
                        text1 = Convert.ToString(dt1.Rows[0][4]);
                    }

                    textBox1.Text = text1;
                 


                    if (dt1.Rows.Count > 1)
                    {
                        for (int i = 1; i < dt1.Rows.Count; ++i)
                        {
                            string text11 = "";

                            if (dt1.Rows[i][4] != DBNull.Value)
                            {
                                text11 = Convert.ToString(dt1.Rows[i][4]);

                            }

                            TextBox tb1 = new TextBox();
                            tb1.Location = new Point(textBox1.Location.X, textBox1.Location.Y + i * (textBox1.Height + extra1));
                            tb1.BackColor = textBox1.BackColor;
                            tb1.ForeColor = textBox1.ForeColor;
                            tb1.Font = textBox1.Font;
                            tb1.Size = textBox1.Size;
                            tb1.ReadOnly = textBox1.ReadOnly;
                            tb1.BorderStyle = textBox1.BorderStyle;
                            tb1.Text = text11;
                            panel_err.Controls.Add(tb1);




                        }
                    }
                }
            }
            catch (System.ComponentModel.Win32Exception ex1)
            {
                MessageBox.Show(ex1.Message);
            }
        }

        private void button_draw_poly_Click(object sender, EventArgs e)
        {
            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {
                set_enable_false(sender);
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(BlockTable1[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                        if (dt_vp.Rows.Count > 0)
                        {
                            Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                            Functions.Create_vp_od_table();

                            Functions.Creaza_layer("CrossingsVP", 7, true);
                            for (int i = 0; i < dt_vp.Rows.Count; ++i)
                            {
                                Point3d pt1 = (Point3d)dt_vp.Rows[i][0];
                                Point3d pt2 = (Point3d)dt_vp.Rows[i][1];
                                Point3d pt3 = (Point3d)dt_vp.Rows[i][2];
                                Point3d pt4 = (Point3d)dt_vp.Rows[i][3];

                                Polyline poly1 = new Polyline();
                                poly1.Layer = "CrossingsVP";
                                poly1.ColorIndex = 256;


                                poly1.AddVertexAt(0, new Point2d(pt1.X, pt1.Y), 0, 0, 0);
                                poly1.AddVertexAt(1, new Point2d(pt2.X, pt2.Y), 0, 0, 0);
                                poly1.AddVertexAt(2, new Point2d(pt3.X, pt3.Y), 0, 0, 0);
                                poly1.AddVertexAt(3, new Point2d(pt4.X, pt4.Y), 0, 0, 0);
                                poly1.Elevation = 0;
                                poly1.Closed = true;

                                BTrecord.AppendEntity(poly1);
                                Trans1.AddNewlyCreatedDBObject(poly1, true);

                                string dwg = Convert.ToString(dt_vp.Rows[i][4]);
                                List<object> Lista_val = new List<object>();
                                List<Autodesk.Gis.Map.Constants.DataType> Lista_type = new List<Autodesk.Gis.Map.Constants.DataType>();

                                Lista_val.Add(dwg);
                                Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);


                                Functions.Populate_object_data_table_from_objectid(Tables1, poly1.ObjectId, "PGEN_VP", Lista_val, Lista_type);
                            }
                        }


                        dt_vp.Rows.Clear();
                        make_first_line_invisible();

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
    }
}
