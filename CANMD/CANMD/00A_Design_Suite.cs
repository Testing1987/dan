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
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;

namespace Alignment_mdi
{
    public partial class ds_main : Form
    {
        private bool clickdragdown;
        private Point lastLocation;
        public static Mat_Design_form tpage_mat_design = null;
        public static Centerline_form tpage_centerline = null;

        public static Export_form tpage_export = null;
        public static Blank_Form tpage_blank = null;
        public static form_transitions tpage_transition = null;

        public static System.Data.DataTable dt_centerline = null;
        public static System.Data.DataTable dt_top = null;

        public static ds_main tpage_main = null;

        public static bool is_usa = true;
        public static bool is3D = false;
        public static string mat_elbow = "";
        public static string col_Cat = "Category";

        private void make_variables_null()
        {
            dt_centerline = null;
            dt_top = null;

            mat_elbow = "";
            is3D = false;
            is_usa = true;
            Mat_Design_form.dt_mat_library = null;
            Mat_Design_form.dt_filter = null;
            Mat_Design_form.category_list = null;
        }

        public ds_main()
        {
            InitializeComponent();
            tpage_main = this;
            textBox_client_name.Focus();

            tpage_mat_design = new Mat_Design_form();
            tpage_mat_design.MdiParent = this;
            tpage_mat_design.Dock = DockStyle.Fill;
            tpage_mat_design.Hide();

            tpage_centerline = new Centerline_form();
            tpage_centerline.MdiParent = this;
            tpage_centerline.Dock = DockStyle.Fill;
            tpage_centerline.Hide();

            tpage_export = new Export_form();
            tpage_export.MdiParent = this;
            tpage_export.Dock = DockStyle.Fill;
            tpage_export.Hide();

            tpage_transition = new form_transitions();
            tpage_transition.MdiParent = this;
            tpage_transition.Dock = DockStyle.Fill;
            tpage_transition.Hide();

            tpage_blank = new Blank_Form();
            tpage_blank.MdiParent = this;
            tpage_blank.Dock = DockStyle.Fill;
            tpage_blank.Show();


            //sets the mdi background color at runtime
            foreach (Control ctrl in this.Controls)
            {
                if (ctrl is MdiClient)
                {
                    ctrl.BackColor = Color.FromArgb(37, 37, 38);
                }
            }

            treeView1.ShowPlusMinus = false;
        }

        [CommandMethod("CANMD")]
        public void ShowForm()
        {
            if (Functions.isSECURE() == true)
            {
                foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
                {
                    if (Forma1 is Alignment_mdi.ds_main)
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
                    Alignment_mdi.ds_main forma2 = new Alignment_mdi.ds_main();
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

        [CommandMethod("CALC_DOC")]
        public void CALC_DEPTH_OF_COVER()
        {
            if (Functions.isSECURE() == true)
            {


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
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;


                            Autodesk.AutoCAD.EditorInput.PromptEntityResult Rez_CL;
                            Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_CL;
                            Prompt_CL = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the PIPE centerline:");
                            Prompt_CL.SetRejectMessage("\nSelect a polyline!");
                            Prompt_CL.AllowNone = true;
                            Prompt_CL.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                            Rez_CL = ThisDrawing.Editor.GetEntity(Prompt_CL);

                            if (Rez_CL.Status != PromptStatus.OK)
                            {
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }



                            Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_ground;
                            Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_ground;
                            Prompt_ground = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the ground polyline:");
                            Prompt_ground.SetRejectMessage("\nSelect a polyline!");
                            Prompt_ground.AllowNone = true;
                            Prompt_ground.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                            Rezultat_ground = ThisDrawing.Editor.GetEntity(Prompt_ground);

                            if (Rezultat_ground.Status != PromptStatus.OK)
                            {

                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }


                            Polyline cl = Trans1.GetObject(Rez_CL.ObjectId, OpenMode.ForRead) as Polyline;
                            Polyline grd = Trans1.GetObject(Rezultat_ground.ObjectId, OpenMode.ForRead) as Polyline;

                            if (cl != null && grd != null)
                            {
                                using (System.Data.DataTable dt1 = new System.Data.DataTable())
                                {
                                    dt1.Columns.Add("sta", typeof(double));
                                    dt1.Columns.Add("elev", typeof(double));
                                    dt1.Columns.Add("cover", typeof(double));
                                    for (int i = 0; i < cl.NumberOfVertices; ++i)
                                    {
                                        dt1.Rows.Add();
                                        dt1.Rows[dt1.Rows.Count - 1][0] = cl.GetPointAtParameter(i).X;
                                        dt1.Rows[dt1.Rows.Count - 1][1] = cl.GetPointAtParameter(i).Y;

                                        Polyline poly1 = new Polyline();
                                        poly1.AddVertexAt(0, cl.GetPoint2dAt(i), 0, 0, 0);
                                        poly1.AddVertexAt(1, new Point2d(cl.GetPoint2dAt(i).X, cl.GetPoint2dAt(i).Y + 100), 0, 0, 0);
                                        Point3dCollection colint = Functions.Intersect_on_both_operands(poly1, grd);

                                        if (colint.Count > 0)
                                        {
                                            dt1.Rows[dt1.Rows.Count - 1][2] = colint[0].Y - cl.GetPointAtParameter(i).Y;
                                        }


                                    }

                                    for (int i = 0; i < grd.NumberOfVertices; ++i)
                                    {

                                        Polyline poly1 = new Polyline();
                                        poly1.AddVertexAt(0, grd.GetPoint2dAt(i), 0, 0, 0);
                                        poly1.AddVertexAt(1, new Point2d(grd.GetPoint2dAt(i).X, grd.GetPoint2dAt(i).Y - 100), 0, 0, 0);
                                        Point3dCollection colint = Functions.Intersect_on_both_operands(poly1, cl);

                                        if (colint.Count > 0)
                                        {
                                            dt1.Rows.Add();
                                            dt1.Rows[dt1.Rows.Count - 1][0] = grd.GetPointAtParameter(i).X;
                                            dt1.Rows[dt1.Rows.Count - 1][1] = colint[0].Y;
                                            dt1.Rows[dt1.Rows.Count - 1][2] = -colint[0].Y + grd.GetPointAtParameter(i).Y;
                                        }


                                    }
                                    using (System.Data.DataTable dt2 = Functions.Sort_data_table(dt1, "sta"))
                                    {
                                        Functions.Transfer_datatable_to_new_excel_spreadsheet(dt2);
                                    }
                                }


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
            var size = new Size(mdiclient.Width + hdiff, mdiclient.Height + vdiff);
            var location = new Point(mdiclient.Left - (hdiff / 2), mdiclient.Top - (vdiff / 2));
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
                this.Location = new Point(
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


        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            switch (e.Node.Name)
            {
                case "Node1":
                    tpage_mat_design.Hide();
                    tpage_centerline.Show();
                    tpage_export.Hide();
                    tpage_blank.Hide();
                    tpage_transition.Hide();
                    break;
                case "Node2":
                    tpage_mat_design.Show();
                    tpage_centerline.Hide();
                    tpage_export.Hide();
                    tpage_blank.Hide();
                    tpage_transition.Hide();
                    break;
                case "Node3":
                    tpage_mat_design.Hide();
                    tpage_centerline.Hide();
                    tpage_export.Show();
                    tpage_blank.Hide();
                    tpage_transition.Hide();
                    break;
                case "Node4":
                    tpage_mat_design.Hide();
                    tpage_centerline.Hide();
                    tpage_export.Hide();
                    tpage_blank.Hide();
                    tpage_transition.Show();
                    break;
                default:
                    tpage_mat_design.Hide();
                    tpage_centerline.Hide();
                    tpage_export.Hide();
                    tpage_transition.Hide();
                    tpage_blank.Show();
                    break;
            }


        }







        public string get_textbox_client_name()
        {
            return textBox_client_name.Text;
        }

        public string get_textbox_pipe_diam()
        {
            return textBox_pipe_diam.Text;
        }

        public string get_textbox_project()
        {
            return textBox_project.Text;
        }
        public string get_textbox_segment()
        {
            return textBox_segment.Text;
        }

        public void set_textbox_client_name(string continut)
        {
            textBox_client_name.Text = continut;
        }

        public void set_textbox_pipe_diam(string continut)
        {
            textBox_pipe_diam.Text = continut;
        }

        public void set_textbox_project(string continut)
        {
            textBox_project.Text = continut;
        }
        public void set_textbox_segment(string continut)
        {
            textBox_segment.Text = continut;
        }

    }
}
