using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Alignment_mdi;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System.Management;



namespace Alignment_mdi
{
    public partial class _SGEN_mainform : Form
    {
        private bool clickdragdown;
        private Point lastLocation;
        //public static Matl_Design_Tool tpage_matl = null;
        //public static Blank_Form tpage_blank = null;

        public static _SGEN_mainform tpage_dsuite = null;
        public static _SGEN_mainform tpage_Main = null;
        public static SGEN_Sheet_Index tpage_sheetindex = null;
        public static Blank_Form tpage_blank = null;
        public static AGEN_TBLK_Attributes tpage_TBLK = null;
        public static Geo_tools_form tpage_geomanager = null;
        public static Lgen_label_page tpage_Labels = null;
        public static SGEN_Shape_Export tpage_shape = null;
        public static SGEN_Settings tpage_settings = null;
        public static SGEN_Drawing_Creation tpage_dc = null;


        public Color Foreground_Main_Blue = new Color();
        public Color Backgournd_Dark = new Color();


        public static string units_of_measurement = "f";
        public static System.Data.DataTable dt_sheet_index;
        public static string sheet_index_excel_name = "sheet_index.xlsx";
        public static int Start_row_Sheet_index = 11;

        public static int round1 = 0;
     
        public static string config_file = "";
        public static string output_folder = "";

        public static System.Data.DataTable Data_Table_regular_bands;
        public static System.Data.DataTable Data_Table_display_bands;

        public static string nume_main_vp = "Plan View";

        public static double Vw_scale = 1;

        public static double Vw_height = 0;
        public static double Vw_width = 0;
        public static double Vw_ps_x = 0;
        public static double Vw_ps_y = 0;
        public static double Band_Separation = 1000;
        public static string Layer_name_ML_rectangle = "AGEN_Index_ML";

        public static string Col_dwg_name = "DwgNo";
        public static string Layer_name_Main_Viewport = "AGEN_mainVP";
        public static string Col_x = "X";
        public static string Col_y = "Y";
        public static string Col_rot = "Rotation";
        public static string block_attributes_excel_name = "TBLK_attributes.xlsx";
        public static int Start_row_block_attributes = 7;

        public static string od_table_sheet_index = "SheetIndex";
        public static string project_main_folder = "";


        public static int no_of_segments =0;

        public static System.Data.DataTable dt_segments;


        public _SGEN_mainform()
        {
            InitializeComponent();


            Foreground_Main_Blue = Color.FromArgb(0, 122, 204);
            Foreground_Main_Blue = Color.FromArgb(37, 37, 38);

            tpage_blank = new Blank_Form();
            tpage_blank.MdiParent = this;
            tpage_blank.Dock = DockStyle.Fill;
            tpage_blank.Show();

            tpage_settings = new SGEN_Settings();
            tpage_settings.MdiParent = this;
            tpage_settings.Dock = DockStyle.Fill;
            tpage_settings.Hide();

            tpage_dc = new SGEN_Drawing_Creation();
            tpage_dc.MdiParent = this;
            tpage_dc.Dock = DockStyle.Fill;
            tpage_dc.Hide();

            tpage_sheetindex = new SGEN_Sheet_Index();
            tpage_sheetindex.MdiParent = this;
            tpage_sheetindex.Dock = DockStyle.Fill;
            if (Functions.is_dan_popescu() == true)
            {
                tpage_sheetindex.make_labels_visible();
            }
            tpage_sheetindex.Hide();


            tpage_TBLK = new AGEN_TBLK_Attributes();
            tpage_TBLK.MdiParent = this;
            tpage_TBLK.Dock = DockStyle.Fill;
            tpage_TBLK.Hide();

            tpage_geomanager = new Geo_tools_form();
            tpage_geomanager.MdiParent = this;
            tpage_geomanager.Dock = DockStyle.Fill;
            tpage_geomanager.Hide();

            tpage_Labels = new Lgen_label_page();
            tpage_Labels.MdiParent = this;
            tpage_Labels.Dock = DockStyle.Fill;
            tpage_Labels.Hide();

            tpage_shape = new SGEN_Shape_Export();
            tpage_shape.MdiParent = this;
            tpage_shape.Dock = DockStyle.Fill;
            tpage_shape.Hide();

            this.FormBorderStyle = FormBorderStyle.None;
            this.DoubleBuffered = true;
            this.SetStyle(ControlStyles.ResizeRedraw, true);
            tpage_Main = this;

            //sets the mdi background color at runtime
            foreach (Control ctrl in this.Controls)
            {
                if (ctrl is MdiClient)
                {
                    ctrl.BackColor = Color.FromArgb(62, 62, 66);
                }
            }

            

        }


        #region resize window
        private const int cGrip = 16;      // Grip size
        private const int cCaption = 32;   // Caption bar height;

        protected override void OnPaint(PaintEventArgs e)
        {
            System.Drawing.Rectangle rc = new System.Drawing.Rectangle(this.ClientSize.Width - cGrip, this.ClientSize.Height - cGrip, cGrip, cGrip);
            ControlPaint.DrawSizeGrip(e.Graphics, this.BackColor, rc);
            rc = new System.Drawing.Rectangle(0, 0, this.ClientSize.Width, cCaption);
            e.Graphics.FillRectangle(Brushes.Transparent, rc);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x84)
            {  // Trap WM_NCHITTEST
                System.Drawing.Point pos = new System.Drawing.Point(m.LParam.ToInt32());
                pos = this.PointToClient(pos);
                if (pos.Y < cCaption)
                {
                    m.Result = (IntPtr)2;  // HTCAPTION
                    return;
                }
                if (pos.X >= this.ClientSize.Width - cGrip && pos.Y >= this.ClientSize.Height - cGrip)
                {
                    m.Result = (IntPtr)17; // HTBOTTOMRIGHT
                    return;
                }
            }
            base.WndProc(ref m);
        }
        #endregion

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



        #region Window Controls
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
            dt_segments = null;
            dt_sheet_index = null;
            this.Close();
           
        }
        private void button_minimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }
        #endregion

        #region Treeview Stuff

        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            var hitTest = e.Node.TreeView.HitTest(e.Location);
            if (hitTest.Location == TreeViewHitTestLocations.PlusMinus)
                return;

            if (e.Node.IsExpanded)
                e.Node.Collapse();
            else
                e.Node.Expand();
        }


        #endregion





        #region Navigation
        private void treeView1_AfterSelect_1(object sender, TreeViewEventArgs e)
        {

            //control page display
            switch (e.Node.Text)
            {
                case "Sheet Cutter":
                    tpage_sheetindex.Hide();
                    tpage_dc.Hide();
                    tpage_blank.Show();
                    tpage_TBLK.Hide();
                    tpage_geomanager.Hide();
                    tpage_Labels.Hide();
                    tpage_shape.Hide();
                    tpage_settings.Hide();
                    break;

                case "Project Setup":
                    tpage_sheetindex.Hide();
                    tpage_dc.Hide();
                    tpage_blank.Hide();
                    tpage_TBLK.Hide();
                    tpage_geomanager.Hide();
                    tpage_Labels.Hide();
                    tpage_shape.Hide();
                    tpage_settings.Show();
                    break;

                case "Sheet Index Setup":
                    tpage_sheetindex.Show();
                    tpage_dc.Hide();
                    tpage_blank.Hide();
                    tpage_TBLK.Hide();
                    tpage_geomanager.Hide();
                    tpage_Labels.Hide();
                    tpage_shape.Hide();
                    tpage_settings.Hide();
                    break;

                case "Drawing Creation":
                    tpage_sheetindex.Hide();
                    tpage_dc.Show();
                    tpage_blank.Hide();
                    tpage_TBLK.Hide();
                    tpage_geomanager.Hide();
                    tpage_Labels.Hide();
                    tpage_shape.Hide();
                    tpage_settings.Hide();
                    break;

                case "Titleblock Manager":
                    tpage_sheetindex.Hide();
                    tpage_dc.Hide();
                    tpage_blank.Hide();
                    tpage_TBLK.Show();
                    tpage_geomanager.Hide();
                    tpage_Labels.Hide();
                    tpage_shape.Hide();
                    tpage_settings.Hide();
                    break;

                case "Data Manager":
                    tpage_sheetindex.Hide();
                    tpage_dc.Hide();
                    tpage_blank.Hide();
                    tpage_TBLK.Hide();
                    tpage_geomanager.Show();
                    tpage_Labels.Hide();
                    tpage_shape.Hide();
                    tpage_settings.Hide();
                    break;

                case "Label Generator":
                    tpage_sheetindex.Hide();
                    tpage_dc.Hide();
                    tpage_blank.Hide();
                    tpage_TBLK.Hide();
                    tpage_geomanager.Hide();
                    tpage_Labels.Show();
                    tpage_shape.Hide();
                    tpage_settings.Hide();
                    break;

                case "Shape Export":
                    tpage_sheetindex.Hide();
                    tpage_dc.Hide();
                    tpage_blank.Hide();
                    tpage_TBLK.Hide();
                    tpage_geomanager.Hide();
                    tpage_Labels.Hide();
                    tpage_shape.Show();
                    tpage_settings.Hide();
                    break;

                default:
                    break;
            }
        }

        #endregion




        public void LGEN_label_excel_file_red()
        {
            label_excel_info.Text = "Not loaded";
            label_excel_info.ForeColor = Color.OrangeRed;

        }

        public void LGEN_label_excel_file_green(string Fisier_layer_alias)
        {
            label_excel_info.Text = System.IO.Path.GetFileName(Fisier_layer_alias);
            label_excel_info.ForeColor = Color.LimeGreen;
        }

    }
}
