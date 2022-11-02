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
using Font = System.Drawing.Font;



namespace Alignment_mdi
{
    public partial class Profiler_main : Form
    {
        bool clickdragdown;
        Point lastLocation;


        public static profiler_form tpage_profiler = null;
        public static cogo_points_form tpage_cogo = null;
        public static Solar_Blank_form tpage_blank = null;
        public static SolarSlope_form tpage_slope_analyze = null;
        public static Slope_modify_form tpage_slope_modify = null;
        public static tcpl_tools_form tpage_tcpl = null;


        public Profiler_main()
        {
            InitializeComponent();
           

            //sets the mdi background color at runtime
            foreach (Control ctrl in this.Controls)
            {
                if (ctrl is MdiClient)
                {
                    ctrl.BackColor = Color.FromArgb(62, 62, 66);
                }
            }



            tpage_profiler = new profiler_form();
            tpage_profiler.MdiParent = this;
            tpage_profiler.Dock = DockStyle.Fill;
            tpage_profiler.Hide();

            tpage_cogo = new cogo_points_form();
            tpage_cogo.MdiParent = this;
            tpage_cogo.Dock = DockStyle.Fill;
            tpage_cogo.Show();

            tpage_slope_analyze = new SolarSlope_form();
            tpage_slope_analyze.MdiParent = this;
            tpage_slope_analyze.Dock = DockStyle.Fill;
            tpage_slope_analyze.Hide();

            tpage_slope_modify = new Slope_modify_form();
            tpage_slope_modify.MdiParent = this;
            tpage_slope_modify.Dock = DockStyle.Fill;
            tpage_slope_modify.Hide();

            tpage_tcpl = new tcpl_tools_form();
            tpage_tcpl.MdiParent = this;
            tpage_tcpl.Dock = DockStyle.Fill;
            tpage_tcpl.Hide();


        }

        private void treeView1_BeforeSelect(object sender, TreeViewCancelEventArgs e)
        {
            if (Color.Gray == e.Node.ForeColor)
                e.Cancel = true;
        }

        private void clickmove_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            clickdragdown = true;
            lastLocation = e.Location;

        }

        private void clickmove_MouseMove(object sender, MouseEventArgs e)
        {
            if (clickdragdown == true)
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
            this.Close();
        }
        private void button_minimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void treeView_inquiry_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Node.Name == "Node1")
            {
                tpage_cogo.Hide();
                tpage_profiler.Show();
                tpage_slope_analyze.Hide();
                tpage_slope_modify.Hide();
                tpage_tcpl.Hide();
            }

            if (e.Node.Name == "Node2")
            {
                tpage_cogo.Show();
                tpage_profiler.Hide();
                tpage_slope_analyze.Hide();
                tpage_slope_modify.Hide();
                tpage_tcpl.Hide();

            }
            if (e.Node.Name == "Node3")
            {
                tpage_cogo.Hide();
                tpage_profiler.Hide();
                tpage_slope_analyze.Show();
                tpage_slope_modify.Hide();
                tpage_tcpl.Hide();

            }
            if (e.Node.Name == "Node4")
            {
                tpage_cogo.Hide();
                tpage_profiler.Hide();
                tpage_slope_analyze.Hide();
                tpage_slope_modify.Show();
                tpage_tcpl.Hide();

            }
            if (e.Node.Name == "Node5")
            {
                tpage_cogo.Hide();
                tpage_profiler.Hide();
                tpage_slope_analyze.Hide();
                tpage_slope_modify.Hide();
                tpage_tcpl.Show();

            }
        }
    }


}
