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

    public partial class Pgen_mainform : Form
    {


        private bool clickdragdown;
        private Point lastLocation;

        public static Pgen_prof_gen tpage_prof_gen;
        public static Pgen_prof_hyd tpage_prof_hyd;
        public static pgen_vp2ms tpage_vp2ms;
        public static Pgen_prof_gen3 tpage_prof3d;
        public static bool ExcelVisible = false;


        public Pgen_mainform()
        {
            InitializeComponent();
            MenuStrip_lgen.BackColor = Color.FromArgb(37, 37, 38);
            MenuStrip_lgen.ForeColor = Color.FromArgb(0, 122, 204);
            MenuStrip_lgen.Renderer = new MyRenderer();

            foreach (Control ctrl in this.Controls)
            {
                if (ctrl is MdiClient)
                {
                    ctrl.BackColor = Color.FromArgb(62, 62, 66);
                }
            }
            tpage_prof3d = new Pgen_prof_gen3();
            tpage_prof3d.MdiParent = this;
            tpage_prof3d.Dock = DockStyle.Fill;
            tpage_prof3d.Show();

            tpage_prof_gen = new Pgen_prof_gen();
            tpage_prof_gen.MdiParent = this;
            tpage_prof_gen.Dock = DockStyle.Fill;
            tpage_prof_gen.Hide();

            tpage_prof_hyd = new Pgen_prof_hyd();
            tpage_prof_hyd.MdiParent = this;
            tpage_prof_hyd.Dock = DockStyle.Fill;
            tpage_prof_hyd.Hide();

            tpage_vp2ms = new pgen_vp2ms();
            tpage_vp2ms.MdiParent = this;
            tpage_vp2ms.Dock = DockStyle.Fill;
            tpage_vp2ms.Hide();

        }

        #region move and close
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
        #endregion

        /// <summary>
        /// Hides the ugly MDI border around the child form
        /// </summary>
        /// <param name="e"></param>
        protected override void OnLoad(EventArgs e)
        {

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

        private void Hydrant_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tpage_prof3d.Hide();
            tpage_prof_hyd.Show();
            tpage_prof_gen.Hide();
            tpage_vp2ms.Hide();
        }

        private void Stream_toolStripMenuItem_Click(object sender, EventArgs e)
        {
            tpage_prof3d.Hide();
            tpage_prof_hyd.Hide();
            tpage_prof_gen.Show();
            tpage_vp2ms.Hide();
        }

        private void toolStripMenuItem_vp2ms_Click(object sender, EventArgs e)
        {
            tpage_prof3d.Hide();
            tpage_prof_hyd.Hide();
            tpage_prof_gen.Hide();
            tpage_vp2ms.Show();
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            tpage_prof3d.Show();
            tpage_prof_hyd.Hide();
            tpage_prof_gen.Hide();
            tpage_vp2ms.Hide();
        }
    }

    public class MyRenderer : ToolStripProfessionalRenderer
    {
        public MyRenderer() : base(new MyColors()) { }
    }

    public class MyColors : ProfessionalColorTable
    {
        public override Color MenuItemSelected
        {
            get
            {
                return Color.FromArgb(0, 122, 204);
            }
        }
        public override Color MenuItemBorder
        {
            get
            {
                return Color.White;
            }
        }
    }
}
