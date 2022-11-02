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

    public partial class BH_mainform : Form
    {


        private bool clickdragdown;
        private Point lastLocation;


        public static HDD_Boreholes tpage_HDD_boreholes;

        public class MyRenderer : ToolStripProfessionalRenderer
        {
            public MyRenderer() : base(new MyColors()) { }
        }

        public BH_mainform()
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
           

            tpage_HDD_boreholes = new HDD_Boreholes();
            tpage_HDD_boreholes.MdiParent = this;
            tpage_HDD_boreholes.Dock = DockStyle.Fill;
            tpage_HDD_boreholes.Show();

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

        private void HDD_ToolStripMenuItem_Click(object sender, EventArgs e)
        {

            foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
            {
                if (Forma1 is Alignment_mdi.HDD_mainform)
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
                Alignment_mdi.HDD_mainform forma2 = new Alignment_mdi.HDD_mainform();
                Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma2);
                forma2.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - forma2.Width) / 2,
                     (Screen.PrimaryScreen.WorkingArea.Height - forma2.Height) / 2);
            }
            catch (System.Exception EX)
            {
                MessageBox.Show(EX.Message);
            }
           
            button_Exit_Click(sender, e);
        }

        private void BH_toolStripMenuItem_Click(object sender, EventArgs e)
        {
            tpage_HDD_boreholes.Show();
        }


    }


}
