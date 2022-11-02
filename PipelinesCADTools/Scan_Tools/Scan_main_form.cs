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

namespace Alignment_mdi
{
    public partial class scan_mainform : Form
    {
        private bool clickdragdown;
        private Point lastLocation;

        public static scan_survey_permission tpage_permission= null;
        public static map_export tpage_map_export= null;



        public scan_mainform()
        {
            InitializeComponent();




            tpage_permission = new scan_survey_permission();
            tpage_permission.MdiParent = this;
            tpage_permission.Dock = DockStyle.Fill;
            tpage_permission.Show();

            tpage_map_export = new map_export();
            tpage_map_export.MdiParent = this;
            tpage_map_export.Dock = DockStyle.Fill;
            tpage_map_export.Hide();

            //sets the mdi background color at runtime
            foreach (Control ctrl in this.Controls)
            {
                if (ctrl is MdiClient)
                {
                    ctrl.BackColor = Color.FromArgb(62, 62, 66);
                }
            }



            treeView1.ShowPlusMinus = false;

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
            this.Close();

        }
        private void button_minimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }



        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (e.Node.Text.ToLower() == "tools")
            {
                tpage_permission.Show();
                tpage_map_export.Hide();

            }



            if (e.Node.Text.ToLower() == "survey permission")
            {
                tpage_permission.Show();
                tpage_map_export.Hide();


            }

            if (e.Node.Text == "Shape Export")
            {
                tpage_permission.Hide();
                tpage_map_export.Show();

            }



        }







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

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process proc = new System.Diagnostics.Process();
            proc.StartInfo.FileName = "mailto:Support.CADTechnolgies@mottmacna.com?subject=Agen Help";
            proc.Start();
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            MessageBox.Show("Please contact Dan Popescu");
        }
    }
}
