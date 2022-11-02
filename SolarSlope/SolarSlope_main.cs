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
    public partial class Solar_main_form : Form
    {
        bool clickdragdown;
        Point lastLocation;


        public static SolarSlope_form tpage_slope = null;
        public static Solar_Blank_form tpage_blank = null;





        public Solar_main_form()
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



            tpage_slope = new SolarSlope_form();
            tpage_slope.MdiParent = this;
            tpage_slope.Dock = DockStyle.Fill;
            tpage_slope.Show();


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

        private void treeView_inquiry_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (e.Node.Text == "Slope Analizer")
            {
                tpage_slope.Show();
              


            }


        }
    }


}
