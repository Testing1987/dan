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

namespace Alignment_mdi
{
    public partial class Igen_main_form : Form
    {
        bool clickdragdown;
        Point lastLocation;

        static public List<string> col_labels_zoom;

        static public List<string> col_station_labels;

        public static Igen_Inquiry_Tool tpage_inquiry = null;
        public static Wgen_Blank_form tpage_blank = null;
        public static Igen_geomanager tpage_geomanager = null;

        public static string cl_id_for_temp = null;

        public Igen_main_form()
        {
            InitializeComponent();
            col_labels_zoom = new List<string>();
            col_station_labels = new List<string>();

            tpage_inquiry = new Igen_Inquiry_Tool();
            tpage_inquiry.MdiParent = this;
            tpage_inquiry.Dock = DockStyle.Fill;
            tpage_inquiry.Hide();

            tpage_geomanager = new Igen_geomanager();
            tpage_geomanager.MdiParent = this;
            tpage_geomanager.Dock = DockStyle.Fill;
            tpage_geomanager.Hide();

            tpage_blank = new Wgen_Blank_form();
            tpage_blank.MdiParent = this;
            tpage_blank.Dock = DockStyle.Fill;
            tpage_blank.Show();

            //sets the mdi background color at runtime
            foreach (Control ctrl in this.Controls)
            {
                if (ctrl is MdiClient)
                {
                    ctrl.BackColor = Color.FromArgb(62, 62, 66);
                }
            }




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
            tpage_inquiry.delete_zoom_labels();
            if (tpage_inquiry.get_checkBox_temp_cl() == true)
            {
                tpage_inquiry.delete_cl_from_redraw(cl_id_for_temp);
            }

            if (tpage_inquiry.get_checkBox_temp_sta() == true)
            {
                tpage_inquiry.delete_station_labels();
            }

            this.Close();
        }
        private void button_minimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }




        private void treeView_inquiry_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (e.Node.Text == "Inquiry Tools")
            {
                tpage_inquiry.Show();
                tpage_geomanager.Hide();
                tpage_blank.Hide();
            }
            if (e.Node.Text == "Object Data Tables")
            {
                tpage_inquiry.Hide();
                tpage_geomanager.Show();
                tpage_blank.Hide();
            }


        }

        private void label_iq_treeviewnav_Click(object sender, EventArgs e)
        {
            tpage_inquiry.Hide();
            tpage_geomanager.Hide();
            tpage_blank.Show();
        }
    }
}
