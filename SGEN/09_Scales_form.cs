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


namespace Alignment_mdi
{
    public partial class scales_form : Form
    {
        bool clickdragdown;
        Point lastLocation;


        public scales_form()
        {
            InitializeComponent();

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



        private void button_close_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button_minimize_Click_1(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        public double get_current_scale()
        {
            double scale1 = 1;
            if (radioButton1.Checked == true) scale1 = 1;
            if (radioButton10.Checked == true) scale1 = 10;
            if (radioButton20.Checked == true) scale1 = 20;
            if (radioButton30.Checked == true) scale1 = 30;
            if (radioButton40.Checked == true) scale1 = 40;
            if (radioButton50.Checked == true) scale1 = 50;
            if (radioButton60.Checked == true) scale1 = 60;
            if (radioButton100.Checked == true) scale1 = 100;
            if (radioButton200.Checked == true) scale1 = 200;
            if (radioButton300.Checked == true) scale1 = 300;
            if (radioButton400.Checked == true) scale1 = 400;
            if (radioButton500.Checked == true) scale1 = 500;
            if (radioButton600.Checked == true) scale1 = 600;
            if (radioButton1000.Checked == true) scale1 = 1000;
            if (radioButton2000.Checked == true) scale1 = 2000;
            if (radioButton3000.Checked == true) scale1 = 3000;
            if (radioButton4000.Checked == true) scale1 = 4000;
            if (radioButton5000.Checked == true) scale1 = 5000;
            if (radioButton6000.Checked == true) scale1 = 6000;

            return 1 / scale1;
        }
        public string get_current_scale_name()
        {
            string scale1 = "1:1";
            if (radioButton1.Checked == true) return "1:1";
            if (radioButton10.Checked == true) return "1:10";
            if (radioButton20.Checked == true) return "1:20"; 
            if (radioButton30.Checked == true) return "1:30";
            if (radioButton40.Checked == true) return "1:40";
            if (radioButton50.Checked == true) return "1:50";
            if (radioButton60.Checked == true) return "1:60";
            if (radioButton100.Checked == true) return "1:100";
            if (radioButton200.Checked == true) return "1:200";
            if (radioButton300.Checked == true) return "1:300";
            if (radioButton400.Checked == true) return "1:400";
            if (radioButton500.Checked == true) return "1:500";
            if (radioButton600.Checked == true) return "1:600";
            if (radioButton1000.Checked == true) return "1:1000";
            if (radioButton2000.Checked == true) return "1:2000";
            if (radioButton3000.Checked == true) return "1:3000";
            if (radioButton4000.Checked == true) return "1:4000";
            if (radioButton5000.Checked == true) return "1:5000";
            if (radioButton6000.Checked == true) return "1:6000";

            return  scale1;
        }

    }
}
