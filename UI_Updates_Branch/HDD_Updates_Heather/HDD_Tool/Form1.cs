using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HDD_Tool
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        private void RadioButton1_CheckedChanged(object sender, EventArgs e)
        {
           
            if (radioButton2.Checked == true)
            {
                panel1.BackgroundImage = Properties.Resources.HDD_Right_to_Left;
            }
            else
            {
                panel1.BackgroundImage = Properties.Resources.HDD_Left_to_Right;
            }

            if (radioButton2.Checked == true)
            {
                button_load_ground_profile.Location = new Point(3, 3);
                button_load_ground_profile.Size = new Size(230, 28);
            }
            else
            {
                button_load_ground_profile.Location = new Point(3, 3);
                button_load_ground_profile.Size = new Size(230, 28);
            }

            if (radioButton2.Checked == true)
            {
                button_calc_HDD.Location = new Point(3, 35);
                button_calc_HDD.Size = new Size(230, 28);
            }
            else
            {
                button_calc_HDD.Location = new Point(3, 35);
                button_calc_HDD.Size = new Size(230, 28);
            }

            if (radioButton2.Checked == true)
            {
                button_load_HDD_design.Location = new Point(3, 67);
            }
            else
            {
                button_load_HDD_design.Location = new Point(3, 67);
            }

            if (radioButton2.Checked == true)
            {
                panel19.Location = new Point(1214, 105);
            }
            else
            {
                panel19.Location = new Point(14, 105);
            }

            if (radioButton2.Checked == true)
            {
                panel4.Location = new Point(1170, 231);
            }
            else
            {
                panel4.Location = new Point(134, 231);
            }

            if (radioButton2.Checked == true)
            {
                label28.Location = new Point(1085, 148);
            }
            else
            {
                label28.Location = new Point(321, 148);
            }

            if (radioButton2.Checked == true)
            {
                textBox_L1.Location = new Point(1073, 165);
            }
            else
            {
                textBox_L1.Location = new Point(309, 165);
            }

            if (radioButton2.Checked == true)
            {
                panel16.Location = new Point(1100, 295);
            }
            else
            {
                panel16.Location = new Point(127, 295);
            }

            if (radioButton2.Checked == true)
            {
                label66.Location = new Point(1011, 245);
            }
            else
            {
                label66.Location = new Point(406, 245);
            }

            if (radioButton2.Checked == true)
            {
                textBox_h1.Location = new Point(980, 262);
            }
            else
            {
                textBox_h1.Location = new Point(403, 262);
            }

            if (radioButton2.Checked == true)
            {
                panel5.Location = new Point(926, 370);
            }
            else
            {
                panel5.Location = new Point(341, 370);
            }

            if (radioButton2.Checked == true)
            {
                label17.Location = new Point(903, 264);
            }
            else
            {
                label17.Location = new Point(513, 264);
            }

            if (radioButton2.Checked == true)
            {
                textBox_L2.Location = new Point(878, 281);
            }
            else
            {
                textBox_L2.Location = new Point(485, 281);
            }

            if (radioButton2.Checked == true)
            {
                label65.Location = new Point(818, 316);
            }
            else
            {
                label65.Location = new Point(593, 316);
            }

            if (radioButton2.Checked == true)
            {
                textBox_h2.Location = new Point(818, 334);
            }
            else
            {
                textBox_h2.Location = new Point(560, 334);
            }

            if (radioButton2.Checked == true)
            {
                panel15.Location = new Point(771, 439);
            }
            else
            {
                panel15.Location = new Point(448, 439);
            }

            if (radioButton2.Checked == true)
            {
                textBox_dev_angle2.Location = new Point(741, 140);
            }
            else
            {
                textBox_dev_angle2.Location = new Point(653, 140);
            }

            if (radioButton2.Checked == true)
            {
                label1.Location = new Point(722, 163);
            }
            else
            {
                label1.Location = new Point(653, 140);
            }

            if (radioButton2.Checked == true)
            {
                label41.Location = new Point(694, 242);
            }
            else
            {
                label41.Location = new Point(712, 242);
            }

            if (radioButton2.Checked == true)
            {
                textBox_L3.Location = new Point(685, 262);
            }
            else
            {
                textBox_L3.Location = new Point(703, 262);
            }

            if (radioButton2.Checked == true)
            {
                panel7.Location = new Point(664, 395);
            }
            else
            {
                panel7.Location = new Point(712, 395);
            }

            if (radioButton2.Checked == true)
            {
                label63.Location = new Point(619, 308);
            }
            else
            {
                label63.Location = new Point(793, 308);
            }

            if (radioButton2.Checked == true)
            {
                textBox_h4.Location = new Point(600, 324);
            }
            else
            {
                textBox_h4.Location = new Point(795, 324);
            }

            if (radioButton2.Checked == true)
            {
                label40.Location = new Point(570, 264);
            }
            else
            {
                label40.Location = new Point(845, 264);
            }

            if (radioButton2.Checked == true)
            {
                textBox_L4.Location = new Point(554, 281);
            }
            else
            {
                textBox_L4.Location = new Point(829, 281);
            }

            if (radioButton2.Checked == true)
            {
                panel9.Location = new Point(429, 445);
            }
            else
            {
                panel9.Location = new Point(848, 445);
            }

            if (radioButton2.Checked == true)
            {
                panel6.Location = new Point(408, 374);
            }
            else
            {
                panel6.Location = new Point(891, 374);
            }

            if (radioButton2.Checked == true)
            {
                label64.Location = new Point(458, 228);
            }
            else
            {
                label64.Location = new Point(957, 228);
            }

            if (radioButton2.Checked == true)
            {
                textBox_h5.Location = new Point(421, 246);
            }
            else
            {
                textBox_h5.Location = new Point(955, 246);
            }

            if (radioButton2.Checked == true)
            {
                label37.Location = new Point(358, 113);
            }
            else
            {
                label37.Location = new Point(1049, 113);
            }

            if (radioButton2.Checked == true)
            {
                textBox_L5.Location = new Point(333, 130);
            }
            else
            {
                textBox_L5.Location = new Point(1024, 130);
            }

            if (radioButton2.Checked == true)
            {
                panel11.Location = new Point(186, 225);
            }
            else
            {
                panel11.Location = new Point(1108, 225);
            }

            if (radioButton2.Checked == true)
            {
                panel10.Location = new Point(159, 292);
            }
            else
            {
                panel10.Location = new Point(1066, 292);
            }

            if (radioButton2.Checked == true)
            {
                panel12.Location = new Point(14, 99);
            }
            else
            {
                panel12.Location = new Point(1216, 99);
            }
        }
    }
}
