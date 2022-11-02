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
    public partial class AGEN_dwg_selection : Form
    {
        bool clickdragdown;
        Point lastLocation;

        int nr_bands = 0;
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

        public AGEN_dwg_selection()
        {
            InitializeComponent();
            add_controls_to_custom_form();
            comboBox1.DropDownWidth = 250;
            comboBox2.DropDownWidth = 250;
        }

        private void add_controls_to_custom_form()
        {
            if (_AGEN_mainform.Data_Table_profile_band != null && _AGEN_mainform.Data_Table_profile_band.Rows.Count > 0)
            {
                if (_AGEN_mainform.Data_Table_profile_band.Rows[0]["DwgNo"] != DBNull.Value)
                {
                    string nume_dwg = Convert.ToString(_AGEN_mainform.Data_Table_profile_band.Rows[0]["DwgNo"]);
                    label1.Text = nume_dwg;
                    if (_AGEN_mainform.lista_gen_prof_band != null && _AGEN_mainform.lista_gen_prof_band.Count > 0 && _AGEN_mainform.lista_gen_prof_band.Contains(0) == true)
                    {
                        checkBox1.Checked = true;
                    }
                    else
                    {
                        checkBox1.Checked = false;
                    }
                    comboBox1.Items.Add(nume_dwg);
                    comboBox2.Items.Add(nume_dwg);
                }

                if (_AGEN_mainform.Data_Table_profile_band.Rows.Count > 1)
                {
                    for (int i = 1; i < _AGEN_mainform.Data_Table_profile_band.Rows.Count; ++i)
                    {
                        if (_AGEN_mainform.Data_Table_profile_band.Rows[i]["DwgNo"] != DBNull.Value)
                        {
                            string nume_dwg = Convert.ToString(_AGEN_mainform.Data_Table_profile_band.Rows[i]["DwgNo"]);

                            bool checked1 = false;
                            if (_AGEN_mainform.lista_gen_prof_band != null && _AGEN_mainform.lista_gen_prof_band.Count > 0 && _AGEN_mainform.lista_gen_prof_band.Contains(i) == true)
                            {
                                checked1 = true;
                            }
                            add_new_control_row(new object(), new EventArgs(), nume_dwg, checked1);
                            comboBox1.Items.Add(nume_dwg);
                            comboBox2.Items.Add(nume_dwg);

                        }
                    }
                }
            }

        }



        private void button_Exit_Click(object sender, EventArgs e)
        {
            maximize_agen();

            this.Close();
        }




        private void add_new_control_row(object sender, EventArgs e, string nume_band, bool checked1)
        {

            Label label2 = new Label();
            label2.Location = new Point(label1.Location.X, label1.Location.Y + (nr_bands + 1) * (checkBox1.Height + 4));
            label2.BackColor = label1.BackColor;
            label2.ForeColor = label1.ForeColor;
            label2.Font = label1.Font;
            label2.Text = nume_band;
            label2.Size = new Size(label1.Size.Width + 20, label1.Size.Height);
            panel_dwg.Controls.Add(label2);

            CheckBox checkbox2 = new CheckBox();
            checkbox2.Location = new Point(checkBox1.Location.X, checkBox1.Location.Y + (nr_bands + 1) * (checkBox1.Height + 4));
            checkbox2.BackColor = checkBox1.BackColor;
            checkbox2.ForeColor = checkBox1.ForeColor;
            checkbox2.Font = checkBox1.Font;
            checkbox2.Size = checkBox1.Size;
            checkbox2.Text = "";
            checkbox2.Checked = checked1;
            panel_dwg.Controls.Add(checkbox2);

            checkbox2.CheckedChanged += delegate (object s, EventArgs e1)
            {
                checkBox_CheckedChanged(checkbox2, e1);
            };

            nr_bands = nr_bands + 1;
        }

        private void maximize_agen()
        {
            foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
            {
                if (Forma1 is Alignment_mdi._AGEN_mainform)
                {
                    Forma1.Focus();
                    Forma1.WindowState = System.Windows.Forms.FormWindowState.Normal;
                }
            }
        }

        private void checkBox_CheckedChanged(object sender, EventArgs e)
        {
            if (_AGEN_mainform.lista_gen_prof_band == null) _AGEN_mainform.lista_gen_prof_band = new List<int>();
            CheckBox chk1 = sender as CheckBox;
            if (chk1 != null)
            {
                int Y = chk1.Location.Y;
                int index1 = (Y - checkBox1.Location.Y) / (checkBox1.Height + 4);

                if (chk1.Checked == true)
                {

                    if (_AGEN_mainform.lista_gen_prof_band.Contains(index1) == false) _AGEN_mainform.lista_gen_prof_band.Add(index1);

                }
                else
                {

                    if (_AGEN_mainform.lista_gen_prof_band.Contains(index1) == true) _AGEN_mainform.lista_gen_prof_band.Remove(index1);

                }
            }

        }

        private void button_cancel_Click(object sender, EventArgs e)
        {
            _AGEN_mainform.lista_gen_prof_band = null;
            maximize_agen();
            this.Close();
        }

        private void button_apply_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text != "" && comboBox2.Text != "")
            {
                string nume1 = comboBox1.Text;
                string nume2 = comboBox2.Text;
                bool colect1 = false;
                for (int i = 0; i < _AGEN_mainform.Data_Table_profile_band.Rows.Count; ++i)
                {
                    if(_AGEN_mainform.Data_Table_profile_band.Rows[i]["DwgNo"]!=DBNull.Value)
                    {
                        string nume = Convert.ToString(_AGEN_mainform.Data_Table_profile_band.Rows[i]["DwgNo"]);
                        if(nume.ToLower()==nume1.ToLower())
                        {
                            colect1 = true;
                        }

                        if (nume.ToLower() == nume2.ToLower())
                        {
                            if (_AGEN_mainform.lista_gen_prof_band == null) _AGEN_mainform.lista_gen_prof_band = new List<int>();
                            if (_AGEN_mainform.lista_gen_prof_band.Contains(i) == false) _AGEN_mainform.lista_gen_prof_band.Add(i);
                            colect1 = false;
                        }

                        if (colect1==true)
                        {
                            if (_AGEN_mainform.lista_gen_prof_band == null) _AGEN_mainform.lista_gen_prof_band = new List<int>();
                            if (_AGEN_mainform.lista_gen_prof_band.Contains(i) == false) _AGEN_mainform.lista_gen_prof_band.Add(i);
                        }

                    }
                }

                foreach(Control ctrl1 in panel_dwg.Controls)
                {
                    CheckBox chk1 = ctrl1 as CheckBox;
                    if (chk1!=null)
                    {
                        int Y = chk1.Location.Y;
                        int index1 = (Y - checkBox1.Location.Y) / (checkBox1.Height + 4);
                        if (_AGEN_mainform.lista_gen_prof_band == null) _AGEN_mainform.lista_gen_prof_band = new List<int>();
                        if (_AGEN_mainform.lista_gen_prof_band.Contains(index1) == true) chk1.Checked = true;

                    }
                }
            }
        }
    }
}
