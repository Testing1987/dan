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
    public partial class AGEN_crossing_list : Form
    {
        bool clickdragdown;
        Point lastLocation;

        static bool Freeze_operations = false;


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
            maximize_agen();

            this.Close();
        }
        private void button_minimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }


        public AGEN_crossing_list()
        {
            InitializeComponent();

            label_layer.Text = _AGEN_mainform.layer_crossing;

            if (_AGEN_mainform.Data_Table_crossings != null)
            {
                if (_AGEN_mainform.Data_Table_crossings.Rows.Count > 0)
                {
                    if (_AGEN_mainform.Data_Table_crossings.Columns.Contains("Layer") == true &&
                        _AGEN_mainform.Data_Table_crossings.Columns.Contains("2DSta") == true &&
                        _AGEN_mainform.Data_Table_crossings.Columns.Contains("3DSta") == true &&
                        _AGEN_mainform.Data_Table_crossings.Columns.Contains("EqSta") == true &&
                        _AGEN_mainform.Data_Table_crossings.Columns.Contains("Desc") == true &&
                        _AGEN_mainform.Data_Table_crossings.Columns.Contains("Prof Block Name") == true &&
                        _AGEN_mainform.Data_Table_crossings.Columns.Contains("Attrib Sta Prof") == true &&
                        _AGEN_mainform.Data_Table_crossings.Columns.Contains("Attrib Desc Prof") == true)
                    {
                        int rowno = 0;


                        for (int i = 0; i < _AGEN_mainform.Data_Table_crossings.Rows.Count; ++i)
                        {
                            if (_AGEN_mainform.Data_Table_crossings.Rows[i]["Layer"] != DBNull.Value)
                            {
                                string ln = Convert.ToString(_AGEN_mainform.Data_Table_crossings.Rows[i]["Layer"]);
                                if (ln == _AGEN_mainform.layer_crossing)
                                {
                                    string block = "";
                                    string sta2d = "";
                                    string eqsta = "";
                                    string desc = "";
                                    string atr_sta = "";
                                    string atr_desc = "";

                                    if (_AGEN_mainform.Data_Table_crossings.Rows[i]["Prof Block Name"] != DBNull.Value)
                                    {
                                        block = Convert.ToString(_AGEN_mainform.Data_Table_crossings.Rows[i]["Prof Block Name"]);
                                        if (rowno == 0)
                                        {
                                            if (comboBox_prof_block1.Items.Contains(block) == false) comboBox_prof_block1.Items.Add(block);
                                            comboBox_prof_block1.SelectedIndex = comboBox_prof_block1.Items.IndexOf(block);
                                        }
                                    }
                                    if (_AGEN_mainform.Data_Table_crossings.Rows[i]["Attrib Sta Prof"] != DBNull.Value)
                                    {
                                        atr_sta = Convert.ToString(_AGEN_mainform.Data_Table_crossings.Rows[i]["Attrib Sta Prof"]);
                                        if (rowno == 0)
                                        {
                                            if (comboBox_prof_atr_sta1.Items.Contains(atr_sta) == false) comboBox_prof_atr_sta1.Items.Add(atr_sta);
                                            comboBox_prof_atr_sta1.SelectedIndex = comboBox_prof_atr_sta1.Items.IndexOf(atr_sta);
                                        }
                                    }
                                    if (_AGEN_mainform.Data_Table_crossings.Rows[i]["Attrib Desc Prof"] != DBNull.Value)
                                    {
                                        atr_desc = Convert.ToString(_AGEN_mainform.Data_Table_crossings.Rows[i]["Attrib Desc Prof"]);
                                        if (rowno == 0)
                                        {
                                            if (comboBox_prof_atr_desc1.Items.Contains(atr_desc) == false) comboBox_prof_atr_desc1.Items.Add(atr_desc);
                                            comboBox_prof_atr_desc1.SelectedIndex = comboBox_prof_atr_desc1.Items.IndexOf(atr_desc);
                                        }

                                    }
                                    if (_AGEN_mainform.Data_Table_crossings.Rows[i]["2DSta"] != DBNull.Value)
                                    {
                                        sta2d = Convert.ToString(_AGEN_mainform.Data_Table_crossings.Rows[i]["2DSta"]);
                                        if (rowno == 0)
                                        {
                                            textBox_s1.Text = sta2d;
                                        }
                                    }
                                    if (_AGEN_mainform.Data_Table_crossings.Rows[i]["EqSta"] != DBNull.Value)
                                    {
                                        eqsta = Convert.ToString(_AGEN_mainform.Data_Table_crossings.Rows[i]["EqSta"]);
                                        if (rowno == 0)
                                        {
                                            textBox_eq1.Text = eqsta;
                                        }
                                    }
                                    if (_AGEN_mainform.Data_Table_crossings.Rows[i]["Desc"] != DBNull.Value)
                                    {
                                        desc = Convert.ToString(_AGEN_mainform.Data_Table_crossings.Rows[i]["Desc"]);
                                        if (rowno == 0)
                                        {
                                            textBox_d1.Text = desc;
                                        }
                                    }

                                    if (rowno > 0)
                                    {
                                        ComboBox combo1 = new ComboBox();
                                        combo1.Location = new Point(comboBox_prof_block1.Location.X, comboBox_prof_block1.Location.Y + rowno * (comboBox_prof_block1.Height + 8));
                                        combo1.BackColor = comboBox_prof_block1.BackColor;
                                        combo1.ForeColor = comboBox_prof_block1.ForeColor;
                                        combo1.Font = comboBox_prof_block1.Font;
                                        combo1.Size = comboBox_prof_block1.Size;
                                        combo1.FlatStyle = comboBox_prof_block1.FlatStyle;
                                        if (block != "")
                                        {
                                            if (combo1.Items.Contains(block) == false) combo1.Items.Add(block);
                                            combo1.SelectedIndex = combo1.Items.IndexOf(block);
                                        }
                                        panel_crossing_list.Controls.Add(combo1);

                                        combo1.SelectedIndexChanged += delegate (object s, EventArgs e1)
                                        {
                                            combo1_SelectedIndexChanged(combo1, e1);
                                        };

                                        ComboBox combo2 = new ComboBox();
                                        combo2.Location = new Point(comboBox_prof_atr_desc1.Location.X, comboBox_prof_atr_desc1.Location.Y + rowno * (comboBox_prof_atr_desc1.Height + 8));
                                        combo2.BackColor = comboBox_prof_atr_desc1.BackColor;
                                        combo2.ForeColor = comboBox_prof_atr_desc1.ForeColor;
                                        combo2.Font = comboBox_prof_atr_desc1.Font;
                                        combo2.Size = comboBox_prof_atr_desc1.Size;
                                        combo2.FlatStyle = comboBox_prof_atr_desc1.FlatStyle;
                                        if (atr_desc != "")
                                        {
                                            if (combo2.Items.Contains(atr_desc) == false) combo2.Items.Add(atr_desc);
                                            combo2.SelectedIndex = combo2.Items.IndexOf(atr_desc);
                                        }
                                        panel_crossing_list.Controls.Add(combo2);

                                        ComboBox combo3 = new ComboBox();
                                        combo3.Location = new Point(comboBox_prof_atr_sta1.Location.X, comboBox_prof_atr_sta1.Location.Y + rowno * (comboBox_prof_atr_sta1.Height + 8));
                                        combo3.BackColor = comboBox_prof_atr_sta1.BackColor;
                                        combo3.ForeColor = comboBox_prof_atr_sta1.ForeColor;
                                        combo3.Font = comboBox_prof_atr_sta1.Font;
                                        combo3.Size = comboBox_prof_atr_sta1.Size;
                                        combo3.FlatStyle = comboBox_prof_atr_sta1.FlatStyle;
                                        if (atr_sta != "")
                                        {
                                            if (combo3.Items.Contains(atr_sta) == false) combo3.Items.Add(atr_sta);
                                            combo3.SelectedIndex = combo3.Items.IndexOf(atr_sta);
                                        }
                                        panel_crossing_list.Controls.Add(combo3);

                                        Button bt1 = new Button();
                                        bt1.Location = new Point(button_refresh_blocks.Location.X, button_refresh_blocks.Location.Y + rowno * (button_refresh_blocks.Height + 8));
                                        bt1.BackColor = button_refresh_blocks.BackColor;
                                        bt1.ForeColor = button_refresh_blocks.ForeColor;
                                        bt1.Font = button_refresh_blocks.Font;
                                        bt1.Size = button_refresh_blocks.Size;
                                        bt1.FlatStyle = button_refresh_blocks.FlatStyle;
                                        bt1.FlatAppearance.BorderColor = button_refresh_blocks.FlatAppearance.BorderColor;
                                        bt1.FlatAppearance.BorderSize = button_refresh_blocks.FlatAppearance.BorderSize;
                                        bt1.FlatAppearance.MouseDownBackColor = button_refresh_blocks.FlatAppearance.MouseDownBackColor;
                                        bt1.FlatAppearance.MouseOverBackColor = button_refresh_blocks.FlatAppearance.MouseOverBackColor;
                                        bt1.BackgroundImage = button_refresh_blocks.BackgroundImage;
                                        bt1.BackgroundImageLayout = button_refresh_blocks.BackgroundImageLayout;
                                        panel_crossing_list.Controls.Add(bt1);

                                        bt1.Click += delegate (object s, EventArgs e1)
                                        {
                                            button_refresh_line_Click(bt1, e1);
                                        };

                                        TextBox tb1 = new TextBox();
                                        tb1.Location = new Point(textBox_s1.Location.X, textBox_s1.Location.Y + rowno * (textBox_s1.Height + 8));
                                        tb1.BackColor = textBox_s1.BackColor;
                                        tb1.ForeColor = textBox_s1.ForeColor;
                                        tb1.Font = textBox_s1.Font;
                                        tb1.Size = textBox_s1.Size;
                                        tb1.ReadOnly = textBox_s1.ReadOnly;

                                        if (sta2d != "")
                                        {
                                            tb1.Text = sta2d;
                                        }
                                        panel_crossing_list.Controls.Add(tb1);

                                        TextBox tb2 = new TextBox();
                                        tb2.Location = new Point(textBox_eq1.Location.X, textBox_eq1.Location.Y + rowno * (textBox_eq1.Height + 8));
                                        tb2.BackColor = textBox_eq1.BackColor;
                                        tb2.ForeColor = textBox_eq1.ForeColor;
                                        tb2.Font = textBox_eq1.Font;
                                        tb2.Size = textBox_eq1.Size;
                                        tb2.ReadOnly = textBox_eq1.ReadOnly;

                                        if (eqsta != "")
                                        {
                                            tb2.Text = eqsta;
                                        }
                                        panel_crossing_list.Controls.Add(tb2);

                                        TextBox tb3 = new TextBox();
                                        tb3.Location = new Point(textBox_d1.Location.X, textBox_d1.Location.Y + rowno * (textBox_d1.Height + 8));
                                        tb3.BackColor = textBox_d1.BackColor;
                                        tb3.ForeColor = textBox_d1.ForeColor;
                                        tb3.Font = textBox_d1.Font;
                                        tb3.Size = textBox_d1.Size;
                                        tb3.ReadOnly = textBox_d1.ReadOnly;

                                        if (desc != "")
                                        {
                                            tb3.Text = desc;
                                        }
                                        panel_crossing_list.Controls.Add(tb3);

                                    }


                                    rowno = rowno + 1;
                                }
                            }
                        }
                    }
                }
            }




        }

        private void button_refresh_line_Click(object sender, EventArgs e)
        {
            Button c1 = sender as Button;
            int Y = c1.Location.Y;
            foreach (Control ctrl in panel_crossing_list.Controls)
            {
                ComboBox combo_b = ctrl as ComboBox;
                if (combo_b != null)
                {
                    if (combo_b.Location.Y == Y && combo_b.Location.X == comboBox_prof_block1.Location.X)
                    {
                        Functions.Incarca_existing_Blocks_with_attributes_to_combobox(combo_b);
                    }
                    if (combo_b.Location.Y == Y && (combo_b.Location.X == comboBox_prof_atr_sta1.Location.X || combo_b.Location.X == comboBox_prof_atr_desc1.Location.X))
                    {
                        combo_b.Text = "";
                    }
                }
            }

        }

        private void combo1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox c1 = sender as ComboBox;
            int Y = c1.Location.Y;
            foreach (Control ctrl in panel_crossing_list.Controls)
            {
                ComboBox combo_atr = ctrl as ComboBox;
                if (combo_atr != null)
                {
                    if (combo_atr.Location.Y == Y && (combo_atr.Location.X == comboBox_prof_atr_sta1.Location.X || combo_atr.Location.X == comboBox_prof_atr_desc1.Location.X))
                    {
                        Functions.Incarca_existing_Atributes_to_combobox(c1.Text, combo_atr);
                    }
                }
            }
        }


        private void display_layer_data_from_dt_crossing(string station)
        {
            if (_AGEN_mainform.Data_Table_crossings != null)
            {
                if (_AGEN_mainform.Data_Table_crossings.Rows.Count > 0)
                {
                    if (_AGEN_mainform.Data_Table_crossings.Columns.Contains("Layer") == true &&
                        _AGEN_mainform.Data_Table_crossings.Columns.Contains("Prof Block Name") == true &&
                        _AGEN_mainform.Data_Table_crossings.Columns.Contains("Attrib Sta Prof") == true &&
                        _AGEN_mainform.Data_Table_crossings.Columns.Contains("Attrib Desc Prof") == true)
                    {
                        for (int i = 0; i < _AGEN_mainform.Data_Table_crossings.Rows.Count; ++i)
                        {
                            if (_AGEN_mainform.Data_Table_crossings.Rows[i]["Layer"] != DBNull.Value)
                            {
                                string ln = Convert.ToString(_AGEN_mainform.Data_Table_crossings.Rows[i]["Layer"]);
                                if (ln == station)
                                {
                                    if (_AGEN_mainform.Data_Table_crossings.Rows[i]["Prof Block Name"] != DBNull.Value)
                                    {

                                        string bn = Convert.ToString(_AGEN_mainform.Data_Table_crossings.Rows[i]["Prof Block Name"]);

                                        if (_AGEN_mainform.Data_Table_crossings.Rows[i]["Attrib Sta Prof"] != DBNull.Value)
                                        {
                                            string a1 = Convert.ToString(_AGEN_mainform.Data_Table_crossings.Rows[i]["Attrib Sta Prof"]);

                                        }

                                        if (_AGEN_mainform.Data_Table_crossings.Rows[i]["Attrib Desc Prof"] != DBNull.Value)
                                        {
                                            string a2 = Convert.ToString(_AGEN_mainform.Data_Table_crossings.Rows[i]["Attrib Desc Prof"]);

                                        }

                                        i = _AGEN_mainform.Data_Table_crossings.Rows.Count;
                                    }
                                    else
                                    {

                                    }

                                }
                            }

                        }
                    }
                }
            }
        }



        private void button_update_crossing_list_excel_Click(object sender, EventArgs e)
        {
            if (Freeze_operations == false)
            {
                Freeze_operations = true;
                try
                {
                    string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();

                    if (System.IO.Directory.Exists(ProjFolder) == true)
                    {
                        if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                        {
                            ProjFolder = ProjFolder + "\\";
                        }
                    }
                    else
                    {
                        Freeze_operations = false;
                        MessageBox.Show("the project folder does not exist");
                        return;
                    }

                    string fisier_cs = ProjFolder + _AGEN_mainform.crossing_excel_name;

                    if (System.IO.File.Exists(fisier_cs) == false)
                    {
                        Freeze_operations = false;
                        MessageBox.Show("the centerline data file does not exist");
                        return;
                    }

                    update_data_in_dtcrossing(label_layer.Text);

                    Functions.create_backup(fisier_cs);
                    _AGEN_mainform.tpage_crossing_scan.Populate_crossing_file(fisier_cs);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                Freeze_operations = false;
            }

            maximize_agen();

            // Functions.Transfer_datatable_to_new_excel_spreadsheet(AGEN_mainform.Data_Table_custom_bands);
            button_Exit_Click(sender, e);
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

        private void update_data_in_dtcrossing(string lname)
        {
            if (_AGEN_mainform.Data_Table_crossings != null)
            {
                if (_AGEN_mainform.Data_Table_crossings.Rows.Count > 0)
                {
                    if (_AGEN_mainform.Data_Table_crossings.Columns.Contains("Layer") == true &&
                        _AGEN_mainform.Data_Table_crossings.Columns.Contains("2DSta") == true &&
                        _AGEN_mainform.Data_Table_crossings.Columns.Contains("3DSta") == true &&
                        _AGEN_mainform.Data_Table_crossings.Columns.Contains("Desc") == true &&
                        _AGEN_mainform.Data_Table_crossings.Columns.Contains("DispProf") == true &&
                        _AGEN_mainform.Data_Table_crossings.Columns.Contains("Prof Block Name") == true &&
                        _AGEN_mainform.Data_Table_crossings.Columns.Contains("Attrib Sta Prof") == true &&
                        _AGEN_mainform.Data_Table_crossings.Columns.Contains("Attrib Desc Prof") == true)
                    {
                        for (int i = 0; i < _AGEN_mainform.Data_Table_crossings.Rows.Count; ++i)
                        {
                            if (_AGEN_mainform.Data_Table_crossings.Rows[i]["Layer"] != DBNull.Value)
                            {
                                if (_AGEN_mainform.Data_Table_crossings.Rows[i]["2DSta"] != DBNull.Value || _AGEN_mainform.Data_Table_crossings.Rows[i]["3DSta"] != DBNull.Value)
                                {
                                    if (_AGEN_mainform.Data_Table_crossings.Rows[i]["Desc"] != DBNull.Value)
                                    {
                                        string ln = Convert.ToString(_AGEN_mainform.Data_Table_crossings.Rows[i]["Layer"]);

                                        if (ln == lname)
                                        {
                                            string sta = "";
                                            if (_AGEN_mainform.Data_Table_crossings.Rows[i]["2DSta"] != DBNull.Value)
                                            {
                                                sta = Convert.ToString(_AGEN_mainform.Data_Table_crossings.Rows[i]["2DSta"]);
                                            }

                                            if (_AGEN_mainform.Data_Table_crossings.Rows[i]["3DSta"] != DBNull.Value)
                                            {
                                                sta = Convert.ToString(_AGEN_mainform.Data_Table_crossings.Rows[i]["3DSta"]);
                                            }

                                            string desc = Convert.ToString(_AGEN_mainform.Data_Table_crossings.Rows[i]["Desc"]);
                                            if (sta != "")
                                            {
                                                int rowno = 0;
                                                foreach (Control ctrl1 in panel_crossing_list.Controls)
                                                {
                                                    TextBox t1 = ctrl1 as TextBox;
                                                    if (t1 != null)
                                                    {
                                                        if (t1.Location.X == textBox_s1.Location.X && t1.Location.Y == textBox_s1.Location.Y + rowno * (textBox_s1.Height + 8))
                                                        {
                                                            string sta1 = t1.Text;
                                                            if (sta1 == sta)
                                                            {
                                                                foreach (Control ctrl2 in panel_crossing_list.Controls)
                                                                {
                                                                    TextBox t2 = ctrl2 as TextBox;
                                                                    if (t2 != null)
                                                                    {
                                                                        if (t2.Location.X == textBox_d1.Location.X && t2.Location.Y == textBox_d1.Location.Y + rowno * (textBox_d1.Height + 8))
                                                                        {
                                                                            string desc1 = t2.Text;
                                                                            if (sta == sta1 && desc == desc1)
                                                                            {
                                                                                string block1 = "";
                                                                                string atr_desc1 = "";
                                                                                string atr_sta1 = "";
                                                                                foreach (Control ctrl3 in panel_crossing_list.Controls)
                                                                                {
                                                                                    ComboBox c3 = ctrl3 as ComboBox;
                                                                                    if (c3 != null)
                                                                                    {
                                                                                        if (c3.Location.X == comboBox_prof_block1.Location.X && c3.Location.Y == comboBox_prof_block1.Location.Y + rowno * (comboBox_prof_block1.Height + 8))
                                                                                        {
                                                                                            block1 = c3.Text;
                                                                                        }
                                                                                        if (c3.Location.X == comboBox_prof_atr_desc1.Location.X && c3.Location.Y == comboBox_prof_atr_desc1.Location.Y + rowno * (comboBox_prof_atr_desc1.Height + 8))
                                                                                        {
                                                                                            atr_desc1 = c3.Text;
                                                                                        }
                                                                                        if (c3.Location.X == comboBox_prof_atr_sta1.Location.X && c3.Location.Y == comboBox_prof_atr_sta1.Location.Y + rowno * (comboBox_prof_atr_sta1.Height + 8))
                                                                                        {
                                                                                            atr_sta1 = c3.Text;
                                                                                        }
                                                                                    }
                                                                                }

                                                                                if (block1 != "")
                                                                                {
                                                                                    _AGEN_mainform.Data_Table_crossings.Rows[i]["Prof Block Name"] = block1;
                                                                                    if (atr_sta1 != "")
                                                                                    {
                                                                                        _AGEN_mainform.Data_Table_crossings.Rows[i]["Attrib Sta Prof"] = atr_sta1;

                                                                                    }
                                                                                    else
                                                                                    {
                                                                                        _AGEN_mainform.Data_Table_crossings.Rows[i]["Attrib Sta Prof"] = DBNull.Value;

                                                                                    }
                                                                                    if (atr_desc1 != "")
                                                                                    {
                                                                                        _AGEN_mainform.Data_Table_crossings.Rows[i]["Attrib Desc Prof"] = atr_desc1;

                                                                                    }
                                                                                    else
                                                                                    {
                                                                                        _AGEN_mainform.Data_Table_crossings.Rows[i]["Attrib Desc Prof"] = DBNull.Value;

                                                                                    }
                                                                                    _AGEN_mainform.Data_Table_crossings.Rows[i]["DispProf"] = "YES";
                                                                                }
                                                                                else
                                                                                {
                                                                                    _AGEN_mainform.Data_Table_crossings.Rows[i]["Prof Block Name"] = DBNull.Value;
                                                                                    _AGEN_mainform.Data_Table_crossings.Rows[i]["Attrib Sta Prof"] = DBNull.Value;
                                                                                    _AGEN_mainform.Data_Table_crossings.Rows[i]["Attrib Desc Prof"] = DBNull.Value;
                                                                                    _AGEN_mainform.Data_Table_crossings.Rows[i]["DispProf"] = "NO";
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            else
                                                            {
                                                                rowno = rowno + 1;
                                                            }
                                                        }
                                                    }
                                                }
                                            }


                                        }
                                    }
                                }

                            }
                        }

                        for (int i = 0; i < _AGEN_mainform.Data_Table_crossings.Rows.Count; ++i)
                        {
                            if (_AGEN_mainform.Data_Table_crossings.Rows[i]["DispProf"] == DBNull.Value)
                            {
                                _AGEN_mainform.Data_Table_crossings.Rows[i]["DispProf"] = "NO";
                            }
                        }
                    }
                }
            }
        }


     

        private void button_refresh_blocks_Click(object sender, EventArgs e)
        {
            Functions.Incarca_existing_Blocks_with_attributes_to_combobox(comboBox_prof_block1);
            comboBox_prof_atr_sta1.Text = "";
            comboBox_prof_atr_desc1.Text = "";
        }

        private void comboBox_prof_block1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Functions.Incarca_existing_Atributes_to_combobox(comboBox_prof_block1.Text, comboBox_prof_atr_sta1);
            Functions.Incarca_existing_Atributes_to_combobox(comboBox_prof_block1.Text, comboBox_prof_atr_desc1);
        }
    }
}
