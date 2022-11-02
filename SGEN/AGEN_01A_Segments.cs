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
    public partial class AGEN_segments_form : Form
    {
        bool clickdragdown;
        Point lastLocation;

        int idx_segm = 0;
        double one_row_height = 0;
        double last_Y1 = -1;
        double last_Y2 = -1;

        public AGEN_segments_form()
        {
            InitializeComponent();
            last_Y1 = label_name1.Location.Y;
            last_Y2 = textBox_name1.Location.Y;
            one_row_height = panel_segments.Height;
            add_controls_to_segment_form();
        }

        private void add_controls_to_segment_form()
        {
            if (_SGEN_mainform.dt_segments != null && _SGEN_mainform.dt_segments.Rows.Count > 0)
            {

                textBox_name1.Text = Convert.ToString(_SGEN_mainform.dt_segments.Rows[0]["Segment Name"]);

                if (_SGEN_mainform.dt_segments.Rows.Count > 1)
                {
                    for (int i = 1; i < _SGEN_mainform.dt_segments.Rows.Count; ++i)
                    {
                        if (_SGEN_mainform.dt_segments.Rows[i]["Segment Name"] != DBNull.Value)
                        {
                            idx_segm = i;
                            string nume_segm = Convert.ToString(_SGEN_mainform.dt_segments.Rows[i]["Segment Name"]);
                            add_new_control_row(new object(), new EventArgs(), nume_segm, Alignment_mdi.Properties.Resources.check);
                        }

                    }
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

        private void button_minimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void button_add_custom_control_Click(object sender, EventArgs e)
        {
            if (idx_segm > 54)
            {
                MessageBox.Show("you can't have more than 54 segments");
                return;
            }
            ++idx_segm;
            add_new_control_row(sender, e, "", Alignment_mdi.Properties.Resources.selectbluexs);
        }

        private void add_new_control_row(object sender, EventArgs e, string nume_band, System.Drawing.Bitmap bitmap1)
        {

            Label label1 = new Label();
            label1.Location = new Point(label_name1.Location.X, label_name1.Location.Y + idx_segm * (textBox_name1.Height + 4));
            label1.BackColor = label_name1.BackColor;
            label1.ForeColor = label_name1.ForeColor;
            label1.Font = label_name1.Font;
            label1.Text = "Segment " + (idx_segm + 1).ToString() + " Name";
            label1.Size = new Size(label_name1.Size.Width + 20, label_name1.Size.Height);
            panel_segments.Controls.Add(label1);
            last_Y1 = label1.Location.Y;

            TextBox textbox1 = new TextBox();
            textbox1.Location = new Point(textBox_name1.Location.X, textBox_name1.Location.Y + idx_segm * (textBox_name1.Height + 4));
            textbox1.BackColor = textBox_name1.BackColor;
            textbox1.ForeColor = textBox_name1.ForeColor;
            textbox1.Font = textBox_name1.Font;
            textbox1.Size = textBox_name1.Size;
            textbox1.Text = nume_band;
            panel_segments.Controls.Add(textbox1);
            last_Y2 = textbox1.Location.Y;


            button_create.Location = new Point(button_create.Location.X, button_create.Location.Y + textBox_name1.Height + 4);

            panel_segments.Size = new Size(panel_segments.Width, panel_segments.Height + textBox_name1.Height + 4);
            this.Size = new Size(this.Width, this.Height + textBox_name1.Height + 4);
            textbox1.Select();
        }

        private void button_remove_Click(object sender, EventArgs e)
        {
            if (panel_segments.Height > one_row_height)
            {
                for (int i = panel_segments.Controls.Count - 1; i >= 0; --i)
                {
                    Control CTRL1 = panel_segments.Controls[i];
                    if (CTRL1.Location.Y == last_Y1 || CTRL1.Location.Y == last_Y2)
                    {
                        panel_segments.Controls.Remove(CTRL1);

                    }
                }

                idx_segm = idx_segm - 1;
                last_Y1 = last_Y1 - textBox_name1.Height - 4;
                last_Y2 = last_Y2 - textBox_name1.Height - 4;

                button_create.Location = new Point(button_create.Location.X, button_create.Location.Y - textBox_name1.Height - 4);

                panel_segments.Size = new Size(panel_segments.Width, panel_segments.Height - textBox_name1.Height - 4);
                this.Size = new Size(this.Width, this.Height - textBox_name1.Height - 4);
            }
        }

        private void button_create_segments_Click(object sender, EventArgs e)
        {

            System.Data.DataTable dt1 = _SGEN_mainform.dt_segments.Copy();

            _SGEN_mainform.dt_segments = new System.Data.DataTable();
            _SGEN_mainform.dt_segments.Columns.Add("Template", typeof(string));
            _SGEN_mainform.dt_segments.Columns.Add("Output folder", typeof(string));
            _SGEN_mainform.dt_segments.Columns.Add("Prefix File Name", typeof(string));
            _SGEN_mainform.dt_segments.Columns.Add("Suffix File Name", typeof(string));
            _SGEN_mainform.dt_segments.Columns.Add("Start numbering", typeof(string));
            _SGEN_mainform.dt_segments.Columns.Add("Increment", typeof(string));
            _SGEN_mainform.dt_segments.Columns.Add("Segment Name", typeof(string));





            for (int i = 0; i < panel_segments.Controls.Count; ++i)
            {
                TextBox tb1 = panel_segments.Controls[i] as TextBox;
                if (tb1 != null)
                {
                    if (tb1.Text != "" &&
                        tb1.Text.Contains("\\") == false &&
                        tb1.Text.Contains("/") == false &&
                        tb1.Text.Contains(":") == false &&
                        tb1.Text.Contains("*") == false &&
                        tb1.Text.Contains("?") == false &&
                        tb1.Text.Contains("<") == false &&
                        tb1.Text.Contains(">") == false &&
                        tb1.Text.Contains("|") == false &&
                        tb1.Text.Contains("\"") == false)
                    {
                        bool exista = false;
                        if (_SGEN_mainform.dt_segments.Rows.Count > 0)
                        {
                            for (int j = 0; j < _SGEN_mainform.dt_segments.Rows.Count; ++j)
                            {
                                string nume_existent = Convert.ToString(_SGEN_mainform.dt_segments.Rows[j]["Segment Name"]);
                                if (nume_existent.ToUpper() == tb1.Text.ToUpper())
                                {
                                    exista = true;
                                }
                            }
                        }



                        if (tb1.Text.Contains("\\") == true &&
                            tb1.Text.Contains("/") == true &&
                            tb1.Text.Contains(":") == true &&
                            tb1.Text.Contains("*") == true &&
                            tb1.Text.Contains("?") == true &&
                            tb1.Text.Contains("<") == true &&
                            tb1.Text.Contains(">") == true &&
                            tb1.Text.Contains("|") == true &&
                            tb1.Text.Contains("\"") == true)
                        {
                            MessageBox.Show("you can't have the following characters into the segment name:\r\n\\/:*?\"<>|");
                            return;

                        }

                        if (exista == false)
                        {
                            if (tb1.Text != null)
                            {
                                _SGEN_mainform.dt_segments.Rows.Add();
                                _SGEN_mainform.dt_segments.Rows[_SGEN_mainform.dt_segments.Rows.Count - 1]["Segment Name"] = tb1.Text;
                                _SGEN_mainform.no_of_segments = _SGEN_mainform.dt_segments.Rows.Count;
                            }
                        }

                    }

                }
            }

            for (int i = 0; i < _SGEN_mainform.dt_segments.Rows.Count; ++i)
            {
                if (_SGEN_mainform.dt_segments.Rows[i]["Segment Name"] != DBNull.Value)
                {
                    string segment1 = Convert.ToString(_SGEN_mainform.dt_segments.Rows[i]["Segment Name"]);

                    for (int j = 0; j < dt1.Rows.Count; ++j)
                    {
                        if (dt1.Rows[j]["Segment Name"] != DBNull.Value)
                        {
                            string segment2 = Convert.ToString(dt1.Rows[j]["Segment Name"]);

                            if(segment1.ToUpper()==segment2.ToUpper())
                            {
                                for(int k=0; k<6;++k)
                                {
                                    _SGEN_mainform.dt_segments.Rows[i][k] = dt1.Rows[j][k];
                                }
                            }

                        }
                    }
                }

            }



            _SGEN_mainform.tpage_settings.Fill_combobox_segments();
          


            for (int i = 0; i < _SGEN_mainform.dt_segments.Rows.Count; ++i)
            {
                if (_SGEN_mainform.dt_segments.Rows[i]["Segment Name"] != DBNull.Value)
                {
                    string segment1 = Convert.ToString(_SGEN_mainform.dt_segments.Rows[i]["Segment Name"]);

                    if (i == 0)
                    {
                        _SGEN_mainform.tpage_settings.populate_controls_on_page(segment1);
                    }

                    string ProjF = _SGEN_mainform.project_main_folder;


                    if (System.IO.Directory.Exists(ProjF) == true)
                    {

                        if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                        {
                            ProjF = ProjF + "\\";
                        }

                        if (segment1 != "")
                        {
                            if (System.IO.Directory.Exists(ProjF + segment1) == false)
                            {
                                System.IO.Directory.CreateDirectory(ProjF + segment1);

                            }
                        }
                    }


                }
            }

         


            button_Exit_Click(sender, e);
        }

        private void button_Exit_Click(object sender, EventArgs e)
        {
            maximize_Sgen();

            this.Close();
        }

        private void maximize_Sgen()
        {
            foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
            {
                if (Forma1 is Alignment_mdi._SGEN_mainform)
                {
                    Forma1.Focus();
                    Forma1.WindowState = System.Windows.Forms.FormWindowState.Normal;
                }
            }
        }



    }
}
