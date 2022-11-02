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
using Autodesk.AutoCAD.EditorInput;

namespace Alignment_mdi
{
    public partial class Wgen_duplicates : Form
    {



        System.Data.DataTable dt_errors;
        System.Data.DataTable dt_export;


        int extra1 = 6;
        int start_row = 2;
        int end_row = 2;



        string col1 = "PNT";
        string col2 = "NORTHING";
        string col3 = "EASTING";
        string col4 = "ELEVATION";
        string col5 = "FEATURE CODE";
        string col6 = "FILENAME";
        string col7 = "LOCATION";
        string col8 = "NOTES";
        string col9 = "DESCRIPTION";
        string col10 = "MISC1";
        string col11 = "H_ANGLE";
        string col12 = "V_ANGLE";
        string col13 = "MISC4";
        string col14 = "MISC5";
        string col15 = "MISC6";
        string col16 = "MISC7";

        string colGT1 = "MMID";
        string colGT2 = "Pipe";
        string colGT3 = "Heat";
        string colGT4 = "OriginalLength";
        string colGT5 = "NewLength";
        string colGT6 = "WallThickness";
        string colGT7 = "Diameter";
        string colGT8 = "Grade";
        string colGT9 = "Coating";
        string colGT10 = "Manufacture";

        List<string> lista_duplicates;
        List<string> lista_removed;
        List<string> lista_resolved;
        List<int> lista_rows;

        int current_index = -1;

        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_load_all_points);
            lista_butoane.Add(button_all_pts_l);
            lista_butoane.Add(button_all_pts_nl);
            lista_butoane.Add(button_refresh_ws1);
            lista_butoane.Add(button_export_errors_to_xl);
            lista_butoane.Add(button_rem_l);
            lista_butoane.Add(button_rem_nl);
            lista_butoane.Add(button_res_l);
            lista_butoane.Add(button_res_nl);
            lista_butoane.Add(button_load_resolved);
            lista_butoane.Add(button_refresh_resolved);
            lista_butoane.Add(button_refresh_removed);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = false;
            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_load_all_points);
            lista_butoane.Add(button_all_pts_l);
            lista_butoane.Add(button_all_pts_nl);
            lista_butoane.Add(button_refresh_ws1);
            lista_butoane.Add(button_export_errors_to_xl);
            lista_butoane.Add(button_rem_l);
            lista_butoane.Add(button_rem_nl);
            lista_butoane.Add(button_res_l);
            lista_butoane.Add(button_res_nl);
            lista_butoane.Add(button_load_resolved);
            lista_butoane.Add(button_refresh_resolved);
            lista_butoane.Add(button_refresh_removed);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }

        public Wgen_duplicates()
        {
            InitializeComponent();
        }

        private void button_load_all_pts_Click(object sender, EventArgs e)
        {
            Wgen_main_form.dt_all_points = Functions.Creaza_all_points_datatable_structure();
            lista_removed = new List<string>();
            lista_resolved = new List<string>();
            lista_rows = new List<int>();

            if (Wgen_main_form.dt_pt_move == null || Wgen_main_form.dt_pt_move.Rows.Count == 0)
            {
                Wgen_main_form.dt_pt_move = Functions.Creaza_all_points_datatable_structure();
            }
            else
            {
                for (int i = 0; i < Wgen_main_form.dt_pt_move.Rows.Count; ++i)
                {
                    lista_removed.Add(Convert.ToString(Wgen_main_form.dt_pt_move.Rows[i][col1]));
                }

            }

            if (Wgen_main_form.dt_pt_resolved == null || Wgen_main_form.dt_pt_resolved.Rows.Count == 0)
            {
                Wgen_main_form.dt_pt_resolved = Functions.Creaza_all_points_datatable_structure();
            }
            else
            {
                for (int i = 0; i < Wgen_main_form.dt_pt_resolved.Rows.Count; ++i)
                {
                    lista_resolved.Add(Convert.ToString(Wgen_main_form.dt_pt_resolved.Rows[i][col1]));
                }

            }

            make_first_line_invisible();
            textBox_AP_no_duplicates.Text = "";
            textBox_AP_no_rows.Text = "";
            lista_duplicates = new List<string>();


            if (comboBox_all_pts.Text != "")
            {
                string string1 = comboBox_all_pts.Text;
                if (string1.Contains("[") == true && string1.Contains("]") == true)
                {
                    string filename = string1.Substring(string1.IndexOf("]") + 4, string1.Length - (string1.IndexOf("]") + 4));
                    string sheet_name = string1.Substring(1, string1.IndexOf("]") - 1);
                    if (filename.Length > 0 && sheet_name.Length > 0)
                    {
                        set_enable_false();
                        Microsoft.Office.Interop.Excel.Worksheet W1 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, sheet_name);
                        if (W1 != null)
                        {

                            Wgen_main_form.dt_weld_map = null;

                            Wgen_main_form.dt_all_points = Functions.Populate_data_table_from_excel(Wgen_main_form.dt_all_points, W1, start_row, textBox_1.Text, textBox_2.Text, textBox_3.Text, textBox_4.Text, textBox_5.Text, "", "", "", "", "", "", true);



                            Wgen_main_form.dt_all_points.TableName = "TABLA_allpt";
                            end_row = start_row + Wgen_main_form.dt_all_points.Rows.Count - 1;

                            Wgen_main_form.dt_pt_keep = Wgen_main_form.dt_all_points.Copy();
                            Wgen_main_form.dt_pt_keep.TableName = "TABLA_removal";

                            Wgen_main_form.dt_pt_keep.Columns.Add("index1", typeof(int));
                            Wgen_main_form.dt_pt_keep.Columns.Add("duplicate_pt", typeof(string));

                            Wgen_main_form.dt_pt_move.Columns.Add("index1", typeof(int));
                            Wgen_main_form.dt_pt_move.Columns.Add("rempt", typeof(string));

                            Wgen_main_form.dt_pt_resolved.Columns.Add("index1", typeof(int));
                            Wgen_main_form.dt_pt_resolved.Columns.Add("respt", typeof(string));

                            if (Wgen_main_form.dt_pt_keep.Rows.Count > 0)
                            {

                                for (int i = 0; i < Wgen_main_form.dt_pt_keep.Rows.Count; ++i)
                                {
                                    Wgen_main_form.dt_pt_keep.Rows[i]["index1"] = i;
                                }

                                for (int i = Wgen_main_form.dt_pt_keep.Rows.Count - 1; i >= 0; --i)
                                {
                                    if (Wgen_main_form.dt_pt_keep.Rows[i][col1] != DBNull.Value &&
                                       lista_resolved.Contains(Convert.ToString(Wgen_main_form.dt_pt_keep.Rows[i][col1])) == false &&
                                        Wgen_main_form.dt_pt_keep.Rows[i][col5] != DBNull.Value &&
                                        (Convert.ToString(Wgen_main_form.dt_pt_keep.Rows[i][col5]).ToUpper() == "WELD" ||
                                            Convert.ToString(Wgen_main_form.dt_pt_keep.Rows[i][col5]).ToUpper() == "LOOSE_END") &&
                                            Wgen_main_form.dt_pt_keep.Rows[i][col2] != DBNull.Value &&
                                            Wgen_main_form.dt_pt_keep.Rows[i][col3] != DBNull.Value &&
                                            Functions.IsNumeric(Convert.ToString(Wgen_main_form.dt_pt_keep.Rows[i][col2])) == true &&
                                            Functions.IsNumeric(Convert.ToString(Wgen_main_form.dt_pt_keep.Rows[i][col3])) == true)
                                    {

                                    }
                                    else
                                    {
                                        Wgen_main_form.dt_pt_keep.Rows[i].Delete();
                                    }
                                }

                                // Wgen_main_form.dt_pt_keep = Functions.Sort_data_table(Wgen_main_form.dt_pt_keep, col_sta);

                                if (Wgen_main_form.dt_pt_keep.Rows.Count > 0)
                                {
                                    for (int i = 0; i < Wgen_main_form.dt_pt_keep.Rows.Count - 1; ++i)
                                    {
                                        double x1 = Convert.ToDouble(Wgen_main_form.dt_pt_keep.Rows[i][col2]);
                                        double y1 = Convert.ToDouble(Wgen_main_form.dt_pt_keep.Rows[i][col3]);
                                        for (int j = i + 1; j < Wgen_main_form.dt_pt_keep.Rows.Count; ++j)
                                        {
                                            if (Wgen_main_form.dt_pt_keep.Rows[j]["duplicate_pt"] == DBNull.Value)
                                            {
                                                double x2 = Convert.ToDouble(Wgen_main_form.dt_pt_keep.Rows[j][col2]);
                                                double y2 = Convert.ToDouble(Wgen_main_form.dt_pt_keep.Rows[j][col3]);
                                                double d1 = Math.Pow(Math.Pow(x1 - x2, 2) + Math.Pow(y1 - y2, 2), 0.5);
                                                if (d1 < 3)
                                                {
                                                    Wgen_main_form.dt_pt_keep.Rows[j]["duplicate_pt"] = Wgen_main_form.dt_pt_keep.Rows[i][col1];
                                                    if (Wgen_main_form.dt_pt_keep.Rows[i]["duplicate_pt"] == DBNull.Value) Wgen_main_form.dt_pt_keep.Rows[i]["duplicate_pt"] = Wgen_main_form.dt_pt_keep.Rows[i][col1];
                                                }
                                            }
                                        }
                                    }

                                    for (int i = Wgen_main_form.dt_pt_keep.Rows.Count - 1; i >= 0; --i)
                                    {
                                        if (Wgen_main_form.dt_pt_keep.Rows[i]["duplicate_pt"] == DBNull.Value)
                                        {
                                            Wgen_main_form.dt_pt_keep.Rows[i].Delete();
                                        }
                                    }

                                    for (int i = 0; i < Wgen_main_form.dt_pt_keep.Rows.Count; ++i)
                                    {
                                        string duplicate = Convert.ToString(Wgen_main_form.dt_pt_keep.Rows[i]["duplicate_pt"]);
                                        if (lista_duplicates.Contains(duplicate) == false)
                                        {
                                            lista_duplicates.Add(duplicate);
                                        }
                                    }
                                    current_index = 0;
                                    transfer_duplicates_to_panel(Wgen_main_form.dt_pt_keep, current_index);
                                    textBox_AP_no_duplicates.Text = lista_duplicates.Count.ToString();
                                    textBox_AP_no_rows.Text = Wgen_main_form.dt_all_points.Rows.Count.ToString();

                                }

                                button_all_pts_l.Visible = true;
                                button_all_pts_nl.Visible = false;
                            }
                            else
                            {
                                button_all_pts_l.Visible = false;
                                button_all_pts_nl.Visible = true;
                            }
                        }
                        set_enable_true();
                    }
                    else
                    {
                        button_all_pts_l.Visible = false;
                        button_all_pts_nl.Visible = true;
                    }
                }
                else
                {
                    button_all_pts_l.Visible = false;
                    button_all_pts_nl.Visible = true;
                }
            }
            else
            {
                button_all_pts_l.Visible = false;
                button_all_pts_nl.Visible = true;
            }

            //Functions.Transfer_datatable_to_new_excel_spreadsheet( Wgen_main_form. dt_pt_removal);
        }



        private void transfer_duplicates_to_panel(System.Data.DataTable dt1, int index_duplicate)
        {
            if (dt1.Rows.Count > 0)
            {
                textBox1.Visible = true;
                textBox2.Visible = true;
                textBox1.Text = "";
                textBox2.Text = "";
                textBox3.Visible = true;
                textBox3.Text = "";
                textBox3.Visible = true;
                textBox3.Text = "";
                textBox4.Visible = true;
                textBox4.Text = "";

                panel_rb.Visible = true;
                rb1.Checked = true;

                button_z1.Visible = true;

                for (int i = panel_duplicates.Controls.Count - 1; i >= 0; --i)
                {
                    Control ctrl1 = panel_duplicates.Controls[i] as Control;
                    if (ctrl1.Location.Y > textBox1.Location.Y + extra1)
                    {
                        panel_duplicates.Controls.Remove(ctrl1);
                        ctrl1.Dispose();
                    }
                }

                if (dt1.Rows.Count > 1)
                {

                    string duplicate1 = lista_duplicates[index_duplicate];


                    int index_y = 1;

                    for (int i = 0; i < dt1.Rows.Count; ++i)
                    {
                        if (Convert.ToString(dt1.Rows[i][17]) == duplicate1)
                        {
                            string text1 = "";
                            string text2 = "";
                            string text3 = "";
                            string text4 = "";
                            string text5 = "";

                            if (dt1.Rows[i][col1] != DBNull.Value)
                            {
                                text1 = Convert.ToString(dt1.Rows[i][col1]);
                            }



                            if (dt1.Rows[i][col5] != DBNull.Value)
                            {
                                text2 = Convert.ToString(dt1.Rows[i][col5]);
                            }

                            if (dt1.Rows[i][col8] != DBNull.Value)
                            {
                                text4 = Convert.ToString(dt1.Rows[i][col8]);
                            }
                            if (dt1.Rows[i][col6] != DBNull.Value)
                            {
                                text5 = Convert.ToString(dt1.Rows[i][col6]);
                            }
                            if (textBox1.Text == "" && textBox2.Text == "")
                            {
                                textBox1.Text = text1;
                                textBox2.Text = text2;
                                textBox3.Text = text3;
                                textBox3.Text = text4;
                                textBox4.Text = text5;
                                if (lista_removed.Contains(text1) == true)
                                {
                                    rb2.Checked = true;
                                }
                                else
                                {
                                    if (lista_resolved.Contains(text1) == true)
                                    {
                                        rb3.Checked = true;
                                    }
                                    else
                                    {
                                        rb1.Checked = true;
                                    }
                                }

                            }
                            else
                            {
                                Button bt1 = new Button();
                                bt1.Location = new Point(button_z1.Location.X, button_z1.Location.Y + index_y * (button_z1.Height + extra1));
                                bt1.BackColor = button_z1.BackColor;
                                bt1.ForeColor = button_z1.ForeColor;
                                bt1.Font = button_z1.Font;
                                bt1.Size = button_z1.Size;
                                bt1.FlatStyle = button_z1.FlatStyle;
                                bt1.FlatAppearance.BorderColor = button_z1.FlatAppearance.BorderColor;
                                bt1.FlatAppearance.BorderSize = button_z1.FlatAppearance.BorderSize;
                                bt1.FlatAppearance.MouseDownBackColor = button_z1.FlatAppearance.MouseDownBackColor;
                                bt1.FlatAppearance.MouseOverBackColor = button_z1.FlatAppearance.MouseOverBackColor;
                                bt1.BackgroundImage = button_z1.BackgroundImage;
                                bt1.BackgroundImageLayout = button_z1.BackgroundImageLayout;
                                panel_duplicates.Controls.Add(bt1);

                                bt1.Click += delegate (object s, EventArgs e1)
                                {
                                    button_zoom_click(bt1, e1);
                                };

                                TextBox tb1 = new TextBox();
                                tb1.Location = new Point(textBox1.Location.X, textBox1.Location.Y + index_y * (textBox1.Height + extra1));
                                tb1.BackColor = textBox1.BackColor;
                                tb1.ForeColor = textBox1.ForeColor;
                                tb1.Font = textBox1.Font;
                                tb1.Size = textBox1.Size;
                                tb1.ReadOnly = textBox1.ReadOnly;
                                tb1.BorderStyle = textBox1.BorderStyle;
                                tb1.Text = text1;
                                panel_duplicates.Controls.Add(tb1);

                                TextBox tb2 = new TextBox();
                                tb2.Location = new Point(textBox2.Location.X, textBox2.Location.Y + index_y * (textBox2.Height + extra1));
                                tb2.BackColor = textBox2.BackColor;
                                tb2.ForeColor = textBox2.ForeColor;
                                tb2.Font = textBox2.Font;
                                tb2.Size = textBox2.Size;
                                tb2.ReadOnly = textBox2.ReadOnly;
                                tb2.BorderStyle = textBox2.BorderStyle;
                                tb2.Text = text2;
                                panel_duplicates.Controls.Add(tb2);

                                TextBox tb3 = new TextBox();
                                tb3.Location = new Point(textBox3.Location.X, textBox3.Location.Y + index_y * (textBox3.Height + extra1));
                                tb3.BackColor = textBox3.BackColor;
                                tb3.ForeColor = textBox3.ForeColor;
                                tb3.Font = textBox3.Font;
                                tb3.Size = textBox3.Size;
                                tb3.ReadOnly = textBox3.ReadOnly;
                                tb3.BorderStyle = textBox3.BorderStyle;
                                tb3.Text = text3;
                                panel_duplicates.Controls.Add(tb3);

                                TextBox tb4 = new TextBox();
                                tb4.Location = new Point(textBox3.Location.X, textBox3.Location.Y + index_y * (textBox3.Height + extra1));
                                tb4.BackColor = textBox3.BackColor;
                                tb4.ForeColor = textBox3.ForeColor;
                                tb4.Font = textBox3.Font;
                                tb4.Size = textBox3.Size;
                                tb4.ReadOnly = textBox3.ReadOnly;
                                tb4.BorderStyle = textBox3.BorderStyle;
                                tb4.Text = text4;
                                panel_duplicates.Controls.Add(tb4);

                                TextBox tb5 = new TextBox();
                                tb5.Location = new Point(textBox4.Location.X, textBox4.Location.Y + index_y * (textBox4.Height + extra1));
                                tb5.BackColor = textBox4.BackColor;
                                tb5.ForeColor = textBox4.ForeColor;
                                tb5.Font = textBox4.Font;
                                tb5.Size = textBox4.Size;
                                tb5.ReadOnly = textBox4.ReadOnly;
                                tb5.BorderStyle = textBox4.BorderStyle;
                                tb5.Text = text5;
                                panel_duplicates.Controls.Add(tb5);



                                Panel pan3 = new Panel();
                                pan3.Location = new Point(panel_rb.Location.X, panel_rb.Location.Y + index_y * (textBox1.Height + extra1));
                                pan3.BackColor = rb1.BackColor;
                                pan3.ForeColor = rb1.ForeColor;
                                pan3.Font = rb1.Font;
                                pan3.Size = panel_rb.Size;
                                pan3.BorderStyle = panel_rb.BorderStyle;
                                pan3.Anchor = panel_rb.Anchor;
                                panel_duplicates.Controls.Add(pan3);

                                RadioButton rad3 = new RadioButton();
                                rad3.Location = new Point(rb1.Location.X, rb1.Location.Y);
                                rad3.BackColor = rb1.BackColor;
                                rad3.ForeColor = rb1.ForeColor;
                                rad3.Font = rb1.Font;
                                rad3.Size = rb1.Size;
                                rad3.Text = rb1.Text;

                                RadioButton rad4 = new RadioButton();
                                rad4.Location = new Point(rb2.Location.X, rb2.Location.Y);
                                rad4.BackColor = rb2.BackColor;
                                rad4.ForeColor = rb2.ForeColor;
                                rad4.Font = rb2.Font;
                                rad4.Size = rb2.Size;
                                rad4.Text = rb2.Text;

                                RadioButton rad5 = new RadioButton();
                                rad5.Location = new Point(rb3.Location.X, rb3.Location.Y);
                                rad5.BackColor = rb3.BackColor;
                                rad5.ForeColor = rb3.ForeColor;
                                rad5.Font = rb3.Font;
                                rad5.Size = rb3.Size;
                                rad5.Text = rb3.Text;

                                if (lista_removed.Contains(text1) == true)
                                {
                                    rad4.Checked = true;
                                }
                                else
                                {
                                    if (lista_resolved.Contains(text1) == true)
                                    {
                                        rad5.Checked = true;
                                    }
                                    else
                                    {
                                        rad3.Checked = true;
                                    }
                                }

                                pan3.Controls.Add(rad3);
                                pan3.Controls.Add(rad4);
                                pan3.Controls.Add(rad5);

                                rad3.CheckedChanged += delegate (object s, EventArgs e1)
                                {
                                    rb1_CheckedChanged(rad3, e1);
                                };

                                rad4.CheckedChanged += delegate (object s, EventArgs e1)
                                {
                                    rb2_CheckedChanged(rad4, e1);
                                };

                                rad5.CheckedChanged += delegate (object s, EventArgs e1)
                                {
                                    rb3_CheckedChanged(rad5, e1);
                                };


                                ++index_y;
                            }
                        }
                    }
                }
            }
        }
        private void button_zoom_click(object ob, EventArgs e)
        {
            Control ctrl1 = ob as Control;
            if (dt_errors == null || dt_errors.Rows.Count == 0) return;
            if (ctrl1 != null)
            {
                int Y = ctrl1.Location.Y;
                int index1 = (Y - textBox1.Location.Y) / (textBox1.Height + extra1);
                try
                {
                    Microsoft.Office.Interop.Excel.Worksheet W1 = (Microsoft.Office.Interop.Excel.Worksheet)dt_errors.Rows[index1]["w1"];
                    if (W1 != null)
                    {
                        if (dt_errors.Rows[index1]["Excel"] != DBNull.Value)
                        {
                            string adresa = Convert.ToString(dt_errors.Rows[index1]["Excel"]);
                            W1.Activate();
                            W1.Range[adresa].Select();
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void make_first_line_invisible()
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            rb1.Checked = false;
            rb2.Checked = false;

            textBox1.Visible = false;
            textBox2.Visible = false;
            textBox3.Visible = false;
            textBox3.Visible = false;
            textBox4.Visible = false;
            panel_rb.Visible = false;
            button_z1.Visible = false;

            for (int i = panel_duplicates.Controls.Count - 1; i >= 0; --i)
            {
                Control ctrl1 = panel_duplicates.Controls[i] as Control;
                if (ctrl1.Location.Y > textBox1.Location.Y + extra1)
                {
                    panel_duplicates.Controls.Remove(ctrl1);
                    ctrl1.Dispose();
                }
            }
        }



        private void button_refresh_ws1_Click(object sender, EventArgs e)
        {
            Functions.Load_opened_worksheets_to_combobox(comboBox_all_pts);
            if (comboBox_all_pts.Items.Count > 0)
            {
                for (int i = 0; i < comboBox_all_pts.Items.Count; ++i)
                {
                    if (comboBox_all_pts.Items[i].ToString().ToUpper().Contains("ALL_POINTS") == true)
                    {
                        comboBox_all_pts.SelectedIndex = i;
                        i = comboBox_all_pts.Items.Count;
                    }
                }
            }
        }

        private void button_refresh_ws2_Click(object sender, EventArgs e)
        {
            Functions.Load_opened_worksheets_to_combobox(comboBox_removed);
            if (comboBox_removed.Items.Count > 0)
            {
                for (int i = comboBox_removed.Items.Count - 1; i >= 0; --i)
                {
                    if (comboBox_removed.Items[i].ToString().Replace(" ", "").Replace("_", "").ToUpper().Contains("REMOVEDPTS") == false)
                    {
                        comboBox_removed.Items.RemoveAt(i);
                    }
                }
            }
            if (comboBox_removed.Items.Count > 0)
            {
                comboBox_removed.SelectedIndex = 0;
            }
        }

        private void button_refresh_ws3_Click(object sender, EventArgs e)
        {
            Functions.Load_opened_worksheets_to_combobox(comboBox_resolved);
            if (comboBox_resolved.Items.Count > 0)
            {
                for (int i = comboBox_resolved.Items.Count - 1; i >= 0; --i)
                {
                    if (comboBox_resolved.Items[i].ToString().Replace(" ", "").Replace("_", "").ToUpper().Contains("RESOLVEDPTS") == false)
                    {
                        comboBox_resolved.Items.RemoveAt(i);
                    }
                }
            }
            if (comboBox_resolved.Items.Count > 0)
            {
                comboBox_resolved.SelectedIndex = 0;
            }
        }

        private void button_plus1_Click(object sender, EventArgs e)
        {

            if (lista_duplicates != null && lista_duplicates.Count > 0 && current_index != -1)
            {
                if (Wgen_main_form.dt_pt_keep != null && Wgen_main_form.dt_pt_keep.Rows.Count > 0)
                {
                    if (current_index < lista_duplicates.Count - 1)
                    {
                        ++current_index;
                    }
                    else
                    {
                        current_index = 0;
                    }
                    transfer_duplicates_to_panel(Wgen_main_form.dt_pt_keep, current_index);
                }
            }
        }

        private void button_minus1_Click(object sender, EventArgs e)
        {
            if (lista_duplicates != null && lista_duplicates.Count > 0 && current_index != -1)
            {
                if (Wgen_main_form.dt_pt_keep != null && Wgen_main_form.dt_pt_keep.Rows.Count > 0)
                {
                    if (current_index > 0)
                    {
                        --current_index;
                    }
                    else
                    {
                        current_index = lista_duplicates.Count - 1;
                    }
                    transfer_duplicates_to_panel(Wgen_main_form.dt_pt_keep, current_index);
                }
            }
        }

        private void rb1_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton radb1 = sender as RadioButton;

            if (radb1 != null && radb1.Checked == true)
            {
                int y_panel = 4;

                foreach (Control ctrl1 in panel_duplicates.Controls)
                {
                    Panel panel_x = ctrl1 as Panel;
                    if (panel_x != null)
                    {
                        if (panel_x.Controls.Contains(radb1) == true)
                        {
                            y_panel = panel_x.Location.Y + 1;// aici am adaugat 1 pt ca panel.y nu e acelasi ca si textbox.y (4 vs 5)
                        }
                    }
                }

                string point1 = "";
                foreach (Control ctrl1 in panel_duplicates.Controls)
                {
                    TextBox textb1 = ctrl1 as TextBox;
                    if (textb1 != null)
                    {
                        if (textb1.Location.X == textBox1.Location.X && textb1.Location.Y == y_panel)
                        {
                            point1 = textb1.Text;
                        }
                    }
                }

                for (int i = 0; i < Wgen_main_form.dt_pt_keep.Rows.Count; ++i)
                {
                    string point2 = Convert.ToString(Wgen_main_form.dt_pt_keep.Rows[i][col1]);
                    if (point1.ToUpper() == point2.ToUpper())
                    {
                        if (lista_removed.Contains(point1) == false)
                        {
                            Wgen_main_form.dt_pt_move.ImportRow(Wgen_main_form.dt_pt_keep.Rows[i]);
                            lista_removed.Add(point1);
                            lista_rows.Add(Convert.ToInt32(Wgen_main_form.dt_pt_keep.Rows[i]["index1"]));
                        }
                    }
                }



                for (int i = 0; i < Wgen_main_form.dt_pt_keep.Rows.Count; ++i)
                {

                    string point2 = Convert.ToString(Wgen_main_form.dt_pt_keep.Rows[i][col1]);
                    if (point1.ToUpper() == point2.ToUpper())
                    {
                        if (Wgen_main_form.dt_pt_move.Rows.Count > 0)
                        {
                            for (int j = Wgen_main_form.dt_pt_move.Rows.Count - 1; j >= 0; --j)
                            {
                                string point111 = Convert.ToString(Wgen_main_form.dt_pt_move.Rows[j][col1]);
                                if (point1.ToUpper() == point111.ToUpper())
                                {
                                    if (lista_removed.Contains(point111) == true)
                                    {
                                        int idx1 = Convert.ToInt32(Wgen_main_form.dt_pt_move.Rows[j]["index1"]);
                                        Wgen_main_form.dt_pt_move.Rows[j].Delete();
                                        lista_removed.Remove(point111);
                                        lista_rows.Remove(idx1);
                                    }
                                }
                            }
                        }
                    }

                }

                for (int i = 0; i < Wgen_main_form.dt_pt_resolved.Rows.Count; ++i)
                {
                    string point2 = Convert.ToString(Wgen_main_form.dt_pt_resolved.Rows[i][col1]);
                    if (point1.ToUpper() == point2.ToUpper())
                    {

                        if (lista_resolved.Contains(point2) == true)
                        {
                            Wgen_main_form.dt_pt_resolved.Rows[i].Delete();
                            lista_resolved.Remove(point2);
                            i = Wgen_main_form.dt_pt_resolved.Rows.Count;
                        }

                    }
                }

                textBox_AP_no_resolved.Text = Convert.ToString(Wgen_main_form.dt_pt_resolved.Rows.Count);
                textBox_AP_no_removed.Text = Convert.ToString(Wgen_main_form.dt_pt_move.Rows.Count);
            }
        }

        private void rb2_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton radb2 = sender as RadioButton;

            if (radb2 != null && radb2.Checked == true)
            {
                int y_panel = 4;

                foreach (Control ctrl1 in panel_duplicates.Controls)
                {
                    Panel panel_x = ctrl1 as Panel;
                    if (panel_x != null)
                    {
                        if (panel_x.Controls.Contains(radb2) == true)
                        {
                            y_panel = panel_x.Location.Y + 1;// aici am adaugat 1 pt ca panel.y nu e acelasi ca si textbox.y (4 vs 5)
                        }
                    }
                }

                string point1 = "";
                foreach (Control ctrl1 in panel_duplicates.Controls)
                {
                    TextBox textb1 = ctrl1 as TextBox;
                    if (textb1 != null)
                    {
                        if (textb1.Location.X == textBox1.Location.X && textb1.Location.Y == y_panel)
                        {
                            point1 = textb1.Text;
                        }
                    }
                }

                for (int i = 0; i < Wgen_main_form.dt_pt_keep.Rows.Count; ++i)
                {
                    string point2 = Convert.ToString(Wgen_main_form.dt_pt_keep.Rows[i][col1]);
                    if (point1.ToUpper() == point2.ToUpper())
                    {
                        if (lista_removed.Contains(point1) == false)
                        {
                            Wgen_main_form.dt_pt_move.ImportRow(Wgen_main_form.dt_pt_keep.Rows[i]);
                            lista_removed.Add(point1);
                            lista_rows.Add(Convert.ToInt32(Wgen_main_form.dt_pt_keep.Rows[i]["index1"]));
                        }
                    }
                }

                for (int i = 0; i < Wgen_main_form.dt_pt_resolved.Rows.Count; ++i)
                {
                    string point2 = Convert.ToString(Wgen_main_form.dt_pt_resolved.Rows[i][col1]);
                    if (point1.ToUpper() == point2.ToUpper())
                    {
                        if (radb2.Checked == true)
                        {
                            if (lista_resolved.Contains(point2) == true)
                            {
                                Wgen_main_form.dt_pt_resolved.Rows[i].Delete();
                                lista_resolved.Remove(point2);
                                i = Wgen_main_form.dt_pt_resolved.Rows.Count;
                            }
                        }
                    }
                }

                textBox_AP_no_resolved.Text = Convert.ToString(Wgen_main_form.dt_pt_resolved.Rows.Count);
                textBox_AP_no_removed.Text = Convert.ToString(Wgen_main_form.dt_pt_move.Rows.Count);
            }
        }

        private void rb3_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton radb3 = sender as RadioButton;
            if (radb3 != null && radb3.Checked == true)
            {
                int y_panel = 4;

                foreach (Control ctrl1 in panel_duplicates.Controls)
                {
                    Panel panel_x = ctrl1 as Panel;
                    if (panel_x != null)
                    {
                        if (panel_x.Controls.Contains(radb3) == true)
                        {
                            y_panel = panel_x.Location.Y + 1;// aici am adaugat 1 pt ca panel.y nu e acelasi ca si textbox.y (4 vs 5)
                        }
                    }
                }

                string point1 = "";
                foreach (Control ctrl1 in panel_duplicates.Controls)
                {
                    TextBox textb1 = ctrl1 as TextBox;
                    if (textb1 != null)
                    {
                        if (textb1.Location.X == textBox1.Location.X && textb1.Location.Y == y_panel)
                        {
                            point1 = textb1.Text;
                        }
                    }
                }



                for (int i = 0; i < Wgen_main_form.dt_pt_keep.Rows.Count; ++i)
                {
                    string point2 = Convert.ToString(Wgen_main_form.dt_pt_keep.Rows[i][col1]);
                    if (point1.ToUpper() == point2.ToUpper())
                    {
                        if (radb3.Checked == true)
                        {
                            if (lista_resolved.Contains(point1) == false)
                            {
                                Wgen_main_form.dt_pt_resolved.ImportRow(Wgen_main_form.dt_pt_keep.Rows[i]);
                                lista_resolved.Add(point1);
                            }
                        }

                        if (Wgen_main_form.dt_pt_move.Rows.Count > 0)
                        {
                            for (int j = Wgen_main_form.dt_pt_move.Rows.Count - 1; j >= 0; --j)
                            {
                                string point111 = Convert.ToString(Wgen_main_form.dt_pt_move.Rows[j][col1]);
                                if (point1.ToUpper() == point111.ToUpper())
                                {
                                    if (lista_removed.Contains(point111) == true)
                                    {
                                        int idx1 = Convert.ToInt32(Wgen_main_form.dt_pt_move.Rows[j]["index1"]);
                                        Wgen_main_form.dt_pt_move.Rows[j].Delete();
                                        lista_removed.Remove(point111);
                                        lista_rows.Remove(idx1);
                                    }
                                }
                            }
                        }
                    }
                }


                textBox_AP_no_resolved.Text = Convert.ToString(Wgen_main_form.dt_pt_resolved.Rows.Count);
                textBox_AP_no_removed.Text = Convert.ToString(Wgen_main_form.dt_pt_move.Rows.Count);
            }
        }

        private void button_export_errors_to_xl_Click(object sender, EventArgs e)
        {
            if (Wgen_main_form.dt_pt_move.Rows.Count > 0)
            {
                Wgen_main_form.dt_pt_move = Functions.Sort_data_table(Wgen_main_form.dt_pt_move, "index1");
                for (int i = Wgen_main_form.dt_pt_move.Rows.Count - 1; i >= 0; --i)
                {
                    if (Wgen_main_form.dt_pt_move.Rows[i]["index1"] != DBNull.Value)
                    {
                        int index1 = Convert.ToInt32(Wgen_main_form.dt_pt_move.Rows[i]["index1"]);
                        Wgen_main_form.dt_all_points.Rows[index1].Delete();
                    }
                }

                if (Wgen_main_form.dt_pt_move.Columns.Contains("index1") == true) Wgen_main_form.dt_pt_move.Columns.Remove("index1");
                if (Wgen_main_form.dt_pt_move.Columns.Contains("rempt") == true) Wgen_main_form.dt_pt_move.Columns.Remove("rempt");

                string filename_move = "";
                string sheetname_move = "";

                string string1 = comboBox_removed.Text;

                if (string1.Contains("[") == true && string1.Contains("]") == true)
                {
                    filename_move = string1.Substring(string1.IndexOf("]") + 4, string1.Length - (string1.IndexOf("]") + 4));
                    sheetname_move = string1.Substring(1, string1.IndexOf("]") - 1);
                }

                if (filename_move != "")
                {
                    Functions.Transfer_datatable_to_existing_excel_spreadsheet_by_name(Wgen_main_form.dt_pt_move, filename_move, sheetname_move, false, start_row, end_row);
                }
                else
                {
                    Functions.Transfer_datatable_to_new_excel_spreadsheet_named(Wgen_main_form.dt_pt_move, "RemovedPTS");
                }
            }

            if (Wgen_main_form.dt_pt_resolved.Rows.Count > 0)
            {

                if (Wgen_main_form.dt_pt_resolved.Columns.Contains("index1") == true) Wgen_main_form.dt_pt_resolved.Columns.Remove("index1");
                if (Wgen_main_form.dt_pt_resolved.Columns.Contains("respt") == true) Wgen_main_form.dt_pt_resolved.Columns.Remove("respt");

                string filename_resolved = "";
                string sheetname_resolved = "";

                string string1 = comboBox_resolved.Text;

                if (string1.Contains("[") == true && string1.Contains("]") == true)
                {
                    filename_resolved = string1.Substring(string1.IndexOf("]") + 4, string1.Length - (string1.IndexOf("]") + 4));
                    sheetname_resolved = string1.Substring(1, string1.IndexOf("]") - 1);
                }

                if (filename_resolved != "")
                {
                    Functions.Transfer_datatable_to_existing_excel_spreadsheet_by_name(Wgen_main_form.dt_pt_resolved, filename_resolved, sheetname_resolved, false, start_row, end_row);
                }
                else
                {
                    Functions.Transfer_datatable_to_new_excel_spreadsheet_named(Wgen_main_form.dt_pt_resolved, "ResolvedPTS");
                }
            }

            string filename_allpt = "";
            string sheetname_allpt = "";

            string string2 = comboBox_all_pts.Text;

            if (string2.Contains("[") == true && string2.Contains("]") == true)
            {
                filename_allpt = string2.Substring(string2.IndexOf("]") + 4, string2.Length - (string2.IndexOf("]") + 4));
                sheetname_allpt = string2.Substring(1, string2.IndexOf("]") - 1);
            }
            if (filename_allpt != "")
            {
                Functions.create_backup(filename_allpt);
                Functions.erase_rows_from_excel(filename_allpt, sheetname_allpt, lista_rows, start_row);
            }

            Wgen_main_form.dt_pt_move = Functions.Creaza_all_points_datatable_structure();
            Wgen_main_form.dt_pt_resolved = Functions.Creaza_all_points_datatable_structure();
            Wgen_main_form.dt_pt_keep = Functions.Creaza_all_points_datatable_structure();
            Wgen_main_form.dt_all_points = Functions.Creaza_all_points_datatable_structure();
            lista_rows = new List<int>();
            lista_removed = new List<string>();
            lista_duplicates = new List<string>();
            lista_resolved = new List<string>();
            make_first_line_invisible();
            textBox_AP_no_removed.Text = "";
            textBox_AP_no_rows.Text = "";
            textBox_AP_no_duplicates.Text = "";
            button_all_pts_l.Visible = false;
            button_all_pts_nl.Visible = true;
            button_rem_l.Visible = false;
            button_rem_nl.Visible = true;
            button_res_l.Visible = false;
            button_res_nl.Visible = true;
            comboBox_all_pts.Items.Clear();
            comboBox_removed.Items.Clear();
            comboBox_resolved.Items.Clear();

        }

        private void button_load_removed_Click(object sender, EventArgs e)
        {
            Wgen_main_form.dt_pt_move = Functions.Creaza_all_points_datatable_structure();
            lista_removed = new List<string>();

            if (comboBox_removed.Text != "")
            {
                string string1 = comboBox_removed.Text;
                if (string1.Contains("[") == true && string1.Contains("]") == true)
                {
                    string filename = string1.Substring(string1.IndexOf("]") + 4, string1.Length - (string1.IndexOf("]") + 4));
                    string sheet_name = string1.Substring(1, string1.IndexOf("]") - 1);
                    if (filename.Length > 0 && sheet_name.Length > 0)
                    {
                        set_enable_false();
                        Microsoft.Office.Interop.Excel.Worksheet W1 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, sheet_name);
                        if (W1 != null)
                        {
                            Wgen_main_form.dt_pt_move = Functions.Populate_data_table_from_excel(Wgen_main_form.dt_pt_move, W1, start_row, textBox_1.Text, textBox_2.Text, textBox_3.Text, textBox_4.Text, textBox_5.Text, textBox_6.Text, "", "", "", "", "", false);
                            Wgen_main_form.dt_pt_move.TableName = "rempt";

                            if (Wgen_main_form.dt_pt_move.Rows.Count > 0)
                            {
                                button_rem_l.Visible = true;
                                button_rem_nl.Visible = false;
                            }
                            else
                            {
                                button_rem_l.Visible = false;
                                button_rem_nl.Visible = true;
                            }
                        }
                        set_enable_true();
                    }
                    else
                    {
                        button_rem_l.Visible = false;
                        button_rem_nl.Visible = true;
                    }
                }
                else
                {
                    button_rem_l.Visible = false;
                    button_rem_nl.Visible = true;
                }
            }
            else
            {
                button_rem_l.Visible = false;
                button_rem_nl.Visible = true;
            }

            textBox_AP_no_removed.Text = Convert.ToString(Wgen_main_form.dt_pt_move.Rows.Count);

            //Functions.Transfer_datatable_to_new_excel_spreadsheet( Wgen_main_form. dt_pt_removal);
        }
        private void button_load_resolved_Click(object sender, EventArgs e)
        {
            Wgen_main_form.dt_pt_resolved = Functions.Creaza_all_points_datatable_structure();
            lista_resolved = new List<string>();

            if (comboBox_resolved.Text != "")
            {
                string string1 = comboBox_resolved.Text;
                if (string1.Contains("[") == true && string1.Contains("]") == true)
                {
                    string filename = string1.Substring(string1.IndexOf("]") + 4, string1.Length - (string1.IndexOf("]") + 4));
                    string sheet_name = string1.Substring(1, string1.IndexOf("]") - 1);
                    if (filename.Length > 0 && sheet_name.Length > 0)
                    {
                        set_enable_false();
                        Microsoft.Office.Interop.Excel.Worksheet W1 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, sheet_name);
                        if (W1 != null)
                        {
                            Wgen_main_form.dt_pt_resolved = Functions.Populate_data_table_from_excel(Wgen_main_form.dt_pt_resolved, W1, start_row, textBox_1.Text, textBox_2.Text, textBox_3.Text, textBox_4.Text, textBox_5.Text, textBox_6.Text, "", "", "", "", "", false);
                            Wgen_main_form.dt_pt_resolved.TableName = "respt";

                            if (Wgen_main_form.dt_pt_resolved.Rows.Count > 0)
                            {
                                button_res_l.Visible = true;
                                button_res_nl.Visible = false;
                            }
                            else
                            {
                                button_res_l.Visible = false;
                                button_res_nl.Visible = true;
                            }
                        }
                        set_enable_true();
                    }
                    else
                    {
                        button_res_l.Visible = false;
                        button_res_nl.Visible = true;
                    }
                }
                else
                {
                    button_res_l.Visible = false;
                    button_res_nl.Visible = true;
                }
            }
            else
            {
                button_res_l.Visible = false;
                button_res_nl.Visible = true;
            }

            textBox_AP_no_resolved.Text = Convert.ToString(Wgen_main_form.dt_pt_resolved.Rows.Count);

            //Functions.Transfer_datatable_to_new_excel_spreadsheet( Wgen_main_form. dt_pt_removal);
        }

    }
}
