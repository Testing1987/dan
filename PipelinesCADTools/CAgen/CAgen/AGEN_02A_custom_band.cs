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
    public partial class AGEN_custom_band_form : Form
    {
        bool clickdragdown;
        Point lastLocation;

        static bool Freeze_operations = false;
        int nr_bands = 0;
        double panelh = 0;

        public AGEN_custom_band_form()
        {
            InitializeComponent();
            panelh = panel_custom_band.Height;
            add_controls_to_custom_form();
        }

        private void add_controls_to_custom_form()
        {



            if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count > 0)
            {
                textBox_name1.Text = _AGEN_mainform.Data_Table_custom_bands.Rows[0]["band_name"].ToString();

                if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count > 1)
                {
                    for (int i = 1; i < _AGEN_mainform.Data_Table_custom_bands.Rows.Count; ++i)
                    {
                        nr_bands = i - 1;

                        string nume_custom = _AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"].ToString();
                        add_new_control_row(new object(), new EventArgs(), nume_custom, Alignment_mdi.Properties.Resources.check);
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
        private void button_Exit_Click(object sender, EventArgs e)
        {
            maximize_agen();

            this.Close();
        }
        private void button_minimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void button_add_custom_control_Click(object sender, EventArgs e)
        {
            if (nr_bands + 2 > 16)
            {
                MessageBox.Show("you can't have more than 16 custom bands");
                return;
            }

            add_new_control_row(sender, e, "", Alignment_mdi.Properties.Resources.selectbluexs);
        }

        private void add_new_control_row(object sender, EventArgs e, string nume_band, System.Drawing.Bitmap bitmap1)
        {

            Label label1 = new Label();
            label1.Location = new Point(label_name1.Location.X, label_name1.Location.Y + (nr_bands + 1) * (textBox_name1.Height + 4));
            label1.BackColor = label_name1.BackColor;
            label1.ForeColor = label_name1.ForeColor;
            label1.Font = label_name1.Font;
            label1.Text = "Custom Band " + (nr_bands + 2).ToString() + " Name";
            label1.Size = new Size(label_name1.Size.Width + 20, label_name1.Size.Height);
            panel_custom_band.Controls.Add(label1);

            TextBox textbox1 = new TextBox();
            textbox1.Location = new Point(textBox_name1.Location.X, textBox_name1.Location.Y + (nr_bands + 1) * (textBox_name1.Height + 4));
            textbox1.BackColor = textBox_name1.BackColor;
            textbox1.ForeColor = textBox_name1.ForeColor;
            textbox1.Font = textBox_name1.Font;
            textbox1.Size = textBox_name1.Size;
            textbox1.Text = nume_band;
            panel_custom_band.Controls.Add(textbox1);







            button_create_bands.Location = new Point(button_create_bands.Location.X, button_create_bands.Location.Y + textBox_name1.Height + 4);

            panel_custom_band.Size = new Size(panel_custom_band.Width, panel_custom_band.Height + textBox_name1.Height + 4);
            this.Size = new Size(this.Width, this.Height + textBox_name1.Height + 4);
            nr_bands = nr_bands + 1;
        }

        private void button_remove_custom_Click(object sender, EventArgs e)
        {
            if (panel_custom_band.Height > panelh)
            {

                double Yt = textBox_name1.Location.Y + nr_bands * (textBox_name1.Height + 4);
                double Yl = label_name1.Location.Y + nr_bands * (textBox_name1.Height + 4);

                int y_textbox = -1;

                for (int i = panel_custom_band.Controls.Count - 1; i >= 0; --i)
                {
                    Control CTRL1 = panel_custom_band.Controls[i];
                    if (CTRL1.Location.Y == Yl || CTRL1.Location.Y == Yt)
                    {
                        TextBox tb1 = panel_custom_band.Controls[i] as TextBox;
                        if (tb1 != null)
                        {
                            y_textbox = tb1.Location.Y;
                        }
                        panel_custom_band.Controls.Remove(CTRL1);
                    }
                }

                nr_bands = nr_bands - 1;

                button_create_bands.Location = new Point(button_create_bands.Location.X, button_create_bands.Location.Y - textBox_name1.Height - 4);

                panel_custom_band.Size = new Size(panel_custom_band.Width, panel_custom_band.Height - textBox_name1.Height - 4);
                this.Size = new Size(this.Width, this.Height - textBox_name1.Height - 4);






            }
        }

        private void button_create_custom_bands_Click(object sender, EventArgs e)
        {
          
            List<string> lista1 = new List<string>();
            List<string> lista2 = new List<string>();
            List<string> lista3 = new List<string>();
            for (int i = 0; i < panel_custom_band.Controls.Count; ++i)
            {
                TextBox tb1 = panel_custom_band.Controls[i] as TextBox;
                if (tb1 != null)
                {
                    if(tb1.Text!= "")
                    {
                    bool exista = false;
                    if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count > 0)
                    {
                        for (int j = 0; j < _AGEN_mainform.Data_Table_custom_bands.Rows.Count; ++j)
                        {
                            string nume_existent = Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[j]["band_name"]);
                            if (nume_existent.ToUpper() == tb1.Text.ToUpper())
                            {
                                exista = true;
                            }
                        }
                    }

                    if (exista == false)
                    {
                        if (tb1.Text != null)
                        {
                            lista1.Add(tb1.Text.ToUpper());
                            lista2.Add(tb1.Text);
                        }
                    }
                    lista3.Add(tb1.Text.ToUpper());
                    }

                }
            }


            for (int j = _AGEN_mainform.Data_Table_custom_bands.Rows.Count-1; j>=0; --j)
            {
                string nume_existent = Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[j]["band_name"]);
                if(lista3.Contains(nume_existent.ToUpper())==false)
                {
                    _AGEN_mainform.Data_Table_custom_bands.Rows[j].Delete();
                }
            }

            if (lista1.Count > 0)
            {
                for (int i = 0; i < lista1.Count; ++i)
                {
                    bool adauga = true;

                    if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count > 0)
                    {
                        for (int j = 0; j < _AGEN_mainform.Data_Table_custom_bands.Rows.Count; ++j)
                        {
                            string nume_existent = Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[j]["band_name"]);
                            if (nume_existent.ToUpper() == lista1[i])
                            {
                                adauga = false;
                            }
                        }
                    }

                    if (adauga == true)
                    {
                        _AGEN_mainform.Data_Table_custom_bands.Rows.Add();
                        _AGEN_mainform.Data_Table_custom_bands.Rows[_AGEN_mainform.Data_Table_custom_bands.Rows.Count - 1]["band_name"] = lista2[i];
                    }
                }


            }

            if (lista3.Count > 0)
            {
                transfera_custom_band_settings_to_excel();
                _AGEN_mainform.tpage_setup.display_checkboxes_into_generation_page();
                _AGEN_mainform.tpage_viewport_settings.creeaza_display_data_table(Functions.Creaza_lista_regular_vp_picked(), Functions.Creaza_lista_custom_vp_picked(), Functions.Creaza_lista_custom_vp_extra_picked());
            }


            maximize_agen();
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

        static public void transfera_custom_band_settings_to_excel()
        {
            if (_AGEN_mainform.Data_Table_custom_bands != null)
            {
                if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count > 0)
                {
                    Functions.Kill_excel();

                    string cfg1 = System.IO.Path.GetFileName(_AGEN_mainform.config_path);
                    if (Functions.Get_if_workbook_is_open_in_Excel(cfg1) == true)
                    {
                        MessageBox.Show("Please close the " + cfg1 + " file");
                        return;
                    }


                    if (Freeze_operations == false)
                    {
                        Freeze_operations = true;
                        try
                        {
                            Microsoft.Office.Interop.Excel.Application Excel1 = null;
                           
                            try
                            {
                                Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                            }
                            catch (System.Exception ex)
                            {
                                Excel1 = new Microsoft.Office.Interop.Excel.Application();
                               
                            }

                            if (Excel1.Workbooks.Count==0) Excel1.Visible = _AGEN_mainform.ExcelVisible;

                            if (System.IO.File.Exists(_AGEN_mainform.config_path) == true)
                            {

                                Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(_AGEN_mainform.config_path);

                                Microsoft.Office.Interop.Excel.Worksheet W1 = null;

                                foreach (Microsoft.Office.Interop.Excel.Worksheet wsh1 in Workbook1.Worksheets)
                                {
                                    if (wsh1.Name == "Custom_band_data")
                                    {
                                        W1 = wsh1;
                                    }
                                }

                                if (W1 == null)
                                {
                                    W1 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[Workbook1.Worksheets.Count], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                                    W1.Name = "Custom_band_data";

                                }


                                W1.Columns["A:XX"].Delete();


                                try
                                {

                                    int maxRows = _AGEN_mainform.Data_Table_custom_bands.Rows.Count;
                                    int maxCols = _AGEN_mainform.Data_Table_custom_bands.Columns.Count;

                                    Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[2, 1], W1.Cells[maxRows + 1, maxCols]];
                                    object[,] values1 = new object[maxRows, maxCols];

                                    for (int i = 0; i < maxRows; ++i)
                                    {
                                        for (int j = 0; j < maxCols; ++j)
                                        {
                                            if (_AGEN_mainform.Data_Table_custom_bands.Rows[i][j] != DBNull.Value)
                                            {
                                                values1[i, j] = _AGEN_mainform.Data_Table_custom_bands.Rows[i][j];
                                            }
                                        }
                                    }

                                    for (int i = 0; i < _AGEN_mainform.Data_Table_custom_bands.Columns.Count; ++i)
                                    {
                                        W1.Cells[1, i + 1].value2 = _AGEN_mainform.Data_Table_custom_bands.Columns[i].ColumnName;
                                    }

                                    range1.Cells.NumberFormat = "@";
                                    range1.Value2 = values1;

                                    Functions.Color_border_range_inside(range1, 0);

                                    Workbook1.Save();
                                    Workbook1.Close();
                                    if (Excel1.Workbooks.Count==0)
                                    {
                                        Excel1.Quit();
                                    }
                                    else
                                    {
                                        Excel1.Visible = true;
                                    }


                                }
                                catch (System.Exception ex)
                                {
                                    System.Windows.Forms.MessageBox.Show(ex.Message);
                                }
                                finally
                                {
                                    if (W1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W1);
                                    if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                                    if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                                }

                            }



                        }
                        catch (System.Exception ex)
                        {
                            System.Windows.Forms.MessageBox.Show(ex.Message);

                        }
                        Freeze_operations = false;

                    }


                }
            }

        }

    }
}
