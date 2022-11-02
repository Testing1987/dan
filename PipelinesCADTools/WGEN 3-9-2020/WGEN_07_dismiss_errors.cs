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
using Font = System.Drawing.Font;

namespace Alignment_mdi
{
    public partial class Wgen_dismiss_errors : Form
    {


        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();

            lista_butoane.Add(button_minus1);
            lista_butoane.Add(button_plus1);
            lista_butoane.Add(button_release_errors);
            lista_butoane.Add(button_release_errors);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = false;
            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Button> lista_butoane = new List<Button>();
            lista_butoane.Add(button_minus1);
            lista_butoane.Add(button_plus1);
            lista_butoane.Add(button_release_errors);
            lista_butoane.Add(button_release_errors);

            foreach (System.Windows.Forms.Button bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }

        public Wgen_dismiss_errors()
        {
            InitializeComponent();
        }

        public void display_dt_to_datagridView(System.Data.DataTable dt1)
        {
            dataGridView_dissmissed_errors.DataSource = dt1;
            if (dt1 != null)
            {
                dataGridView_dissmissed_errors.Columns[0].Width = 150;
                dataGridView_dissmissed_errors.Columns[1].Width = 175;
                dataGridView_dissmissed_errors.Columns[2].Width = 400;

            }

            dataGridView_dissmissed_errors.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_dissmissed_errors.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView_dissmissed_errors.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_dissmissed_errors.DefaultCellStyle.ForeColor = Color.White;
            dataGridView_dissmissed_errors.EnableHeadersVisualStyles = false;
        }



        private void button_release_errors_Click(object sender, EventArgs e)
        {
            try
            {
                List<int> lista1 = new List<int>();

                foreach (DataGridViewCell cell1 in dataGridView_dissmissed_errors.SelectedCells)
                {
                    int row_index = cell1.RowIndex;
                    if (lista1.Contains(row_index) == false)
                    {
                        lista1.Add(row_index);
                    }
                }

                if (lista1.Count > 0)
                {
                    set_enable_false();

                    if (label1.Text == "Ground Tally")
                    {
                        for (int i = 0; i < lista1.Count; ++i)
                        {
                            string val0 = Convert.ToString(dataGridView_dissmissed_errors.Rows[lista1[i]].Cells[0].Value);
                            string val1 = Convert.ToString(dataGridView_dissmissed_errors.Rows[lista1[i]].Cells[1].Value);
                            string val2 = Convert.ToString(dataGridView_dissmissed_errors.Rows[lista1[i]].Cells[2].Value);

                            for (int j = 0; j < Wgen_main_form.tpage_pipe_tally.dt_dismissed_errors.Rows.Count; ++j)
                            {
                                string val0a = Convert.ToString(Wgen_main_form.tpage_pipe_tally.dt_dismissed_errors.Rows[j][0]);
                                string val1a = Convert.ToString(Wgen_main_form.tpage_pipe_tally.dt_dismissed_errors.Rows[j][1]);
                                string val2a = Convert.ToString(Wgen_main_form.tpage_pipe_tally.dt_dismissed_errors.Rows[j][2]);

                                if (val0a == val0 && val1a == val1 && val2a == val2)
                                {
                                    Wgen_main_form.tpage_pipe_tally.dt_dismissed_errors.Rows[j].Delete();
                                    j = Wgen_main_form.tpage_pipe_tally.dt_dismissed_errors.Rows.Count;
                                }

                            }

                        }
                        if (Wgen_main_form.tpage_pipe_tally.filename != "" && Wgen_main_form.tpage_pipe_tally.dismiss_errors_tab != "")
                        {
                            Functions.Transfer_datatable_to_existing_excel_spreadsheet_by_name(Wgen_main_form.tpage_pipe_tally.dt_dismissed_errors,
                                                                                                                  Wgen_main_form.tpage_pipe_tally.filename,
                                                                                                                                 Wgen_main_form.tpage_pipe_tally.dismiss_errors_tab, false);
                        }
                        display_dt_to_datagridView(Wgen_main_form.tpage_pipe_tally.dt_dismissed_errors);
                    }
                    if (label1.Text == "Pipe Manifest")
                    {
                        for (int i = 0; i < lista1.Count; ++i)
                        {
                            string val0 = Convert.ToString(dataGridView_dissmissed_errors.Rows[lista1[i]].Cells[0].Value);
                            string val1 = Convert.ToString(dataGridView_dissmissed_errors.Rows[lista1[i]].Cells[1].Value);
                            string val2 = Convert.ToString(dataGridView_dissmissed_errors.Rows[lista1[i]].Cells[2].Value);

                            for (int j = 0; j < Wgen_main_form.tpage_pipe_manifest.dt_dismissed_errors.Rows.Count; ++j)
                            {
                                string val0a = Convert.ToString(Wgen_main_form.tpage_pipe_manifest.dt_dismissed_errors.Rows[j][0]);
                                string val1a = Convert.ToString(Wgen_main_form.tpage_pipe_manifest.dt_dismissed_errors.Rows[j][1]);
                                string val2a = Convert.ToString(Wgen_main_form.tpage_pipe_manifest.dt_dismissed_errors.Rows[j][2]);

                                if (val0a == val0 && val1a == val1 && val2a == val2)
                                {
                                    Wgen_main_form.tpage_pipe_manifest.dt_dismissed_errors.Rows[j].Delete();
                                    j = Wgen_main_form.tpage_pipe_manifest.dt_dismissed_errors.Rows.Count;
                                }

                            }

                        }
                        if (Wgen_main_form.tpage_pipe_manifest.filename != "" && Wgen_main_form.tpage_pipe_manifest.dismiss_errors_tab != "")
                        {
                            Functions.Transfer_datatable_to_existing_excel_spreadsheet_by_name(Wgen_main_form.tpage_pipe_manifest.dt_dismissed_errors,
                                                                                                                  Wgen_main_form.tpage_pipe_manifest.filename,
                                                                                                                                 Wgen_main_form.tpage_pipe_manifest.dismiss_errors_tab, false);
                        }
                        display_dt_to_datagridView(Wgen_main_form.tpage_pipe_manifest.dt_dismissed_errors);
                    }

                    if (label1.Text == "All Points")
                    {
                        for (int i = 0; i < lista1.Count; ++i)
                        {
                            string val0 = Convert.ToString(dataGridView_dissmissed_errors.Rows[lista1[i]].Cells[0].Value);
                            string val1 = Convert.ToString(dataGridView_dissmissed_errors.Rows[lista1[i]].Cells[1].Value);
                            string val2 = Convert.ToString(dataGridView_dissmissed_errors.Rows[lista1[i]].Cells[2].Value);

                            for (int j = 0; j < Wgen_main_form.tpage_allpts.dt_dismissed_errors.Rows.Count; ++j)
                            {
                                string val0a = Convert.ToString(Wgen_main_form.tpage_allpts.dt_dismissed_errors.Rows[j][0]);
                                string val1a = Convert.ToString(Wgen_main_form.tpage_allpts.dt_dismissed_errors.Rows[j][1]);
                                string val2a = Convert.ToString(Wgen_main_form.tpage_allpts.dt_dismissed_errors.Rows[j][2]);

                                if (val0a == val0 && val1a == val1 && val2a == val2)
                                {
                                    Wgen_main_form.tpage_allpts.dt_dismissed_errors.Rows[j].Delete();
                                    j = Wgen_main_form.tpage_allpts.dt_dismissed_errors.Rows.Count;
                                }

                            }

                        }
                        if (Wgen_main_form.tpage_allpts.filename != "" && Wgen_main_form.tpage_allpts.dismiss_errors_tab != "")
                        {
                            Functions.Transfer_datatable_to_existing_excel_spreadsheet_by_name(Wgen_main_form.tpage_allpts.dt_dismissed_errors,
                                                                                                                  Wgen_main_form.tpage_allpts.filename,
                                                                                                                                 Wgen_main_form.tpage_allpts.dismiss_errors_tab, false);
                        }
                        display_dt_to_datagridView(Wgen_main_form.tpage_allpts.dt_dismissed_errors);
                    }

                }






            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            set_enable_true();

        }

        private void button_minus1_Click(object sender, EventArgs e)
        {
            switch (label1.Text)
            {
                case "Ground Tally":
                    label1.Text = "Pipe Manifest";
                    Wgen_main_form.tpage_dismiss_errors.display_dt_to_datagridView(Wgen_main_form.tpage_pipe_manifest.dt_dismissed_errors);
                    break;
                case "All Points":
                    label1.Text = "Ground Tally";
                    Wgen_main_form.tpage_dismiss_errors.display_dt_to_datagridView(Wgen_main_form.tpage_pipe_tally.dt_dismissed_errors);
                    break;
                case "Pipe Manifest":
                    label1.Text = "All Points";
                    Wgen_main_form.tpage_dismiss_errors.display_dt_to_datagridView(Wgen_main_form.tpage_allpts.dt_dismissed_errors);
                    break;
                default:
                    break;
            }
        }

        private void button_plus1_Click(object sender, EventArgs e)
        {
            switch (label1.Text)
            {
                case "Ground Tally":
                    label1.Text = "All Points";
                    Wgen_main_form.tpage_dismiss_errors.display_dt_to_datagridView(Wgen_main_form.tpage_allpts.dt_dismissed_errors);

                    break;
                case "All Points":
                    label1.Text = "Pipe Manifest";
                    Wgen_main_form.tpage_dismiss_errors.display_dt_to_datagridView(Wgen_main_form.tpage_pipe_manifest.dt_dismissed_errors);
                    break;
                case "Pipe Manifest":
                    label1.Text = "Ground Tally";
                    Wgen_main_form.tpage_dismiss_errors.display_dt_to_datagridView(Wgen_main_form.tpage_pipe_tally.dt_dismissed_errors);

                    break;
                default:
                    break;
            }
        }
    }
}
