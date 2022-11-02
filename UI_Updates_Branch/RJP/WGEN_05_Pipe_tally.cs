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

namespace Alignment_mdi
{
    public partial class Wgen_pipe_tally : Form
    {
        int start_row = 2;

        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_create_pipe_tally);
            lista_butoane.Add(button_pipe_tally_l);
            lista_butoane.Add(button_pipe_tally_nl);
            lista_butoane.Add(button_refresh_ws1);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = false;
            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_create_pipe_tally);
            lista_butoane.Add(button_pipe_tally_l);
            lista_butoane.Add(button_pipe_tally_nl);
            lista_butoane.Add(button_refresh_ws1);


            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }

        public Wgen_pipe_tally()
        {
            InitializeComponent();
        }

        private void button_create_pipe_tally_Click(object sender, EventArgs e)
        {
            string col1 = "MMID";
            string col2 = "Pipe";
            string col3 = "Heat";
            string col4 = "OriginalLength";
            string col5 = "NewLength";
            string col6 = "WallThickness";
            string col7 = "Diameter";
            string col8 = "Grade";
            string col9 = "Coating";
            string col10 = "Manufacture";
            string col11 = "DoubleJointNo";
            Wgen_main_form.dt_ground_tally = Functions.Creaza_weldmap_pipe_tally_datatable_structure();


            string c1 = textBoxf1.Text;
            string c2 = textBoxf2.Text;
            string c3 = textBoxf3.Text;
            string c4 = textBoxf4.Text;
            string c5 = textBoxf5.Text;
            string c6 = textBoxf6.Text;
            string c7 = textBoxf7.Text;
            string c8 = textBoxf8.Text;
            string c9 = textBoxf9.Text;
            string c10 = textBoxf10.Text;
            string c11 = textBoxf11.Text;


            if (comboBox_ws1.Text != "")
            {
                string string1 = comboBox_ws1.Text;
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
                            Wgen_main_form.dt_ground_tally = Functions.build_data_table_from_excel_based_on_11_columns_for_pipe_tally(Wgen_main_form.dt_ground_tally, W1, start_row,
                                                                col1, c1, col2, c2, col3, c3, col4, c4, col5, c5, col6, c6, col7, c7,
                                                                col8, c8, col9, c9, col10, c10, col11, c11);

                            for (int i = 0; i < Wgen_main_form.dt_ground_tally.Rows.Count; ++i)
                            {
                                if (Wgen_main_form.dt_ground_tally.Rows[i]["WallThickness"] != DBNull.Value)
                                {
                                    string wt = Convert.ToString(Wgen_main_form.dt_ground_tally.Rows[i]["WallThickness"]);
                                    if (Functions.IsNumeric(wt) == false)
                                    {
                                        if (wt.Contains(" WT") == true && wt.Contains("x") == true)
                                        {
                                            int index1 = wt.IndexOf("x");
                                            int index2 = wt.IndexOf(" WT");
                                            if (index1 < index2)
                                            {
                                                string new_wt = wt.Substring(index1 + 2, index2 - index1 - 2);
                                                Wgen_main_form.dt_ground_tally.Rows[i]["WallThickness"] = new_wt;
                                            }

                                        }
                                    }
                                }

                                if (Wgen_main_form.dt_ground_tally.Rows[i]["Diameter"] != DBNull.Value)
                                {
                                    string dia = Convert.ToString(Wgen_main_form.dt_ground_tally.Rows[i]["Diameter"]);
                                    if (Functions.IsNumeric(dia) == false)
                                    {
                                        if (dia.Contains(" OD") == true)
                                        {
                                            int index1 = dia.IndexOf(" OD");

                                            string new_diam = dia.Substring(0, index1);
                                            Wgen_main_form.dt_ground_tally.Rows[i]["Diameter"] = new_diam;

                                        }
                                    }
                                }

                                if (Wgen_main_form.dt_ground_tally.Rows[i]["Grade"] != DBNull.Value)
                                {
                                    string grade = Convert.ToString(Wgen_main_form.dt_ground_tally.Rows[i]["Grade"]);

                                    if (grade.Contains(" WT x ") == true)
                                    {
                                        int index1 = grade.IndexOf(" WT x ");
                                        int index2 = grade.IndexOf(" x ", index1);
                                        int index3 = grade.IndexOf(" ", index2 + 3);
                                        if (index2 < index3)
                                        {
                                            string new_grade = grade.Substring(index2 + 3, index3 - index2 - 3);
                                            Wgen_main_form.dt_ground_tally.Rows[i]["Grade"] = new_grade;
                                        }

                                    }

                                }

                            }

                        }
                    }
                }
            }
            Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_for_ground_tally(Wgen_main_form.dt_ground_tally);
            set_enable_true();
            button_pipe_tally_l.Visible = true;
            button_pipe_tally_nl.Visible = false;
        }

        private void button_refresh_ws1_Click(object sender, EventArgs e)
        {
            Functions.Load_opened_worksheets_to_combobox(comboBox_ws1);
            if (comboBox_ws1.Items.Count > 0)
            {
                for (int i = 0; i < comboBox_ws1.Items.Count; ++i)
                {
                    if (comboBox_ws1.Items[i].ToString().ToUpper().Contains("GROUND_TALLY") == true)
                    {
                        comboBox_ws1.SelectedIndex = i;
                        i = comboBox_ws1.Items.Count;
                    }
                }
            }
        }





    }
}
