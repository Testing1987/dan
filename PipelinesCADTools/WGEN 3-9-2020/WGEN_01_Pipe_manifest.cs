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
    public partial class Wgen_pipemanifest : Form
    {

        private ContextMenuStrip ContextMenuStrip_go_to_error;

        System.Data.DataTable dt_errors;
        System.Data.DataTable dt_export;
        int start_row = 2;

        public System.Data.DataTable dt_dismissed_errors = null;
        public string dismiss_errors_tab = "Dsd errors pm";
        Microsoft.Office.Interop.Excel.Worksheet W2 = null;
        public string filename = "";


        private void set_enable_false(object sender)
        {
            List<System.Windows.Forms.Button> lista_butoane = new List<Button>();

            lista_butoane.Add(button_load_pipe_manifest);
            lista_butoane.Add(button_pipe_manifest_l);
            lista_butoane.Add(button_pipe_manifest_nl);
            lista_butoane.Add(button_refresh_ws1);
            lista_butoane.Add(button_export_errors_to_xl);

            foreach (System.Windows.Forms.Button bt1 in lista_butoane)
            {
                if (sender as System.Windows.Forms.Button != bt1)
                {
                    bt1.Enabled = false;
                }
            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Button> lista_butoane = new List<Button>();

            lista_butoane.Add(button_load_pipe_manifest);
            lista_butoane.Add(button_pipe_manifest_l);
            lista_butoane.Add(button_pipe_manifest_nl);
            lista_butoane.Add(button_refresh_ws1);
            lista_butoane.Add(button_export_errors_to_xl);


            foreach (System.Windows.Forms.Button bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }

        public Wgen_pipemanifest()
        {
            InitializeComponent();
            var toolStripMenuItem2 = new ToolStripMenuItem { Text = "Go to error" };
            toolStripMenuItem2.Click += go_to_excel_point;


            ContextMenuStrip_go_to_error = new ContextMenuStrip();
            ContextMenuStrip_go_to_error.Items.AddRange(new ToolStripItem[] { toolStripMenuItem2 });
        }

        private void button_load_pipe_manifest_Click(object sender, EventArgs e)
        {

            string col1 = "Pipe ID";
            string col2 = "Heat";
            string col3 = "Length";
            string col4 = "Wall Thickness";
            string col5 = "Diameter";
            string col6 = "Grade";
            string col7 = "Coating";
            string col8 = "Manufacture";
            string col9 = "DoubleJointNo";
            Wgen_main_form.dt_double_joint = null;
            Wgen_main_form.dt_pipe_list = Functions.Creaza_weldmap_pipelist_datatable_structure();
            make_first_line_invisible();
            if (comboBox_ws1.Text != "")
            {
                string string1 = comboBox_ws1.Text;
                if (string1.Contains("[") == true && string1.Contains("]") == true)
                {
                     filename = string1.Substring(string1.IndexOf("]") + 4, string1.Length - (string1.IndexOf("]") + 4));
                    string sheet_name = string1.Substring(1, string1.IndexOf("]") - 1);
                    if (filename.Length > 0 && sheet_name.Length > 0)
                    {
                        set_enable_false(sender);
                        Microsoft.Office.Interop.Excel.Worksheet W1 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, sheet_name);
                        W2 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, dismiss_errors_tab);

                        if (W1 != null)
                        {

                            Wgen_main_form.dt_ground_tally = null;
                            Wgen_main_form.dt_all_points = null;
                            Wgen_main_form.dt_weld_map = null;
                            Wgen_main_form.dt_pt_keep = null;
                            Wgen_main_form.dt_pt_move = null;

                            Wgen_main_form.dt_pipe_list = Functions.Populate_data_table_from_excel(Wgen_main_form.dt_pipe_list, W1, start_row, textBox_1.Text, textBox_2.Text, textBox_3.Text, textBox_4.Text, textBox_5.Text, textBox_6.Text, textBox_7.Text, textBox_8.Text, textBox_9.Text, "", "",
                                "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", true);

                            if (W2 != null)
                            {
                                dt_dismissed_errors = new System.Data.DataTable();
                                dt_dismissed_errors.Columns.Add("Point", typeof(string));
                                dt_dismissed_errors.Columns.Add("Value", typeof(string));
                                dt_dismissed_errors.Columns.Add("Error", typeof(string));

                                dt_dismissed_errors = Functions.Populate_data_table_from_excel(dt_dismissed_errors, W2, start_row,
                                    "A", "B", "C", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", false);
                                if (dt_dismissed_errors.Rows.Count == 0) dt_dismissed_errors = null;
                            }
                            else
                            {
                                dt_dismissed_errors = null;
                            }


                            if (Wgen_main_form.dt_pipe_list.Rows.Count > 0)
                            {
                                for (int i = 0; i < Wgen_main_form.dt_pipe_list.Rows.Count; ++i)
                                {
                                    if (Wgen_main_form.dt_pipe_list.Rows[i]["DoubleJointNo"] != DBNull.Value)
                                    {
                                        string dj1 = Convert.ToString(Wgen_main_form.dt_pipe_list.Rows[i]["DoubleJointNo"]);
                                        if (dj1.ToLower() == "0" || dj1.ToLower() == "na" || dj1.ToLower() == @"n/a" || dj1.ToLower() == @"n\a")
                                        {
                                            Wgen_main_form.dt_pipe_list.Rows[i]["DoubleJointNo"] = DBNull.Value;
                                        }

                                    }
                                }


                                int nr_duplicates_dj = 0;
                                int nr_duplicates_pipe_heat = 0;
                                int nr_null_values = 0;
                                dt_errors = new System.Data.DataTable();
                                dt_errors.Columns.Add("Value1", typeof(string));
                                dt_errors.Columns.Add("Value2", typeof(string));
                                dt_errors.Columns.Add("Excel", typeof(string));
                                dt_errors.Columns.Add("w1", typeof(Microsoft.Office.Interop.Excel.Worksheet));
                                dt_errors.Columns.Add("Error", typeof(string));

                                for (int i = 0; i < Wgen_main_form.dt_pipe_list.Rows.Count; ++i)
                                {
                                    if (Wgen_main_form.dt_pipe_list.Rows[i][col4] != DBNull.Value)
                                    {
                                        string wallT = Convert.ToString(Wgen_main_form.dt_pipe_list.Rows[i][col4]).Replace(" ", "");


                                        if (Functions.IsNumeric(wallT) == true)
                                        {
                                            double Wall1 = Convert.ToDouble(wallT);
                                            Wgen_main_form.dt_pipe_list.Rows[i][col4] = Convert.ToString(Wall1);
                                        }

                                    }

                                    if (Wgen_main_form.dt_pipe_list.Rows[i][col6] != DBNull.Value)
                                    {

                                        string grade = Convert.ToString(Wgen_main_form.dt_pipe_list.Rows[i][col6]).Replace(" ", "");
                                        Wgen_main_form.dt_pipe_list.Rows[i][col6] = grade;

                                    }
                                }



                                var duplicates1 = Wgen_main_form.dt_pipe_list.AsEnumerable().GroupBy(i => new { Pipeid = i.Field<string>(col1), Heat = i.Field<string>(col2) }).Where(g => g.Count() > 1).Select(g => new { g.Key.Pipeid, g.Key.Heat }).ToList();
                                var duplicates2 = Wgen_main_form.dt_pipe_list.AsEnumerable().GroupBy(i => new { dbljoint = i.Field<string>(col9) }).Where(g => g.Count() > 1).Select(g => new { g.Key.dbljoint }).ToList();

                                System.Data.DataTable dt2 = new System.Data.DataTable();
                                dt2.Columns.Add(col1, typeof(string));
                                dt2.Columns.Add(col2, typeof(string));
                                dt2.Columns.Add(col9, typeof(string));
                                if (duplicates1.Count > 0)
                                {
                                    for (int i = 0; i < duplicates1.Count; ++i)
                                    {
                                        if (duplicates1[i].Pipeid != null && duplicates1[i].Heat != null)
                                        {
                                            string duplicat_val1 = Convert.ToString(duplicates1[i].Pipeid);
                                            string duplicat_val2 = Convert.ToString(duplicates1[i].Heat);
                                            dt2.Rows.Add();
                                            dt2.Rows[dt2.Rows.Count - 1][0] = duplicat_val1;
                                            dt2.Rows[dt2.Rows.Count - 1][1] = duplicat_val2;
                                        }
                                    }
                                }

                                if (duplicates2.Count > 0)
                                {
                                    for (int i = 0; i < duplicates2.Count; ++i)
                                    {
                                        if (duplicates2[i].dbljoint != null)
                                        {
                                            string duplicat_val1 = Convert.ToString(duplicates2[i].dbljoint);
                                            dt2.Rows.Add();
                                            dt2.Rows[dt2.Rows.Count - 1][2] = duplicat_val1;

                                        }
                                    }
                                }
                                DataSet dataset1 = new DataSet();
                                dataset1.Tables.Add(Wgen_main_form.dt_pipe_list);
                                dataset1.Tables.Add(dt2);



                                DataRelation relation1 = new DataRelation("xxx", Wgen_main_form.dt_pipe_list.Columns[col1], dt2.Columns[col1], false);
                                DataRelation relation2 = new DataRelation("xxx1", Wgen_main_form.dt_pipe_list.Columns[col2], dt2.Columns[col2], false);
                                DataRelation relation3 = new DataRelation("xxx2", Wgen_main_form.dt_pipe_list.Columns[col9], dt2.Columns[col9], false);

                                dataset1.Relations.Add(relation1);
                                dataset1.Relations.Add(relation2);
                                dataset1.Relations.Add(relation3);



                                if (dt2.Rows.Count > 0)
                                {

                                    List<string> lista_dj1 = new List<string>();
                                    List<string> lista_dj2 = new List<string>();
                                    List<string> lista_dj3 = new List<string>();
                                    for (int i = 0; i < Wgen_main_form.dt_pipe_list.Rows.Count; ++i)
                                    {
                                        if (Wgen_main_form.dt_pipe_list.Rows[i].GetChildRows(relation1).Length > 0 && Wgen_main_form.dt_pipe_list.Rows[i].GetChildRows(relation2).Length > 0)
                                        {
                                            for (int j = 0; j < Wgen_main_form.dt_pipe_list.Rows[i].GetChildRows(relation1).Length; ++j)
                                            {
                                                string val1 = Wgen_main_form.dt_pipe_list.Rows[i].GetChildRows(relation1)[j][col1].ToString();
                                                string val2 = Wgen_main_form.dt_pipe_list.Rows[i].GetChildRows(relation1)[j][col2].ToString();

                                                dt_errors.Rows.Add();
                                                dt_errors.Rows[dt_errors.Rows.Count - 1][0] = "PipeID:" + val1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1][1] = "Heat:" + val2;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1][2] = textBox_1.Text + Convert.ToString(i + start_row);
                                                dt_errors.Rows[dt_errors.Rows.Count - 1][3] = W1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1][4] = "Duplicate Pipe ID & Heat";
                                                ++nr_duplicates_pipe_heat;

                                            }
                                        }

                                        if (Wgen_main_form.dt_pipe_list.Rows[i].GetChildRows(relation3).Length > 0)
                                        {
                                            for (int j = 0; j < Wgen_main_form.dt_pipe_list.Rows[i].GetChildRows(relation3).Length; ++j)
                                            {
                                                string val1 = Wgen_main_form.dt_pipe_list.Rows[i].GetChildRows(relation3)[j][col9].ToString();
                                                if (lista_dj1.Contains(val1) == false)
                                                {
                                                    lista_dj1.Add(val1);
                                                }
                                                else
                                                {
                                                    if (val1.ToUpper().Contains("DJ") == true)
                                                    {
                                                        if (lista_dj2.Contains(val1) == false)
                                                        {
                                                            lista_dj2.Add(val1);
                                                        }
                                                        else
                                                        {
                                                            dt_errors.Rows.Add();
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1][0] = "Double Joint: " + val1;
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1][2] = textBox_9.Text + Convert.ToString(i + start_row);
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1][3] = W1;
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1][4] = "Double Joint listed more than twice";
                                                            ++nr_duplicates_dj;
                                                        }
                                                    }
                                                    else if (val1.ToUpper().Contains("TJ") == true)
                                                    {
                                                        if (lista_dj2.Contains(val1) == false)
                                                        {
                                                            lista_dj2.Add(val1);
                                                        }
                                                        else
                                                        {
                                                            if (lista_dj3.Contains(val1) == false)
                                                            {
                                                                lista_dj3.Add(val1);
                                                            }
                                                            else
                                                            {
                                                                dt_errors.Rows.Add();
                                                                dt_errors.Rows[dt_errors.Rows.Count - 1][0] = "Double Joint: " + val1;
                                                                dt_errors.Rows[dt_errors.Rows.Count - 1][2] = textBox_9.Text + Convert.ToString(i + start_row);
                                                                dt_errors.Rows[dt_errors.Rows.Count - 1][3] = W1;
                                                                dt_errors.Rows[dt_errors.Rows.Count - 1][4] = "Triple Joint listed more than 3 times";
                                                                ++nr_duplicates_dj;
                                                            }
                                                        }
                                                    }


                                                }

                                            }
                                        }
                                    }
                                }

                                dataset1.Relations.Remove(relation1);
                                dataset1.Relations.Remove(relation2);
                                dataset1.Relations.Remove(relation3);
                                dataset1.Tables.Remove(Wgen_main_form.dt_pipe_list);
                                dataset1.Tables.Remove(dt2);
                                dt2 = null;




                                for (int i = 0; i < Wgen_main_form.dt_pipe_list.Rows.Count; ++i)
                                {
                                    string pipeID = "";
                                    string heat = "";

                                    if (Wgen_main_form.dt_pipe_list.Rows[i][col1] == DBNull.Value)
                                    {
                                        dt_errors.Rows.Add();
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][2] = textBox_1.Text + Convert.ToString(i + start_row);
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][3] = W1;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][4] = "No Pipe ID Value Specified";
                                        ++nr_null_values;
                                    }
                                    else
                                    {
                                        pipeID = Convert.ToString(Wgen_main_form.dt_pipe_list.Rows[i][col1]);
                                    }

                                    if (Wgen_main_form.dt_pipe_list.Rows[i][col2] == DBNull.Value)
                                    {
                                        dt_errors.Rows.Add();
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][0] = "PipeID:" + pipeID;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][2] = textBox_2.Text + Convert.ToString(i + start_row);
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][3] = W1;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][4] = "No Heat Value Specified";
                                        ++nr_null_values;
                                    }
                                    else
                                    {
                                        heat = Convert.ToString(Wgen_main_form.dt_pipe_list.Rows[i][col2]);
                                    }

                                    if (Wgen_main_form.dt_pipe_list.Rows[i][col3] == DBNull.Value || Functions.IsNumeric(Convert.ToString(Wgen_main_form.dt_pipe_list.Rows[i][col3])) == false)
                                    {
                                        dt_errors.Rows.Add();
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][0] = "PipeID:" + pipeID;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][1] = "Heat:" + pipeID;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][2] = textBox_3.Text + Convert.ToString(i + start_row);
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][3] = W1;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][4] = "No Length Value Specified";
                                        ++nr_null_values;
                                    }

                                    if (Wgen_main_form.dt_pipe_list.Rows[i][col4] == DBNull.Value || Functions.IsNumeric(Convert.ToString(Wgen_main_form.dt_pipe_list.Rows[i][col4])) == false)
                                    {
                                        dt_errors.Rows.Add();
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][0] = "PipeID:" + pipeID;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][1] = "Heat:" + pipeID;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][2] = textBox_4.Text + Convert.ToString(i + start_row);
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][3] = W1;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][4] = "No Wall Thickness Specified";
                                        ++nr_null_values;
                                    }

                                    if (Wgen_main_form.dt_pipe_list.Rows[i][col5] == DBNull.Value || Functions.IsNumeric(Convert.ToString(Wgen_main_form.dt_pipe_list.Rows[i][col5])) == false)
                                    {
                                        dt_errors.Rows.Add();
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][0] = "PipeID:" + pipeID;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][1] = "Heat:" + pipeID;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][2] = textBox_5.Text + Convert.ToString(i + start_row);
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][3] = W1;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][4] = "No Diameter Specified";
                                        ++nr_null_values;
                                    }

                                    if (Wgen_main_form.dt_pipe_list.Rows[i][col6] == DBNull.Value)
                                    {
                                        dt_errors.Rows.Add();
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][0] = "PipeID:" + pipeID;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][1] = "Heat:" + pipeID;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][2] = textBox_6.Text + Convert.ToString(i + start_row);
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][3] = W1;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][4] = "No Grade Specified";
                                        ++nr_null_values;
                                    }

                                    if (Wgen_main_form.dt_pipe_list.Rows[i][col7] == DBNull.Value)
                                    {
                                        dt_errors.Rows.Add();
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][0] = "PipeID:" + pipeID;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][1] = "Heat:" + pipeID;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][2] = textBox_7.Text + Convert.ToString(i + start_row);
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][3] = W1;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][4] = "No Coating Specified";
                                        ++nr_null_values;
                                    }

                                    if (Wgen_main_form.dt_pipe_list.Rows[i][col8] == DBNull.Value)
                                    {
                                        dt_errors.Rows.Add();
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][0] = "PipeID:" + pipeID;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][1] = "Heat:" + pipeID;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][2] = textBox_8.Text + Convert.ToString(i + start_row);
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][3] = W1;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][4] = "No Manufacture Specified";
                                        ++nr_null_values;
                                    }
                                }

                                dt_errors = Functions.Sort_data_table(dt_errors, "Error");
                                transfer_errors_to_panel(dt_errors);
                                dt_export = Functions.creaza_error_export_table_for_pipe_manifest(dt_errors, sheet_name);

                                textBox_PM_no_rows.Text = Convert.ToString(Wgen_main_form.dt_pipe_list.Rows.Count);
                                textBox_PM_no_duplicates.Text = Convert.ToString(nr_duplicates_pipe_heat);
                                textBox_PM_no_dj_duplicates.Text = Convert.ToString(nr_duplicates_dj);
                                textBox_PM_no_null.Text = Convert.ToString(nr_null_values);
                                button_pipe_manifest_l.Visible = true;
                                button_pipe_manifest_nl.Visible = false;
                            }
                            else
                            {
                                button_pipe_manifest_l.Visible = false;
                                button_pipe_manifest_nl.Visible = true;
                            }
                        }
                        set_enable_true();
                    }
                    else
                    {
                        button_pipe_manifest_l.Visible = false;
                        button_pipe_manifest_nl.Visible = true;
                    }
                }
                else
                {
                    button_pipe_manifest_l.Visible = false;
                    button_pipe_manifest_nl.Visible = true;
                }
            }
            else
            {
                button_pipe_manifest_l.Visible = false;
                button_pipe_manifest_nl.Visible = true;
            }


            #region double joint

            List<string> lista1 = new List<string>();

            if (Wgen_main_form.dt_pipe_list != null && Wgen_main_form.dt_pipe_list.Rows.Count > 0)
            {
                for (int i = 0; i < Wgen_main_form.dt_pipe_list.Rows.Count; ++i)
                {
                    if (Wgen_main_form.dt_pipe_list.Rows[i][col9] != DBNull.Value & Wgen_main_form.dt_pipe_list.Rows[i][col2] != DBNull.Value &&
                        Wgen_main_form.dt_pipe_list.Rows[i][col3] != DBNull.Value && Wgen_main_form.dt_pipe_list.Rows[i][col4] != DBNull.Value &&
                        Wgen_main_form.dt_pipe_list.Rows[i][col5] != DBNull.Value && Wgen_main_form.dt_pipe_list.Rows[i][col6] != DBNull.Value &&
                        Wgen_main_form.dt_pipe_list.Rows[i][col7] != DBNull.Value && Wgen_main_form.dt_pipe_list.Rows[i][col8] != DBNull.Value &&
                        Wgen_main_form.dt_pipe_list.Rows[i][col1] != DBNull.Value)
                    {

                        string pipeid1 = Convert.ToString(Wgen_main_form.dt_pipe_list.Rows[i][col1]);
                        string Heat1 = Convert.ToString(Wgen_main_form.dt_pipe_list.Rows[i][col2]);
                        double len1 = Convert.ToDouble(Wgen_main_form.dt_pipe_list.Rows[i][col3]);
                        string wt1 = Convert.ToString(Wgen_main_form.dt_pipe_list.Rows[i][col4]);
                        string od1 = Convert.ToString(Wgen_main_form.dt_pipe_list.Rows[i][col5]);
                        string grade1 = Convert.ToString(Wgen_main_form.dt_pipe_list.Rows[i][col6]);
                        string coat1 = Convert.ToString(Wgen_main_form.dt_pipe_list.Rows[i][col7]);
                        string manuf1 = Convert.ToString(Wgen_main_form.dt_pipe_list.Rows[i][col8]);
                        string dj1 = Convert.ToString(Wgen_main_form.dt_pipe_list.Rows[i][col9]);

                        if (dj1.Replace(" ", "").Length > 0)
                        {
                            List<double> lista_len = new List<double>();
                            lista_len.Add(len1);
                            if (Wgen_main_form.dt_double_joint == null)
                            {
                                Wgen_main_form.dt_double_joint = new System.Data.DataTable();

                                Wgen_main_form.dt_double_joint.Columns.Add("pipeid", typeof(string));
                                Wgen_main_form.dt_double_joint.Columns.Add("heat", typeof(string));
                                Wgen_main_form.dt_double_joint.Columns.Add("total_len", typeof(double));
                                Wgen_main_form.dt_double_joint.Columns.Add("wt", typeof(string));
                                Wgen_main_form.dt_double_joint.Columns.Add("od", typeof(string));
                                Wgen_main_form.dt_double_joint.Columns.Add("grade", typeof(string));
                                Wgen_main_form.dt_double_joint.Columns.Add("coating", typeof(string));
                                Wgen_main_form.dt_double_joint.Columns.Add("manufacturer", typeof(string));
                                Wgen_main_form.dt_double_joint.Columns.Add("double_joint", typeof(string));

                            }

                            if (lista1.Contains(dj1) == false)
                            {
                                Wgen_main_form.dt_double_joint.Rows.Add();
                                Wgen_main_form.dt_double_joint.Rows[Wgen_main_form.dt_double_joint.Rows.Count - 1]["pipeid"] = pipeid1;
                                Wgen_main_form.dt_double_joint.Rows[Wgen_main_form.dt_double_joint.Rows.Count - 1]["heat"] = Heat1;
                                Wgen_main_form.dt_double_joint.Rows[Wgen_main_form.dt_double_joint.Rows.Count - 1]["total_len"] = len1;
                                Wgen_main_form.dt_double_joint.Rows[Wgen_main_form.dt_double_joint.Rows.Count - 1]["wt"] = wt1;
                                Wgen_main_form.dt_double_joint.Rows[Wgen_main_form.dt_double_joint.Rows.Count - 1]["od"] = od1;
                                Wgen_main_form.dt_double_joint.Rows[Wgen_main_form.dt_double_joint.Rows.Count - 1]["grade"] = grade1;
                                Wgen_main_form.dt_double_joint.Rows[Wgen_main_form.dt_double_joint.Rows.Count - 1]["coating"] = coat1;
                                Wgen_main_form.dt_double_joint.Rows[Wgen_main_form.dt_double_joint.Rows.Count - 1]["manufacturer"] = manuf1;
                                Wgen_main_form.dt_double_joint.Rows[Wgen_main_form.dt_double_joint.Rows.Count - 1]["double_joint"] = dj1;
                                lista1.Add(dj1);
                            }
                            else
                            {
                                for (int j = 0; j < Wgen_main_form.dt_double_joint.Rows.Count; ++j)
                                {
                                    string dj2 = Convert.ToString(Wgen_main_form.dt_double_joint.Rows[j]["double_joint"]);
                                    if (dj1.ToLower() == dj2.ToLower() && i != j)
                                    {
                                        string heat2 = Convert.ToString(Wgen_main_form.dt_double_joint.Rows[j]["heat"]);
                                        string pipeid2 = Convert.ToString(Wgen_main_form.dt_double_joint.Rows[j]["pipeid"]);
                                        double len2 = Convert.ToDouble(Wgen_main_form.dt_double_joint.Rows[j]["total_len"]);
                                        lista_len.Add(len2);
                                        List<string> lp2 = new List<string>();
                                        if (pipeid2.Contains("/") == true)
                                        {
                                            string[] h1 = pipeid2.Split(Convert.ToChar("/"));
                                            for (int k = 0; k < h1.Length; ++k)
                                            {
                                                lp2.Add(h1[k]);
                                            }
                                        }
                                        else
                                        {
                                            lp2.Add(pipeid2);
                                        }

                                        List<string> lh2 = new List<string>();
                                        if (heat2.Contains("/") == true)
                                        {
                                            string[] h1 = heat2.Split(Convert.ToChar("/"));
                                            for (int k = 0; k < h1.Length; ++k)
                                            {
                                                lh2.Add(h1[k]);
                                            }
                                        }
                                        else
                                        {
                                            lh2.Add(heat2);
                                        }

                                        if (lp2.Contains(pipeid1) == false)
                                        {
                                            Wgen_main_form.dt_double_joint.Rows[j]["pipeid"] = pipeid2 + "/" + pipeid1;
                                        }

                                        if (lh2.Contains(Heat1) == false)
                                        {
                                            Wgen_main_form.dt_double_joint.Rows[j]["heat"] = heat2 + "/" + Heat1;
                                        }

                                        double tot_len = 0;

                                        for (int k = 0; k < lista_len.Count; ++k)
                                        {
                                            tot_len = tot_len + lista_len[k];
                                        }


                                        Wgen_main_form.dt_double_joint.Rows[j]["total_len"] = tot_len;

                                    }
                                }
                            }
                        }




                    }

                }
            }
            #endregion


            //Functions.Transfer_datatable_to_new_excel_spreadsheet(dt_pipe_manifest);
        }






        private void make_first_line_invisible()
        {
            dataGridView_error_pipe_manifest.DataSource = null;
        }


        private void button_refresh_ws1_Click(object sender, EventArgs e)
        {
            Functions.Load_opened_worksheets_to_combobox(comboBox_ws1);
            if (comboBox_ws1.Items.Count > 0)
            {
                for (int i = 0; i < comboBox_ws1.Items.Count; ++i)
                {
                    if (comboBox_ws1.Items[i].ToString().ToUpper().Contains("PIPE_LIST") == true)
                    {
                        comboBox_ws1.SelectedIndex = i;
                        i = comboBox_ws1.Items.Count;
                    }
                }
            }

        }



        private void checkBox_incomplete_pipe_manifest_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_incomplete_pipe_manifest.Checked == true)
            {
                Wgen_main_form.incomplete_pipe_manifest = true;
            }
            else
            {
                Wgen_main_form.incomplete_pipe_manifest = false;
            }

        }


        private void DataGridView_error_pipe_manifest_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right && e.RowIndex >= 0)
            {
                dataGridView_error_pipe_manifest.CurrentCell = dataGridView_error_pipe_manifest.Rows[e.RowIndex].Cells[0];
                ContextMenuStrip_go_to_error.Show(Cursor.Position);
                ContextMenuStrip_go_to_error.Visible = true;
            }
            else
            {
                ContextMenuStrip_go_to_error.Visible = false;
            }
        }

        private void go_to_excel_point(object sender, EventArgs e)
        {
            if (dt_errors == null || dt_errors.Rows.Count == 0) return;

            int index1 = dataGridView_error_pipe_manifest.CurrentCell.RowIndex;
            try
            {
                if (dt_errors.Rows.Count - 1 >= index1)
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
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }



        }

        public void radioButton_pipe_manifest_CheckedChanged(RadioButton radioButton_enlarged)
        {

            Font regularfont = new Font("Arial", 8.2f, FontStyle.Bold);


            Font englargedFont = new Font("Arial", 10f, FontStyle.Bold);


            Font regularHeader = new Font("Arial", 10f, FontStyle.Bold);

            Font englargedHeader = new Font("Arial", 12f, FontStyle.Bold);
            if (radioButton_enlarged.Checked == true)
            {
                panel7.Location = new Point(3, 3);
                panel7.Size = new Size(723, 30);
            }
            else
            {
                panel7.Location = new Point(3, 3);
                panel7.Size = new Size(723, 25);
            }

            if (radioButton_enlarged.Checked == true)
            {
                label12.Location = new Point(5, 5);
                label12.Size = new Size(93, 23);

                label12.Font = englargedHeader;
            }
            else
            {
                label12.Location = new Point(3, 3);
                label12.Size = new Size(75, 18);

                label12.Font = regularHeader;
            }

            if (radioButton_enlarged.Checked == true)
            {
                panel_pipe_manifest.Location = new Point(3, 35);
                panel_pipe_manifest.Size = new Size(723, 35);
            }
            else
            {
                panel_pipe_manifest.Location = new Point(3, 29);
                panel_pipe_manifest.Size = new Size(723, 33);
            }

            if (radioButton_enlarged.Checked == true)
            {
                button_refresh_ws1.Location = new Point(4, 4);
                button_refresh_ws1.Size = new Size(24, 24);
            }
            else
            {
                button_refresh_ws1.Location = new Point(5, 5);
                button_refresh_ws1.Size = new Size(21, 21);
            }

            if (radioButton_enlarged.Checked == true)
            {
                comboBox_ws1.Location = new Point(32, 4);
                comboBox_ws1.Size = new Size(455, 28);
            }
            else
            {
                comboBox_ws1.Location = new Point(35, 4);
                comboBox_ws1.Size = new Size(510, 24);
            }

            if (radioButton_enlarged.Checked == true)
            {
                button_load_pipe_manifest.Location = new Point(494, 2);
                button_load_pipe_manifest.Size = new Size(198, 30);

                button_load_pipe_manifest.Font = englargedFont;
            }
            else
            {
                button_load_pipe_manifest.Location = new Point(546, 2);
                button_load_pipe_manifest.Size = new Size(145, 28);

                button_load_pipe_manifest.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                button_pipe_manifest_l.Location = new Point(695, 3);
                button_pipe_manifest_l.Size = new Size(24, 24);
            }
            else
            {
                button_pipe_manifest_l.Location = new Point(697, 5);
                button_pipe_manifest_l.Size = new Size(21, 21);
            }

            if (radioButton_enlarged.Checked == true)
            {
                button_pipe_manifest_nl.Location = new Point(695, 3);
                button_pipe_manifest_nl.Size = new Size(24, 24);
            }
            else
            {
                button_pipe_manifest_nl.Location = new Point(697, 5);
                button_pipe_manifest_nl.Size = new Size(21, 21);
            }

            if (radioButton_enlarged.Checked == true)
            {
                panel6.Location = new Point(2, 72);
                panel6.Size = new Size(723, 79);
            }
            else
            {
                panel6.Location = new Point(2, 64);
                panel6.Size = new Size(723, 79);
            }

            if (radioButton_enlarged.Checked == true)
            {
                label18.Location = new Point(3, 3);
                label18.Size = new Size(81, 30);

                label18.Font = englargedHeader;
            }
            else
            {
                label18.Location = new Point(3, 3);
                label18.Size = new Size(81, 25);

                label18.Font = regularHeader;
            }

            if (radioButton_enlarged.Checked == true)
            {
                checkBox_incomplete_pipe_manifest.Location = new Point(525, 3);
                checkBox_incomplete_pipe_manifest.Size = new Size(81, 30);

                checkBox_incomplete_pipe_manifest.Font = englargedFont;
            }
            else
            {
                checkBox_incomplete_pipe_manifest.Location = new Point(546, 3);
                checkBox_incomplete_pipe_manifest.Size = new Size(81, 25);

                checkBox_incomplete_pipe_manifest.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                dataGridView_error_pipe_manifest.Location = new Point(2, 100);
                dataGridView_error_pipe_manifest.Size = new Size(723, 445);

                dataGridView_error_pipe_manifest.DefaultCellStyle.Font = englargedFont;

                dataGridView_error_pipe_manifest.ColumnHeadersDefaultCellStyle.Font = englargedHeader;

            }
            else
            {
                dataGridView_error_pipe_manifest.Location = new Point(2, 91);
                dataGridView_error_pipe_manifest.Size = new Size(723, 468);

                dataGridView_error_pipe_manifest.DefaultCellStyle.Font = regularfont;
                dataGridView_error_pipe_manifest.ColumnHeadersDefaultCellStyle.Font = regularHeader;

            }

            if (radioButton_enlarged.Checked == true)
            {
                panel_stats.Location = new Point(3, 548);
                panel_stats.Size = new Size(722, 125);
            }
            else
            {
                panel_stats.Location = new Point(3, 560);
                panel_stats.Size = new Size(722, 108);
            }

            if (radioButton_enlarged.Checked == true)
            {
                label19.Location = new Point(5, 3);
                label19.Size = new Size(58, 14);

                label19.Font = englargedFont;
            }
            else
            {
                label19.Location = new Point(5, 3);
                label19.Size = new Size(58, 14);

                label19.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_PM_Items.Location = new Point(3, 21);
                textBox_PM_Items.Size = new Size(310, 25);

                textBox_PM_Items.Font = englargedFont;
            }
            else
            {
                textBox_PM_Items.Location = new Point(3, 21);
                textBox_PM_Items.Size = new Size(287, 20);

                textBox_PM_Items.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_PM_no_rows.Location = new Point(325, 21);
                textBox_PM_no_rows.Size = new Size(40, 25);

                textBox_PM_no_rows.Font = englargedFont;
            }
            else
            {
                textBox_PM_no_rows.Location = new Point(311, 21);
                textBox_PM_no_rows.Size = new Size(37, 20);

                textBox_PM_no_rows.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_PM_missing_OD.Location = new Point(3, 47);
                textBox_PM_missing_OD.Size = new Size(310, 25);

                textBox_PM_missing_OD.Font = englargedFont;
            }
            else
            {
                textBox_PM_missing_OD.Location = new Point(3, 42);
                textBox_PM_missing_OD.Size = new Size(287, 20);

                textBox_PM_missing_OD.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_PM_no_duplicates.Location = new Point(325, 47);
                textBox_PM_no_duplicates.Size = new Size(40, 25);

                textBox_PM_no_duplicates.Font = englargedFont;
            }
            else
            {
                textBox_PM_no_duplicates.Location = new Point(311, 42);
                textBox_PM_no_duplicates.Size = new Size(37, 20);

                textBox_PM_no_duplicates.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_PM_double_jpint_duplicates.Location = new Point(3, 73);
                textBox_PM_double_jpint_duplicates.Size = new Size(310, 25);

                textBox_PM_double_jpint_duplicates.Font = englargedFont;
            }
            else
            {
                textBox_PM_double_jpint_duplicates.Location = new Point(3, 63);
                textBox_PM_double_jpint_duplicates.Size = new Size(287, 20);

                textBox_PM_double_jpint_duplicates.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_PM_no_dj_duplicates.Location = new Point(325, 73);
                textBox_PM_no_dj_duplicates.Size = new Size(40, 25);

                textBox_PM_no_dj_duplicates.Font = englargedFont;
            }
            else
            {
                textBox_PM_no_dj_duplicates.Location = new Point(311, 63);
                textBox_PM_no_dj_duplicates.Size = new Size(37, 20);

                textBox_PM_no_dj_duplicates.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_PM_null_value_items.Location = new Point(3, 99);
                textBox_PM_null_value_items.Size = new Size(310, 25);

                textBox_PM_null_value_items.Font = englargedFont;
            }
            else
            {
                textBox_PM_null_value_items.Location = new Point(3, 84);
                textBox_PM_null_value_items.Size = new Size(287, 20);

                textBox_PM_null_value_items.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_PM_no_null.Location = new Point(325, 99);
                textBox_PM_no_null.Size = new Size(40, 25);

                textBox_PM_no_null.Font = englargedFont;
            }
            else
            {
                textBox_PM_no_null.Location = new Point(311, 84);
                textBox_PM_no_null.Size = new Size(37, 20);

                textBox_PM_no_null.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                button_export_errors_to_xl.Location = new Point(505, 90);
                button_export_errors_to_xl.Size = new Size(214, 30);

                button_export_errors_to_xl.Font = englargedFont;
            }
            else
            {
                button_export_errors_to_xl.Location = new Point(556, 75);
                button_export_errors_to_xl.Size = new Size(161, 28);

                button_export_errors_to_xl.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                button_dismiss_errors.Location = new Point(385,-1);
                button_dismiss_errors.Size = new Size(124, 30);

                button_dismiss_errors.Font = englargedFont;
            }
            else
            {
                button_dismiss_errors.Location = new Point(436,-1);
                button_dismiss_errors.Size = new Size(104, 25);

                button_dismiss_errors.Font = regularfont;
            }

        }


        private void button_export_errors_to_xl_Click(object sender, EventArgs e)
        {
            Functions.Transfer_datatable_to_new_excel_spreadsheet_named(dt_export, "PipeManifestErrors");
        }


        private void transfer_errors_to_panel(System.Data.DataTable dt1)
        {
            if (dt1.Rows.Count > 0)
            {
                System.Data.DataTable dt_display = dt1.Copy();
                dt_display.Columns.RemoveAt(3);

                if (dt_dismissed_errors != null && dt_dismissed_errors.Rows.Count > 0)
                {
                    dt_display.TableName = "err0";
                    dt_dismissed_errors.TableName = "t8";
                    DataSet dataset1 = new DataSet();
                    dataset1.Tables.Add(dt_dismissed_errors);
                    dataset1.Tables.Add(dt_display);

                    DataRelation rel0 = new DataRelation("xxx", dt_display.Columns[0], dt_dismissed_errors.Columns[0], false);
                    DataRelation rel1 = new DataRelation("xxx1", dt_display.Columns[1], dt_dismissed_errors.Columns[1], false);
                    DataRelation rel2 = new DataRelation("xxx2", dt_display.Columns[3], dt_dismissed_errors.Columns[2], false);


                    dataset1.Relations.Add(rel0);
                    dataset1.Relations.Add(rel1);
                    dataset1.Relations.Add(rel2);

                    for (int i = dt_display.Rows.Count - 1; i >= 0; --i)
                    {
                        if (dt_display.Rows[i][0] != DBNull.Value && dt_display.Rows[i][1] != DBNull.Value && dt_display.Rows[i][3] != DBNull.Value)
                        {
                            if (dt_display.Rows[i].GetChildRows(rel0).Length > 0 && dt_display.Rows[i].GetChildRows(rel1).Length > 0 && dt_display.Rows[i].GetChildRows(rel2).Length > 0)
                            {
                                // the 3 rel are not the right way... they are taking out extra lines
                                dt_display.Rows[i].Delete();
                            }
                        }
                    }
                  


                    dataset1.Relations.Remove(rel0);
                    dataset1.Relations.Remove(rel1);
                    dataset1.Relations.Remove(rel2);
                    dataset1.Tables.Remove(dt_dismissed_errors);
                    dataset1.Tables.Remove(dt_display);


                }

                dataGridView_error_pipe_manifest.DataSource = dt_display;
                dataGridView_error_pipe_manifest.Columns[0].Width = 75;
                dataGridView_error_pipe_manifest.Columns[1].Width = 75;
                dataGridView_error_pipe_manifest.Columns[2].Width = 60;
                dataGridView_error_pipe_manifest.Columns[3].Width = 300;
                dataGridView_error_pipe_manifest.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_error_pipe_manifest.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                dataGridView_error_pipe_manifest.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_error_pipe_manifest.DefaultCellStyle.ForeColor = Color.White;
                dataGridView_error_pipe_manifest.EnableHeadersVisualStyles = false;
            }
        }


        private void button_dismiss_errors_Click(object sender, EventArgs e)
        {
            try
            {
                if (dt_dismissed_errors == null)
                {
                    dt_dismissed_errors = new System.Data.DataTable();
                    dt_dismissed_errors.Columns.Add("Point", typeof(string));
                    dt_dismissed_errors.Columns.Add("Value", typeof(string));
                    dt_dismissed_errors.Columns.Add("Error", typeof(string));
                }

                List<int> lista1 = new List<int>();

                foreach (DataGridViewCell cell1 in dataGridView_error_pipe_manifest.SelectedCells)
                {
                    int row_index = cell1.RowIndex;
                    if (lista1.Contains(row_index) == false)
                    {
                        lista1.Add(row_index);
                    }
                }


                if (lista1.Count > 0)
                {
                    for (int i = 0; i < lista1.Count; ++i)
                    {
                        dt_dismissed_errors.Rows.Add();
                        dt_dismissed_errors.Rows[dt_dismissed_errors.Rows.Count - 1][0] = dataGridView_error_pipe_manifest.Rows[lista1[i]].Cells[0].Value;
                        dt_dismissed_errors.Rows[dt_dismissed_errors.Rows.Count - 1][1] = dataGridView_error_pipe_manifest.Rows[lista1[i]].Cells[1].Value;
                        dt_dismissed_errors.Rows[dt_dismissed_errors.Rows.Count - 1][2] = dataGridView_error_pipe_manifest.Rows[lista1[i]].Cells[3].Value;
                    }

                    if (W2 == null)
                    {
                        Functions.Create_a_new_worksheet_from_excel_by_name(filename, dismiss_errors_tab);

                    }
                    Functions.Transfer_datatable_to_existing_excel_spreadsheet_by_name(dt_dismissed_errors, filename, dismiss_errors_tab, false);

                    transfer_errors_to_panel(dt_errors);

                }


            }
            catch (System.Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

    }
}
