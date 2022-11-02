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
    public partial class Wgen_pipetally : Form
    {

        private ContextMenuStrip ContextMenuStrip_go_to_error;
        private ContextMenuStrip ContextMenuStrip_load_build_pipe_tally;

        System.Data.DataTable dt_errors;
        System.Data.DataTable dt_export;

        public System.Data.DataTable dt_dismissed_errors = null;
        public string dismiss_errors_tab = "Dsd errors gt";
        Microsoft.Office.Interop.Excel.Worksheet W2 = null;
        public string filename = "";

        int start_row = 2;
        double length_tolerance = 0.5;
        double pup_length_tolerance = 1;

        public Wgen_pipetally()
        {
            InitializeComponent();
            var toolStripMenuItem1 = new ToolStripMenuItem { Text = "Build Pipe Tally" };
            toolStripMenuItem1.Click += show_buld_pipe_tally_Click;


            ContextMenuStrip_load_build_pipe_tally = new ContextMenuStrip();
            ContextMenuStrip_load_build_pipe_tally.Items.AddRange(new ToolStripItem[] { toolStripMenuItem1 });


            var toolStripMenuItem2 = new ToolStripMenuItem { Text = "Go to error" };
            toolStripMenuItem2.Click += go_to_excel_point;


            ContextMenuStrip_go_to_error = new ContextMenuStrip();
            ContextMenuStrip_go_to_error.Items.AddRange(new ToolStripItem[] { toolStripMenuItem2 });

        }

        private void show_buld_pipe_tally_Click(object sender, EventArgs e)
        {
            Wgen_main_form.tpage_pipe_manifest.Hide();
            Wgen_main_form.tpage_pipe_tally.Hide();
            Wgen_main_form.tpage_weldmap.Hide();
            Wgen_main_form.tpage_blank.Hide();
            Wgen_main_form.tpage_allpts.Hide();
            Wgen_main_form.tpage_duplicates.Hide();
            Wgen_main_form.tpage_build_pipe_tally.Show();
        }
        private void set_enable_false(object sender)
        {
            List<System.Windows.Forms.Button> lista_butoane = new List<Button>();
            lista_butoane.Add(button_load_pipe_tally);
            lista_butoane.Add(button_pipe_tally_l);
            lista_butoane.Add(button_pipe_tally_nl);
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
            lista_butoane.Add(button_load_pipe_tally);
            lista_butoane.Add(button_pipe_tally_l);
            lista_butoane.Add(button_pipe_tally_nl);
            lista_butoane.Add(button_refresh_ws1);
            lista_butoane.Add(button_export_errors_to_xl);

            foreach (System.Windows.Forms.Button bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }



        private void button_load_pipe_tally_Click(object sender, EventArgs e)
        {
            Wgen_main_form.dt_ground_tally = Functions.Creaza_weldmap_pipe_tally_datatable_structure();
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

                            Wgen_main_form.dt_all_points = null;
                            Wgen_main_form.dt_weld_map = null;
                            Wgen_main_form.dt_pt_keep = null;
                            Wgen_main_form.dt_pt_move = null;


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

                            string colPM1 = "Pipe ID";
                            string colPM2 = "Heat";
                            string colPM3 = "Length";
                            string colPM4 = "Wall Thickness";
                            string colPM5 = "Diameter";
                            string colPM6 = "Grade";
                            string colPM7 = "Coating";
                            string colPM8 = "Manufacture";
                            string colPM9 = "DoubleJointNo";



                            Wgen_main_form.dt_ground_tally = Functions.Populate_data_table_from_excel(Wgen_main_form.dt_ground_tally, W1, start_row, textBox_1.Text, textBox_2.Text, textBox_3.Text, textBox_4.Text,
                                textBox_5.Text, textBox_6.Text, textBox_7.Text, textBox_8.Text, textBox_9.Text, textBox_10.Text, textBox_11.Text,
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


                            Wgen_main_form.dt_ground_tally.Columns.Add("rowno", typeof(int));

                            if (Wgen_main_form.dt_ground_tally.Rows.Count > 0)
                            {
                                for (int i = 0; i < Wgen_main_form.dt_ground_tally.Rows.Count; ++i)
                                {
                                    if (Wgen_main_form.dt_ground_tally.Rows[i]["DoubleJointNo"] != DBNull.Value)
                                    {
                                        string dj1 = Convert.ToString(Wgen_main_form.dt_ground_tally.Rows[i]["DoubleJointNo"]);
                                        if (dj1.ToLower() == "0" || dj1.ToLower() == "na" || dj1.ToLower() == @"n/a" || dj1.ToLower() == @"n\a")
                                        {
                                            Wgen_main_form.dt_ground_tally.Rows[i]["DoubleJointNo"] = DBNull.Value;
                                        }

                                    }
                                }

                                for (int i = 0; i < Wgen_main_form.dt_ground_tally.Rows.Count; ++i)
                                {

                                    if (Wgen_main_form.dt_ground_tally.Rows[i][col6] != DBNull.Value)
                                    {
                                        string wallT = Convert.ToString(Wgen_main_form.dt_ground_tally.Rows[i][col6]).Replace(" ", "");

                                        if (Functions.IsNumeric(wallT) == true)
                                        {
                                            double Wall1 = Convert.ToDouble(wallT);
                                            Wgen_main_form.dt_ground_tally.Rows[i][col6] = Convert.ToString(Wall1);
                                        }

                                    }

                                    if (Wgen_main_form.dt_ground_tally.Rows[i][col8] != DBNull.Value)
                                    {
                                        string grade = Convert.ToString(Wgen_main_form.dt_ground_tally.Rows[i][col8]).Replace(" ", "");
                                        Wgen_main_form.dt_ground_tally.Rows[i][col8] = grade;
                                    }
                                }

                                int nr_duplicates_pipe = 0;
                                int nr_null_values = 0;
                                int nr_duplicates_mmid = 0;
                                int nr_pipe_id_missing = 0;
                                int nr_heat_missing = 0;
                                int nr_length_off = 0;
                                int nr_not_numeric = 0;

                                int nr_dj_missmatch = 0;

                                dt_errors = new System.Data.DataTable();
                                dt_errors.Columns.Add("Point", typeof(string));
                                dt_errors.Columns.Add("Value", typeof(string));
                                dt_errors.Columns.Add("Excel", typeof(string));
                                dt_errors.Columns.Add("w1", typeof(Microsoft.Office.Interop.Excel.Worksheet));
                                dt_errors.Columns.Add("Error", typeof(string));

                                System.Data.DataTable dt_pipe_id_heat_duplicates = new System.Data.DataTable();
                                dt_pipe_id_heat_duplicates.Columns.Add("mmid", typeof(string));
                                dt_pipe_id_heat_duplicates.Columns.Add("pipe", typeof(string));
                                dt_pipe_id_heat_duplicates.Columns.Add("heat", typeof(string));
                                dt_pipe_id_heat_duplicates.Columns.Add("originall", typeof(double));
                                dt_pipe_id_heat_duplicates.Columns.Add("newl", typeof(double));
                                dt_pipe_id_heat_duplicates.Columns.Add("i", typeof(int));
                                dt_pipe_id_heat_duplicates.Columns.Add("solved", typeof(bool));


                                System.Data.DataTable dt_lengths = new System.Data.DataTable();
                                dt_lengths.Columns.Add("mmid", typeof(string));
                                dt_lengths.Columns.Add("pipe", typeof(string));
                                dt_lengths.Columns.Add("heat", typeof(string));
                                dt_lengths.Columns.Add("pipetallyl", typeof(double));
                                dt_lengths.Columns.Add("originall", typeof(double));
                                dt_lengths.Columns.Add("newl", typeof(double));
                                dt_lengths.Columns.Add("i", typeof(int));
                                dt_lengths.Columns.Add("solved", typeof(bool));
                                dt_lengths.Columns.Add("message", typeof(string));


                                System.Data.DataTable dt_pipe_heat = new System.Data.DataTable();
                                dt_pipe_heat.Columns.Add("mmid", typeof(string));
                                dt_pipe_heat.Columns.Add("pipe", typeof(string));
                                dt_pipe_heat.Columns.Add("heat", typeof(string));
                                dt_pipe_heat.Columns.Add("i", typeof(int));


                                var duplicates2 = Wgen_main_form.dt_ground_tally.AsEnumerable().GroupBy(i => new { MMid = i.Field<string>(col1) }).Where(g => g.Count() > 1).Select(g => new { g.Key.MMid }).ToList();
                                var duplicates1 = Wgen_main_form.dt_ground_tally.AsEnumerable().GroupBy(i => new { Pipeid = i.Field<string>(col2), Heat = i.Field<string>(col3) }).Where(g => g.Count() > 1).Select(g => new { g.Key.Pipeid, g.Key.Heat }).ToList();

                                System.Data.DataTable dt2 = new System.Data.DataTable();
                                dt2.Columns.Add(col2, typeof(string));
                                dt2.Columns.Add(col3, typeof(string));

                                System.Data.DataTable dt3 = new System.Data.DataTable();
                                dt3.Columns.Add(col1, typeof(string));

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
                                        if (duplicates2[i].MMid != null)
                                        {
                                            string duplicat_val1 = Convert.ToString(duplicates2[i].MMid);
                                            dt3.Rows.Add();
                                            dt3.Rows[dt3.Rows.Count - 1][0] = duplicat_val1;
                                        }
                                    }
                                }

                                DataSet dataset1 = new DataSet();
                                if (Wgen_main_form.dt_ground_tally != null) Wgen_main_form.dt_ground_tally.TableName = "t1";
                                if (Wgen_main_form.dt_pipe_list != null) Wgen_main_form.dt_pipe_list.TableName = "t2";
                                if (dt2 != null) dt2.TableName = "t4";
                                if (dt3 != null) dt3.TableName = "t5";
                                if (Wgen_main_form.dt_double_joint != null) Wgen_main_form.dt_pipe_list.TableName = "t6";

                                dataset1.Tables.Add(Wgen_main_form.dt_ground_tally);
                                if (Wgen_main_form.dt_pipe_list != null && Wgen_main_form.dt_pipe_list.Rows.Count > 0) dataset1.Tables.Add(Wgen_main_form.dt_pipe_list);
                                if (Wgen_main_form.dt_double_joint != null && Wgen_main_form.dt_double_joint.Rows.Count > 0) dataset1.Tables.Add(Wgen_main_form.dt_double_joint);



                                dataset1.Tables.Add(dt2);
                                dataset1.Tables.Add(dt3);

                                DataRelation rel_duplic_pipeid = new DataRelation("xxx", Wgen_main_form.dt_ground_tally.Columns[col2], dt2.Columns[col2], false);
                                DataRelation rel_duplic_heat = new DataRelation("xxx1", Wgen_main_form.dt_ground_tally.Columns[col3], dt2.Columns[col3], false);
                                DataRelation rel_duplic_mmid = new DataRelation("xxx2", Wgen_main_form.dt_ground_tally.Columns[col1], dt3.Columns[col1], false);

                                DataRelation rel_pipe_id = null;
                                DataRelation rel_heat = null;
                                DataRelation relation_dj = null;
                                DataRelation relation_wt = null;
                                DataRelation relation8 = null;
                                DataRelation relation_pipe_grade = null;
                                DataRelation relation10 = null;
                                DataRelation relation11 = null;
                                DataRelation relation_double_joint1 = null;
                                DataRelation relation_double_joint2 = null;

                                if (Wgen_main_form.dt_pipe_list != null && Wgen_main_form.dt_pipe_list.Rows.Count > 0)
                                {
                                    rel_pipe_id = new DataRelation("xxx3", Wgen_main_form.dt_ground_tally.Columns[col2], Wgen_main_form.dt_pipe_list.Columns[colPM1], false);
                                    rel_heat = new DataRelation("xxx4", Wgen_main_form.dt_ground_tally.Columns[col3], Wgen_main_form.dt_pipe_list.Columns[colPM2], false);
                                    relation_dj = new DataRelation("xxx5", Wgen_main_form.dt_ground_tally.Columns[col11], Wgen_main_form.dt_pipe_list.Columns[colPM9], false);
                                    relation_wt = new DataRelation("xxx7", Wgen_main_form.dt_ground_tally.Columns[col6], Wgen_main_form.dt_pipe_list.Columns[colPM4], false);
                                    relation8 = new DataRelation("xxx8", Wgen_main_form.dt_ground_tally.Columns[col7], Wgen_main_form.dt_pipe_list.Columns[colPM5], false);
                                    relation_pipe_grade = new DataRelation("xxx9", Wgen_main_form.dt_ground_tally.Columns[col8], Wgen_main_form.dt_pipe_list.Columns[colPM6], false);
                                    relation10 = new DataRelation("xxx10", Wgen_main_form.dt_ground_tally.Columns[col9], Wgen_main_form.dt_pipe_list.Columns[colPM7], false);
                                    relation11 = new DataRelation("xxx11", Wgen_main_form.dt_ground_tally.Columns[col10], Wgen_main_form.dt_pipe_list.Columns[colPM8], false);
                                }

                                if (Wgen_main_form.dt_double_joint != null && Wgen_main_form.dt_double_joint.Rows.Count > 0)
                                {
                                    relation_double_joint1 = new DataRelation("xxx13", Wgen_main_form.dt_double_joint.Columns["double_joint"], Wgen_main_form.dt_ground_tally.Columns[col11], false);
                                    relation_double_joint2 = new DataRelation("xxx14", Wgen_main_form.dt_ground_tally.Columns[col11], Wgen_main_form.dt_double_joint.Columns["double_joint"], false);
                                }


                                dataset1.Relations.Add(rel_duplic_pipeid);
                                dataset1.Relations.Add(rel_duplic_heat);
                                dataset1.Relations.Add(rel_duplic_mmid);

                                if (rel_pipe_id != null) dataset1.Relations.Add(rel_pipe_id);
                                if (rel_heat != null) dataset1.Relations.Add(rel_heat);
                                if (relation_dj != null) dataset1.Relations.Add(relation_dj);
                                if (relation_wt != null) dataset1.Relations.Add(relation_wt);
                                if (relation8 != null) dataset1.Relations.Add(relation8);
                                if (relation_pipe_grade != null) dataset1.Relations.Add(relation_pipe_grade);
                                if (relation10 != null) dataset1.Relations.Add(relation10);
                                if (relation11 != null) dataset1.Relations.Add(relation11);
                                if (relation_double_joint1 != null) dataset1.Relations.Add(relation_double_joint1);
                                if (relation_double_joint2 != null) dataset1.Relations.Add(relation_double_joint2);


                                // nr_duplicates_pipe = dt2.Rows.Count;
                                nr_duplicates_mmid = dt3.Rows.Count;

                                for (int i = 0; i < Wgen_main_form.dt_ground_tally.Rows.Count; ++i)
                                {
                                    Wgen_main_form.dt_ground_tally.Rows[i]["rowno"] = i;

                                    if (Wgen_main_form.dt_ground_tally.Rows[i][col1] != DBNull.Value && Wgen_main_form.dt_ground_tally.Rows[i][col2] != DBNull.Value &&
                                        Wgen_main_form.dt_ground_tally.Rows[i][col3] != DBNull.Value && Wgen_main_form.dt_ground_tally.Rows[i][col4] != DBNull.Value &&
                                        Wgen_main_form.dt_ground_tally.Rows[i][col6] != DBNull.Value &&
                                        Wgen_main_form.dt_ground_tally.Rows[i][col7] != DBNull.Value && Wgen_main_form.dt_ground_tally.Rows[i][col8] != DBNull.Value &&
                                        Wgen_main_form.dt_ground_tally.Rows[i][col9] != DBNull.Value && Wgen_main_form.dt_ground_tally.Rows[i][col10] != DBNull.Value)
                                    {
                                        string string_mmid = Convert.ToString(Wgen_main_form.dt_ground_tally.Rows[i][col1]);
                                        string string_pipe = Convert.ToString(Wgen_main_form.dt_ground_tally.Rows[i][col2]);
                                        string string_heat = Convert.ToString(Wgen_main_form.dt_ground_tally.Rows[i][col3]);
                                        string string_orig_len = Convert.ToString(Wgen_main_form.dt_ground_tally.Rows[i][col4]);
                                        string string_new_len = "";
                                        if (Wgen_main_form.dt_ground_tally.Rows[i][col5] != DBNull.Value) string_new_len = Convert.ToString(Wgen_main_form.dt_ground_tally.Rows[i][col5]);
                                        string string_wt = Convert.ToString(Wgen_main_form.dt_ground_tally.Rows[i][col6]);
                                        string string_diam = Convert.ToString(Wgen_main_form.dt_ground_tally.Rows[i][col7]);
                                        string string_grd = Convert.ToString(Wgen_main_form.dt_ground_tally.Rows[i][col8]);
                                        string string_coat = Convert.ToString(Wgen_main_form.dt_ground_tally.Rows[i][col9]);
                                        string string_man = Convert.ToString(Wgen_main_form.dt_ground_tally.Rows[i][col10]);


                                        #region heat and pipeid for normal double joint = 0
                                        if (Wgen_main_form.dt_ground_tally.Rows[i][col11] == DBNull.Value)
                                        {
                                            dt_pipe_heat.Rows.Add();
                                            dt_pipe_heat.Rows[dt_pipe_heat.Rows.Count - 1]["mmid"] = string_mmid;
                                            dt_pipe_heat.Rows[dt_pipe_heat.Rows.Count - 1]["pipe"] = string_pipe;
                                            dt_pipe_heat.Rows[dt_pipe_heat.Rows.Count - 1]["heat"] = string_heat;
                                            dt_pipe_heat.Rows[dt_pipe_heat.Rows.Count - 1]["i"] = i;
                                        }
                                        #endregion



                                        #region doublejoint normal
                                        if (Wgen_main_form.dt_double_joint == null)
                                        {
                                            if (Wgen_main_form.dt_ground_tally.Rows[i][col11] != DBNull.Value && relation_dj != null)
                                            {
                                                string string_dj = Convert.ToString(Wgen_main_form.dt_ground_tally.Rows[i][col11]);
                                                if (string_dj.Replace(" ", "") != "")
                                                {
                                                    if (Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(relation_dj).Length == 0)
                                                    {
                                                        dt_errors.Rows.Add();
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = string_mmid;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = string_dj;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_11.Text + Convert.ToString(i + start_row);
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Double Joint Number not found in the Pipe Manifest";
                                                        ++nr_dj_missmatch;
                                                    }
                                                    else
                                                    {
                                                        System.Data.DataRow row1 = Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(relation_dj)[0];
                                                        string string_dj1 = Convert.ToString(row1[colPM9]);
                                                        if (string_dj.ToUpper() != string_dj1.ToUpper())
                                                        {
                                                            dt_errors.Rows.Add();
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = string_mmid;
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = string_dj + " vs " + string_dj1;
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_11.Text + Convert.ToString(i + start_row);
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Double Joint Number mismatch with the Pipe Manifest";
                                                            ++nr_dj_missmatch;
                                                        }
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                if (rel_pipe_id != null)
                                                {
                                                    if (Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(rel_pipe_id).Length > 0)
                                                    {
                                                        System.Data.DataRow row1 = Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(rel_pipe_id)[0];
                                                        if (row1[colPM9] != DBNull.Value)
                                                        {
                                                            dt_errors.Rows.Add();
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = string_mmid;
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = "*** vs " + Convert.ToString(row1[colPM9]);
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_11.Text + Convert.ToString(i + start_row);
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Double Joint Number mismatch with the Pipe Manifest";
                                                            ++nr_dj_missmatch;
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                        #endregion

                                        #region wall thickness

                                        if (Wgen_main_form.dt_ground_tally.Rows[i][col11] == DBNull.Value)
                                        {
                                            if (relation_wt != null &&
                                                rel_pipe_id != null && rel_heat != null &&
                                                Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(rel_pipe_id).Length > 0 &&
                                                Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(rel_heat).Length > 0)
                                            {

                                                if (string_wt.Replace(" ", "") != "")
                                                {
                                                    if (Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(relation_wt).Length == 0)
                                                    {

                                                        dt_errors.Rows.Add();
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = string_mmid;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = string_wt;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_6.Text + Convert.ToString(i + start_row);
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Wall thickness not found in the Pipe Manifest";
                                                        ++nr_dj_missmatch;
                                                    }
                                                    else
                                                    {
                                                        System.Data.DataRow row1 = Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(relation_wt)[0];
                                                        string string_wall1 = Convert.ToString(row1[colPM4]);
                                                        if (string_wt.ToUpper() != string_wall1.ToUpper())
                                                        {
                                                            dt_errors.Rows.Add();
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = string_mmid;
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = string_wt + " vs " + string_wall1;
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_6.Text + Convert.ToString(i + start_row);
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Wall thickness mismatch with the Pipe Manifest";
                                                            ++nr_dj_missmatch;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        #endregion

                                        #region diameter
                                        if (Wgen_main_form.dt_ground_tally.Rows[i][col11] == DBNull.Value)
                                        {
                                            if (Wgen_main_form.dt_ground_tally.Rows[i][col7] != DBNull.Value && relation8 != null)
                                            {
                                                string string_wall = Convert.ToString(Wgen_main_form.dt_ground_tally.Rows[i][col7]);
                                                if (string_wall.Replace(" ", "") != "")
                                                {
                                                    if (Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(relation8).Length == 0)
                                                    {
                                                        dt_errors.Rows.Add();
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = string_mmid;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = string_wall;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_7.Text + Convert.ToString(i + start_row);
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Diameter not found in the Pipe Manifest";
                                                        ++nr_dj_missmatch;
                                                    }
                                                    else
                                                    {
                                                        System.Data.DataRow row1 = Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(relation8)[0];
                                                        string string_wall1 = Convert.ToString(row1[colPM5]);
                                                        if (string_wall.ToUpper() != string_wall1.ToUpper())
                                                        {
                                                            dt_errors.Rows.Add();
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = string_mmid;
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = string_wall + " vs " + string_wall1;
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_7.Text + Convert.ToString(i + start_row);
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Diameter mismatch with the Pipe Manifest";
                                                            ++nr_dj_missmatch;
                                                        }
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                if (rel_pipe_id != null)
                                                {
                                                    if (Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(rel_pipe_id).Length > 0)
                                                    {
                                                        System.Data.DataRow row1 = Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(rel_pipe_id)[0];
                                                        if (row1[colPM5] != DBNull.Value)
                                                        {
                                                            dt_errors.Rows.Add();
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = string_mmid;
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = "*** vs " + Convert.ToString(row1[colPM5]);
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_7.Text + Convert.ToString(i + start_row);
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Diameter mismatch with the Pipe Manifest";
                                                            ++nr_dj_missmatch;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        #endregion

                                        #region grade
                                        if (Wgen_main_form.dt_ground_tally.Rows[i][col11] == DBNull.Value)
                                        {
                                            if (Wgen_main_form.dt_ground_tally.Rows[i][col8] != DBNull.Value && relation_pipe_grade != null)
                                            {
                                                string string_wall = Convert.ToString(Wgen_main_form.dt_ground_tally.Rows[i][col8]);
                                                if (string_wall.Replace(" ", "") != "")
                                                {
                                                    if (Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(relation_pipe_grade).Length == 0)
                                                    {
                                                        dt_errors.Rows.Add();
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = string_mmid;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = string_wall;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_8.Text + Convert.ToString(i + start_row);
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Grade not found in the Pipe Manifest";
                                                        ++nr_dj_missmatch;
                                                    }
                                                    else
                                                    {
                                                        System.Data.DataRow row1 = Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(relation_pipe_grade)[0];
                                                        string string_wall1 = Convert.ToString(row1[colPM6]);
                                                        if (string_wall.ToUpper() != string_wall1.ToUpper())
                                                        {
                                                            dt_errors.Rows.Add();
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = string_mmid;
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = string_wall + " vs " + string_wall1;
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_8.Text + Convert.ToString(i + start_row);
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Grade mismatch with the Pipe Manifest";
                                                            ++nr_dj_missmatch;
                                                        }
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                if (rel_pipe_id != null)
                                                {
                                                    if (Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(rel_pipe_id).Length > 0)
                                                    {
                                                        System.Data.DataRow row1 = Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(rel_pipe_id)[0];
                                                        if (row1[colPM6] != DBNull.Value)
                                                        {
                                                            dt_errors.Rows.Add();
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = string_mmid;
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = "*** vs " + Convert.ToString(row1[colPM6]);
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_8.Text + Convert.ToString(i + start_row);
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Grade mismatch with the Pipe Manifest";
                                                            ++nr_dj_missmatch;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        #endregion

                                        #region Coating
                                        if (Wgen_main_form.dt_ground_tally.Rows[i][col11] == DBNull.Value)
                                        {
                                            if (Wgen_main_form.dt_ground_tally.Rows[i][col9] != DBNull.Value && relation10 != null)
                                            {
                                                string string_wall = Convert.ToString(Wgen_main_form.dt_ground_tally.Rows[i][col9]);
                                                if (string_wall.Replace(" ", "") != "")
                                                {
                                                    if (Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(relation10).Length == 0)
                                                    {
                                                        dt_errors.Rows.Add();
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = string_mmid;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = string_wall;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_9.Text + Convert.ToString(i + start_row);
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Coating not found in the Pipe Manifest";
                                                        ++nr_dj_missmatch;
                                                    }
                                                    else
                                                    {
                                                        System.Data.DataRow row1 = Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(relation10)[0];
                                                        string string_wall1 = Convert.ToString(row1[colPM7]);
                                                        if (string_wall.ToUpper() != string_wall1.ToUpper())
                                                        {
                                                            dt_errors.Rows.Add();
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = string_mmid;
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = string_wall + " vs " + string_wall1;
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_9.Text + Convert.ToString(i + start_row);
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Coating mismatch with the Pipe Manifest";
                                                            ++nr_dj_missmatch;
                                                        }
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                if (rel_pipe_id != null)
                                                {
                                                    if (Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(rel_pipe_id).Length > 0)
                                                    {
                                                        System.Data.DataRow row1 = Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(rel_pipe_id)[0];
                                                        if (row1[colPM7] != DBNull.Value)
                                                        {
                                                            dt_errors.Rows.Add();
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = string_mmid;
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = "*** vs " + Convert.ToString(row1[colPM7]);
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_9.Text + Convert.ToString(i + start_row);
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Coating mismatch with the Pipe Manifest";
                                                            ++nr_dj_missmatch;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        #endregion

                                        #region Manufacture
                                        if (Wgen_main_form.dt_ground_tally.Rows[i][col11] == DBNull.Value)
                                        {
                                            if (Wgen_main_form.dt_ground_tally.Rows[i][col10] != DBNull.Value && relation11 != null)
                                            {
                                                string string_wall = Convert.ToString(Wgen_main_form.dt_ground_tally.Rows[i][col10]);
                                                if (string_wall.Replace(" ", "") != "")
                                                {
                                                    if (Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(relation11).Length == 0)
                                                    {
                                                        dt_errors.Rows.Add();
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = string_mmid;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = string_wall;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_10.Text + Convert.ToString(i + start_row);
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Manufacture not found in the Pipe Manifest";
                                                        ++nr_dj_missmatch;
                                                    }
                                                    else
                                                    {
                                                        System.Data.DataRow row1 = Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(relation11)[0];
                                                        string string_wall1 = Convert.ToString(row1[colPM8]);
                                                        if (string_wall.ToUpper() != string_wall1.ToUpper())
                                                        {
                                                            dt_errors.Rows.Add();
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = string_mmid;
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = string_wall + " vs " + string_wall1;
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_10.Text + Convert.ToString(i + start_row);
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Manufacture mismatch with the Pipe Manifest";
                                                            ++nr_dj_missmatch;
                                                        }
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                if (rel_pipe_id != null)
                                                {
                                                    if (Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(rel_pipe_id).Length > 0)
                                                    {
                                                        System.Data.DataRow row1 = Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(rel_pipe_id)[0];
                                                        if (row1[colPM8] != DBNull.Value)
                                                        {
                                                            dt_errors.Rows.Add();
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = string_mmid;
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = "*** vs " + Convert.ToString(row1[colPM8]);
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_10.Text + Convert.ToString(i + start_row);
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Manufacture mismatch with the Pipe Manifest";
                                                            ++nr_dj_missmatch;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        #endregion


                                        if (Wgen_main_form.dt_ground_tally.Rows[i][col11] == DBNull.Value)
                                        {
                                            #region non numeric length
                                            if (Functions.IsNumeric(string_orig_len) == false)
                                            {
                                                dt_errors.Rows.Add();
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = string_mmid;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = string_orig_len;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_4.Text + Convert.ToString(i + start_row);
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Non Numeric Length";
                                                ++nr_not_numeric;
                                            }
                                            #endregion

                                            #region non numeric pipe diam
                                            if (Functions.IsNumeric(string_diam) == false)
                                            {
                                                dt_errors.Rows.Add();
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = string_mmid;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = string_diam;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_7.Text + Convert.ToString(i + start_row);
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Non Numeric Pipe Diameter";
                                                ++nr_not_numeric;
                                            }
                                            else
                                            {
                                                if (i == 0) Wgen_main_form.pipe_diam = Convert.ToDouble(string_diam);
                                            }
                                            #endregion

                                            #region non numeric length
                                            if (Wgen_main_form.dt_ground_tally.Rows[i][col5] != DBNull.Value)
                                            {
                                                if (Functions.IsNumeric(string_new_len) == false)
                                                {
                                                    dt_errors.Rows.Add();
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = string_mmid;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = string_new_len;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_5.Text + Convert.ToString(i + start_row);
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Non Numeric Length";
                                                    ++nr_not_numeric;
                                                }
                                            }
                                            #endregion


                                            if (Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(rel_duplic_pipeid).Length > 0 && Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(rel_duplic_heat).Length > 0)
                                            {
                                                for (int j = 0; j < Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(rel_duplic_pipeid).Length; ++j)
                                                {
                                                    #region builds dt_pipeid_heat
                                                    //if (Functions.IsNumeric(string_new_len) == true && Functions.IsNumeric(string_orig_len) == true)
                                                    //{

                                                    double ol = 0;
                                                    double nl = 0;
                                                    if (Functions.IsNumeric(string_new_len) == true)
                                                    {
                                                        nl = Convert.ToDouble(string_new_len);
                                                    }
                                                    if (Functions.IsNumeric(string_orig_len) == true)
                                                    {
                                                        ol = Convert.ToDouble(string_orig_len);
                                                    }
                                                    dt_pipe_id_heat_duplicates.Rows.Add();
                                                    dt_pipe_id_heat_duplicates.Rows[dt_pipe_id_heat_duplicates.Rows.Count - 1][0] = string_mmid;
                                                    dt_pipe_id_heat_duplicates.Rows[dt_pipe_id_heat_duplicates.Rows.Count - 1][1] = string_pipe;
                                                    dt_pipe_id_heat_duplicates.Rows[dt_pipe_id_heat_duplicates.Rows.Count - 1][2] = string_heat;
                                                    dt_pipe_id_heat_duplicates.Rows[dt_pipe_id_heat_duplicates.Rows.Count - 1][3] = ol;
                                                    dt_pipe_id_heat_duplicates.Rows[dt_pipe_id_heat_duplicates.Rows.Count - 1][4] = nl;
                                                    dt_pipe_id_heat_duplicates.Rows[dt_pipe_id_heat_duplicates.Rows.Count - 1][5] = i;
                                                    //}
                                                    #endregion
                                                }
                                            }

                                            #region duplicates mmid
                                            if (Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(rel_duplic_mmid).Length > 0)
                                            {
                                                for (int j = 0; j < Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(rel_duplic_mmid).Length; ++j)
                                                {
                                                    string val1 = Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(rel_duplic_mmid)[j][col1].ToString();
                                                    dt_errors.Rows.Add();
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = val1;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = val1;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_1.Text + Convert.ToString(i + start_row);
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Duplicate MMID";
                                                }
                                            }
                                            #endregion

                                            bool noid1 = true;
                                            bool noh1 = true;
                                            if (rel_pipe_id != null)
                                            {
                                                if (Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(rel_pipe_id).Length == 0)
                                                {
                                                    noid1 = false;
                                                }
                                            }

                                            if (rel_heat != null)
                                            {
                                                if (Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(rel_heat).Length == 0)
                                                {
                                                    noh1 = false;
                                                }
                                            }

                                            if (rel_pipe_id != null && rel_heat != null)
                                            {
                                                #region length calcs for non double joints
                                                // aici verifici sa nu ai dj sau 3j

                                                if (Wgen_main_form.dt_ground_tally.Rows[i][col11] == DBNull.Value &&
                                                    Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(rel_pipe_id).Length > 0 &&
                                                    Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(rel_heat).Length > 0)
                                                {
                                                    if (Functions.IsNumeric(string_orig_len) == true)
                                                    {
                                                        double orig_len = Convert.ToDouble(string_orig_len);
                                                        double new_len = 0;
                                                        if (Functions.IsNumeric(string_new_len) == true)
                                                        {
                                                            new_len = Convert.ToDouble(string_new_len);
                                                        }

                                                        for (int j = 0; j < Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(rel_pipe_id).Length; ++j)
                                                        {
                                                            string heat1 = Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(rel_pipe_id)[j][colPM2].ToString();
                                                            if (string_heat == heat1)
                                                            {
                                                                string string_len_pipe_manifest = Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(rel_pipe_id)[j][colPM3].ToString();

                                                                if (Functions.IsNumeric(string_len_pipe_manifest) == true)
                                                                {
                                                                    double len_pipe_manifest = Convert.ToDouble(string_len_pipe_manifest);

                                                                    if (Math.Abs(Math.Round(Convert.ToDouble(string_orig_len), 2) - Math.Round(len_pipe_manifest, 2)) > length_tolerance)
                                                                    {
                                                                        dt_errors.Rows.Add();
                                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = string_mmid;
                                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = string_orig_len + " vs " + string_len_pipe_manifest;
                                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_4.Text + Convert.ToString(i + start_row);
                                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Original Length not matching the pipe manifest value";
                                                                    }
                                                                    else
                                                                    {
                                                                        dt_lengths.Rows.Add();
                                                                        dt_lengths.Rows[dt_lengths.Rows.Count - 1][0] = string_mmid;
                                                                        dt_lengths.Rows[dt_lengths.Rows.Count - 1][1] = string_pipe;
                                                                        dt_lengths.Rows[dt_lengths.Rows.Count - 1][2] = string_heat;
                                                                        dt_lengths.Rows[dt_lengths.Rows.Count - 1][3] = len_pipe_manifest;
                                                                        dt_lengths.Rows[dt_lengths.Rows.Count - 1][4] = orig_len;
                                                                        dt_lengths.Rows[dt_lengths.Rows.Count - 1][5] = new_len;
                                                                        dt_lengths.Rows[dt_lengths.Rows.Count - 1][6] = i;
                                                                    }
                                                                }

                                                                string wt = Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(rel_pipe_id)[j][colPM4].ToString();
                                                                if (string_wt != wt)
                                                                {
                                                                    dt_errors.Rows.Add();
                                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = string_mmid;
                                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = string_wt + " vs " + wt;
                                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_6.Text + Convert.ToString(i + start_row);
                                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Wall Thickness not matching the pipe manifest value";
                                                                    ++nr_dj_missmatch;
                                                                }

                                                                string dia = Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(rel_pipe_id)[j][colPM5].ToString();
                                                                if (string_diam != dia)
                                                                {
                                                                    dt_errors.Rows.Add();
                                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = string_mmid;
                                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = string_diam + " vs " + dia;
                                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_7.Text + Convert.ToString(i + start_row);
                                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Diameter not matching the pipe manifest value";
                                                                    ++nr_dj_missmatch;
                                                                }
                                                                string grd = Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(rel_pipe_id)[j][colPM6].ToString();
                                                                if (string_grd != grd)
                                                                {
                                                                    dt_errors.Rows.Add();
                                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = string_mmid;
                                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = string_grd + " vs " + grd;
                                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_8.Text + Convert.ToString(i + start_row);
                                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Grade not matching the pipe manifest value";
                                                                    ++nr_dj_missmatch;
                                                                }
                                                                string coat = Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(rel_pipe_id)[j][colPM7].ToString();
                                                                if (string_coat != coat)
                                                                {
                                                                    dt_errors.Rows.Add();
                                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = string_mmid;
                                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = string_coat + " vs " + coat;
                                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_9.Text + Convert.ToString(i + start_row);
                                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Coating not matching the pipe manifest value";
                                                                    ++nr_dj_missmatch;
                                                                }
                                                                string man = Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(rel_pipe_id)[j][colPM8].ToString();
                                                                if (string_man != man)
                                                                {
                                                                    dt_errors.Rows.Add();
                                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = string_mmid;
                                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = string_man + " vs " + man;
                                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_10.Text + Convert.ToString(i + start_row);
                                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Manufacture not matching the pipe manifest value";
                                                                    ++nr_dj_missmatch;
                                                                }

                                                            } //if (string_heat == heat1)
                                                        }
                                                    }

                                                } // if (Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(rel_pipe_id).Length > 0 && Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(rel_heat).Length > 0)

                                                #endregion
                                            }
                                            if (Wgen_main_form.dt_pipe_list != null && Wgen_main_form.dt_pipe_list.Rows.Count > 0)
                                            {
                                                #region pipe id not found in pipe manifest
                                                if (noid1 == false && Wgen_main_form.incomplete_pipe_manifest == false)
                                                {
                                                    string val1 = Wgen_main_form.dt_ground_tally.Rows[i][col2].ToString();
                                                    dt_errors.Rows.Add();
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = string_mmid;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = val1;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_2.Text + Convert.ToString(i + start_row);
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Pipe ID not found in Pipe manifest";
                                                    ++nr_pipe_id_missing;
                                                }
                                                #endregion

                                                #region heat not found in pipe manifest
                                                if (noh1 == false && Wgen_main_form.incomplete_pipe_manifest == false)
                                                {
                                                    string val1 = Wgen_main_form.dt_ground_tally.Rows[i][col3].ToString();
                                                    dt_errors.Rows.Add();
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = string_mmid;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = val1;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_3.Text + Convert.ToString(i + start_row);
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Heat # not found in Pipe manifest";
                                                    ++nr_heat_missing;
                                                }
                                                #endregion
                                            }
                                        }
                                    }
                                } //for (int i = 0; i < Wgen_main_form.dt_ground_tally.Rows.Count; ++i)


                                #region double joints 
                                if (Wgen_main_form.dt_double_joint != null && Wgen_main_form.dt_double_joint.Rows.Count > 0)
                                {
                                    for (int i = 0; i < Wgen_main_form.dt_double_joint.Rows.Count; ++i)
                                    {
                                        if (Wgen_main_form.dt_double_joint.Rows[i]["pipeid"] != DBNull.Value && Wgen_main_form.dt_double_joint.Rows[i]["heat"] != DBNull.Value &&
                                            Wgen_main_form.dt_double_joint.Rows[i]["total_len"] != DBNull.Value && Wgen_main_form.dt_double_joint.Rows[i]["double_joint"] != DBNull.Value)
                                        {
                                            string dj1 = Convert.ToString(Wgen_main_form.dt_double_joint.Rows[i]["double_joint"]);
                                            string pipe1 = Convert.ToString(Wgen_main_form.dt_double_joint.Rows[i]["pipeid"]);
                                            string h1 = Convert.ToString(Wgen_main_form.dt_double_joint.Rows[i]["heat"]);
                                            List<string> Listap1 = new List<string>();
                                            List<string> Listah1 = new List<string>();



                                            if (pipe1.Contains("/") == true)
                                            {
                                                string[] pp1 = pipe1.Split(Convert.ToChar("/"));
                                                for (int k = 0; k < pp1.Length; ++k)
                                                {
                                                    if (Listap1.Contains(pp1[k]) == false)
                                                    {
                                                        Listap1.Add(pp1[k]);
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                if (Listap1.Contains(pipe1) == false)
                                                {
                                                    Listap1.Add(pipe1);
                                                }
                                            }



                                            if (h1.Contains("/") == true)
                                            {
                                                string[] pp1 = h1.Split(Convert.ToChar("/"));
                                                for (int k = 0; k < pp1.Length; ++k)
                                                {
                                                    if (Listah1.Contains(pp1[k]) == false)
                                                    {
                                                        Listah1.Add(pp1[k]);
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                if (Listah1.Contains(h1) == false)
                                                {
                                                    Listah1.Add(h1);
                                                }
                                            }

                                            for (int k = 0; k < dt_pipe_heat.Rows.Count; ++k)
                                            {
                                                string mmid2 = Convert.ToString(dt_pipe_heat.Rows[k]["mmid"]);
                                                string pipe2 = Convert.ToString(dt_pipe_heat.Rows[k]["pipe"]);
                                                string h2 = Convert.ToString(dt_pipe_heat.Rows[k]["heat"]);
                                                int index2 = Convert.ToInt32(dt_pipe_heat.Rows[k]["i"]);

                                                if (Listap1.Contains(pipe2) == true && Listah1.Contains(h2) == true)
                                                {
                                                    dt_errors.Rows.Add();
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = mmid2;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = "PipeID: " + pipe2 + " & Heat:" + h2;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_11.Text + Convert.ToString(index2 + start_row);
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Double joint number missing";
                                                }
                                            }


                                            if (Wgen_main_form.dt_double_joint.Rows[i].GetChildRows(relation_double_joint1).Length > 0)
                                            {
                                                string pipeid1 = Convert.ToString(Wgen_main_form.dt_double_joint.Rows[i]["pipeid"]);
                                                List<string> Lista_pipeid = new List<string>();
                                                List<int> Lista_index1 = new List<int>();

                                                string heat1 = Convert.ToString(Wgen_main_form.dt_double_joint.Rows[i]["heat"]);
                                                List<string> Lista_heat = new List<string>();
                                                List<int> Lista_index2 = new List<int>();

                                                string wt1 = Convert.ToString(Wgen_main_form.dt_double_joint.Rows[i]["wt"]);
                                                List<string> Lista_wt = new List<string>();
                                                List<int> Lista_index3 = new List<int>();

                                                string od1 = Convert.ToString(Wgen_main_form.dt_double_joint.Rows[i]["od"]);
                                                List<string> Lista_od = new List<string>();
                                                List<int> Lista_index4 = new List<int>();

                                                string grade1 = Convert.ToString(Wgen_main_form.dt_double_joint.Rows[i]["grade"]);
                                                List<string> Lista_grade = new List<string>();
                                                List<int> Lista_index5 = new List<int>();

                                                string coat1 = Convert.ToString(Wgen_main_form.dt_double_joint.Rows[i]["coating"]);
                                                List<string> Lista_coat = new List<string>();
                                                List<int> Lista_index6 = new List<int>();

                                                string manuf1 = Convert.ToString(Wgen_main_form.dt_double_joint.Rows[i]["manufacturer"]);
                                                List<string> Lista_manuf = new List<string>();
                                                List<int> Lista_index7 = new List<int>();

                                                double len1 = Convert.ToDouble(Wgen_main_form.dt_double_joint.Rows[i]["total_len"]);
                                                double len2 = 0;

                                                for (int j = 0; j < Wgen_main_form.dt_double_joint.Rows[i].GetChildRows(relation_double_joint1).Length; ++j)
                                                {
                                                    int index1 = Convert.ToInt32(Wgen_main_form.dt_double_joint.Rows[i].GetChildRows(relation_double_joint1)[j]["rowno"]);
                                                    len2 = len2 + Convert.ToDouble(Wgen_main_form.dt_double_joint.Rows[i].GetChildRows(relation_double_joint1)[j][col4]);
                                                    string pipeid2 = Convert.ToString(Wgen_main_form.dt_double_joint.Rows[i].GetChildRows(relation_double_joint1)[j][col2]);
                                                    if (pipeid2.Contains("/") == true)
                                                    {
                                                        string[] pid2 = pipeid2.Split(Convert.ToChar("/"));
                                                        for (int k = 0; k < pid2.Length; ++k)
                                                        {
                                                            if (pipeid1.ToLower().Replace(" ", "").Contains(pid2[k].ToLower().Replace(" ", "")) == false)
                                                            {
                                                                Lista_pipeid.Add(pid2[k]);
                                                                Lista_index1.Add(index1);
                                                            }
                                                        }
                                                    }
                                                    else
                                                    {
                                                        if (pipeid1.ToLower().Replace(" ", "").Contains(pipeid2.ToLower().Replace(" ", "")) == false)
                                                        {
                                                            Lista_pipeid.Add(pipeid2);
                                                            Lista_index1.Add(index1);
                                                        }
                                                    }

                                                    string heat2 = Convert.ToString(Wgen_main_form.dt_double_joint.Rows[i].GetChildRows(relation_double_joint1)[j][col3]);
                                                    if (heat2.Contains("/") == true)
                                                    {
                                                        string[] hid2 = heat2.Split(Convert.ToChar("/"));
                                                        for (int k = 0; k < hid2.Length; ++k)
                                                        {
                                                            if (heat1.ToLower().Replace(" ", "").Contains(hid2[k].ToLower().Replace(" ", "")) == false)
                                                            {
                                                                Lista_heat.Add(hid2[k]);
                                                                Lista_index1.Add(index1);
                                                            }
                                                        }
                                                    }
                                                    else
                                                    {
                                                        if (heat1.ToLower().Replace(" ", "").Contains(heat2.ToLower().Replace(" ", "")) == false)
                                                        {
                                                            Lista_heat.Add(heat2);
                                                            Lista_index2.Add(index1);
                                                        }
                                                    }

                                                    string wt2 = Convert.ToString(Wgen_main_form.dt_double_joint.Rows[i].GetChildRows(relation_double_joint1)[j][col6]);
                                                    if (wt1.ToLower().Replace(" ", "") != wt2.ToLower().Replace(" ", ""))
                                                    {
                                                        Lista_wt.Add(wt2);
                                                        Lista_index3.Add(index1);
                                                    }

                                                    string od2 = Convert.ToString(Wgen_main_form.dt_double_joint.Rows[i].GetChildRows(relation_double_joint1)[j][col7]);
                                                    if (od1.ToLower().Replace(" ", "") != od2.ToLower().Replace(" ", ""))
                                                    {
                                                        Lista_od.Add(od2);
                                                        Lista_index4.Add(index1);
                                                    }

                                                    string grade2 = Convert.ToString(Wgen_main_form.dt_double_joint.Rows[i].GetChildRows(relation_double_joint1)[j][col8]);
                                                    if (grade1.ToLower().Replace(" ", "") != grade2.ToLower().Replace(" ", ""))
                                                    {
                                                        Lista_grade.Add(grade2);
                                                        Lista_index5.Add(index1);
                                                    }

                                                    string coat2 = Convert.ToString(Wgen_main_form.dt_double_joint.Rows[i].GetChildRows(relation_double_joint1)[j][col9]);
                                                    if (coat1.ToLower().Replace(" ", "") != coat2.ToLower().Replace(" ", ""))
                                                    {
                                                        Lista_coat.Add(coat2);
                                                        Lista_index6.Add(index1);
                                                    }

                                                    string manuf2 = Convert.ToString(Wgen_main_form.dt_double_joint.Rows[i].GetChildRows(relation_double_joint1)[j][col10]);
                                                    if (manuf1.ToLower().Replace(" ", "") != manuf2.ToLower().Replace(" ", ""))
                                                    {
                                                        Lista_manuf.Add(manuf2);
                                                        Lista_index7.Add(index1);
                                                    }
                                                }

                                                if (Math.Abs(len1 - len2) > length_tolerance)
                                                {
                                                    string MMID1 = Convert.ToString(Wgen_main_form.dt_double_joint.Rows[i].GetChildRows(relation_double_joint1)[0][col1]);
                                                    dt_errors.Rows.Add();
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = MMID1;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = Convert.ToString(len1) + " vs " + Convert.ToString(len2) + " at " + dj1;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_11.Text + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_double_joint.Rows[i].GetChildRows(relation_double_joint1)[0]["rowno"]) + start_row);
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Double joint item: length missmatch between pipe manifest and pipe tally";
                                                }

                                                if (Lista_pipeid.Count > 0)
                                                {
                                                    for (int k = 0; k < Lista_pipeid.Count; ++k)
                                                    {
                                                        string MMID1 = Convert.ToString(Wgen_main_form.dt_double_joint.Rows[i].GetChildRows(relation_double_joint1)[k][col1]);
                                                        dt_errors.Rows.Add();
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = MMID1;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = Lista_pipeid[k] + " at " + dj1;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_2.Text + Convert.ToString(Lista_index1[k] + start_row);
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Double joint item: Pipe Id not matching the pipe manifest";
                                                    }
                                                }

                                                if (Lista_heat.Count > 0)
                                                {
                                                    for (int k = 0; k < Lista_heat.Count; ++k)
                                                    {
                                                        string MMID1 = Convert.ToString(Wgen_main_form.dt_double_joint.Rows[i].GetChildRows(relation_double_joint1)[k][col1]);
                                                        dt_errors.Rows.Add();
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = MMID1;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = Lista_heat[k] + " at " + dj1;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_3.Text + Convert.ToString(Lista_index2[k] + start_row);
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Double joint item: Heat not matching the pipe manifest";
                                                    }
                                                }

                                                if (Lista_wt.Count > 0)
                                                {
                                                    for (int k = 0; k < Lista_wt.Count; ++k)
                                                    {
                                                        string MMID1 = Convert.ToString(Wgen_main_form.dt_double_joint.Rows[i].GetChildRows(relation_double_joint1)[k][col1]);
                                                        dt_errors.Rows.Add();
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = MMID1;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = Lista_wt[k] + " at " + dj1;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_6.Text + Convert.ToString(Lista_index3[k] + start_row);
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Double joint item: Wall Thickness not matching the pipe manifest";
                                                    }
                                                }

                                                if (Lista_od.Count > 0)
                                                {
                                                    for (int k = 0; k < Lista_od.Count; ++k)
                                                    {
                                                        string MMID1 = Convert.ToString(Wgen_main_form.dt_double_joint.Rows[i].GetChildRows(relation_double_joint1)[k][col1]);
                                                        dt_errors.Rows.Add();
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = MMID1;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = Lista_od[k] + " at " + dj1;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_7.Text + Convert.ToString(Lista_index4[k] + start_row);
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Double joint item: Diameter not matching the pipe manifest";
                                                    }
                                                }

                                                if (Lista_grade.Count > 0)
                                                {
                                                    for (int k = 0; k < Lista_grade.Count; ++k)
                                                    {
                                                        string MMID1 = Convert.ToString(Wgen_main_form.dt_double_joint.Rows[i].GetChildRows(relation_double_joint1)[k][col1]);
                                                        dt_errors.Rows.Add();
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = MMID1;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = Lista_grade[k] + " at " + dj1;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_8.Text + Convert.ToString(Lista_index5[k] + start_row);
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Double joint item: Grade not matching the pipe manifest";
                                                    }
                                                }

                                                if (Lista_coat.Count > 0)
                                                {
                                                    for (int k = 0; k < Lista_coat.Count; ++k)
                                                    {
                                                        string MMID1 = Convert.ToString(Wgen_main_form.dt_double_joint.Rows[i].GetChildRows(relation_double_joint1)[k][col1]);
                                                        dt_errors.Rows.Add();
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = MMID1;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = Lista_coat[k] + " at " + dj1;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_9.Text + Convert.ToString(Lista_index6[k] + start_row);
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Double joint item: Coating not matching the pipe manifest";
                                                    }
                                                }


                                                if (Lista_manuf.Count > 0)
                                                {
                                                    for (int k = 0; k < Lista_manuf.Count; ++k)
                                                    {
                                                        string MMID1 = Convert.ToString(Wgen_main_form.dt_double_joint.Rows[i].GetChildRows(relation_double_joint1)[k][col1]);
                                                        dt_errors.Rows.Add();
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = MMID1;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = Lista_manuf[k] + " at " + dj1;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_10.Text + Convert.ToString(Lista_index7[k] + start_row);
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Double joint item: Manufacturer not matching the pipe manifest";
                                                    }
                                                }


                                            }
                                        }
                                    }


                                    for (int i = 0; i < Wgen_main_form.dt_ground_tally.Rows.Count; ++i)
                                    {
                                        if (Wgen_main_form.dt_ground_tally.Rows[i][col11] != DBNull.Value)
                                        {

                                            if (Wgen_main_form.dt_ground_tally.Rows[i].GetChildRows(relation_double_joint2).Length == 0)
                                            {
                                                string mmid2 = "xxx";
                                                if (Wgen_main_form.dt_ground_tally.Rows[i][col1] != DBNull.Value)
                                                {
                                                    mmid2 = Convert.ToString(Wgen_main_form.dt_ground_tally.Rows[i][col1]);
                                                }

                                                string dj2 = Convert.ToString(Wgen_main_form.dt_ground_tally.Rows[i][col11]);

                                                dt_errors.Rows.Add();
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = mmid2;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = dj2;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_11.Text + Convert.ToString(i + start_row);
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Double joint number not found inside pipe manifest";
                                            }
                                        }
                                    }

                                }
                                #endregion

                                dataset1.Relations.Remove(rel_duplic_pipeid);
                                dataset1.Relations.Remove(rel_duplic_heat);
                                dataset1.Relations.Remove(rel_duplic_mmid);
                                if (rel_pipe_id != null) dataset1.Relations.Remove(rel_pipe_id);
                                if (rel_heat != null) dataset1.Relations.Remove(rel_heat);
                                if (relation_dj != null) dataset1.Relations.Remove(relation_dj);
                                if (relation_wt != null) dataset1.Relations.Remove(relation_wt);
                                if (relation8 != null) dataset1.Relations.Remove(relation8);
                                if (relation_pipe_grade != null) dataset1.Relations.Remove(relation_pipe_grade);
                                if (relation10 != null) dataset1.Relations.Remove(relation10);
                                if (relation11 != null) dataset1.Relations.Remove(relation11);
                                if (relation_double_joint1 != null) dataset1.Relations.Remove(relation_double_joint1);
                                if (relation_double_joint2 != null) dataset1.Relations.Remove(relation_double_joint2);

                                dataset1.Tables.Remove(Wgen_main_form.dt_ground_tally);
                                dataset1.Tables.Remove(dt2);
                                dataset1.Tables.Remove(dt3);
                                if (Wgen_main_form.dt_pipe_list != null && Wgen_main_form.dt_pipe_list.Rows.Count > 0) dataset1.Tables.Remove(Wgen_main_form.dt_pipe_list);

                                dt2 = null;
                                dt3 = null;

                                if (Wgen_main_form.dt_double_joint != null && Wgen_main_form.dt_double_joint.Rows.Count > 0) dataset1.Tables.Remove(Wgen_main_form.dt_double_joint);

                                Wgen_main_form.dt_ground_tally.Columns.Remove("rowno");

                                if (dt_pipe_id_heat_duplicates.Rows.Count > 0)
                                {
                                    dt_pipe_id_heat_duplicates = Functions.Sort_data_table(dt_pipe_id_heat_duplicates, "mmid");

                                    #region dt_duplicates_for_pups deleting rows
                                    List<string> lista_pups = new List<string>();
                                    List<int> lista_index = new List<int>();
                                    for (int i = 0; i < dt_pipe_id_heat_duplicates.Rows.Count; ++i)
                                    {
                                        string id1 = Convert.ToString(dt_pipe_id_heat_duplicates.Rows[i][0]);

                                        if (id1.Contains(".") == true)
                                        {
                                            string parinte = id1.Substring(0, id1.IndexOf("."));
                                            if (lista_pups.Contains(parinte) == true)
                                            {
                                                if (lista_index.Contains(lista_pups.IndexOf(parinte)) == false) lista_index.Add(lista_pups.IndexOf(parinte));
                                                if (lista_index.Contains(i) == false) lista_index.Add(i);
                                            }
                                        }
                                        else
                                        {
                                            string last_letter = id1.Substring(id1.Length - 1, 1);
                                            if (Functions.IsNumeric(last_letter) == false)
                                            {
                                                string parinte = id1.Substring(0, id1.Length - 1);
                                                if (lista_pups.Contains(parinte) == true)
                                                {
                                                    if (lista_index.Contains(lista_pups.IndexOf(parinte)) == false) lista_index.Add(lista_pups.IndexOf(parinte));
                                                    if (lista_index.Contains(i) == false) lista_index.Add(i);
                                                }

                                            }
                                        }
                                        lista_pups.Add(id1);
                                    }

                                    if (lista_index.Count > 0)
                                    {
                                        for (int i = lista_index.Count - 1; i >= 0; --i)
                                        {
                                            dt_pipe_id_heat_duplicates.Rows[lista_index[i]].Delete();
                                        }
                                    }
                                    #endregion
                                }

                                if (dt_lengths.Rows.Count > 0)
                                {
                                    #region lengths
                                    dt_lengths = Functions.Sort_data_table(dt_lengths, "mmid");
                                    List<string> lista_pups = new List<string>();
                                    List<int> lista_index = new List<int>();
                                    for (int i = 0; i < dt_lengths.Rows.Count; ++i)
                                    {
                                        string id1 = Convert.ToString(dt_lengths.Rows[i][0]);

                                        if (id1.Contains(".") == true)
                                        {
                                            string parinte = id1.Substring(0, id1.IndexOf("."));
                                            if (lista_pups.Contains(parinte) == true)
                                            {
                                                if (lista_index.Contains(lista_pups.IndexOf(parinte)) == false) lista_index.Add(lista_pups.IndexOf(parinte));
                                                if (lista_index.Contains(i) == false) lista_index.Add(i);
                                            }
                                        }
                                        else
                                        {
                                            string last_letter = id1.Substring(id1.Length - 1, 1);
                                            if (Functions.IsNumeric(last_letter) == false)
                                            {
                                                string parinte = id1.Substring(0, id1.Length - 1);
                                                if (lista_pups.Contains(parinte) == true)
                                                {
                                                    if (lista_index.Contains(lista_pups.IndexOf(parinte)) == false) lista_index.Add(lista_pups.IndexOf(parinte));
                                                    if (lista_index.Contains(i) == false) lista_index.Add(i);
                                                }

                                            }
                                        }
                                        lista_pups.Add(id1);
                                    }

                                    if (lista_index.Count > 0)
                                    {
                                        for (int i = dt_lengths.Rows.Count - 1; i >= 0; --i)
                                        {
                                            if (lista_index.Contains(i) == false) dt_lengths.Rows[i].Delete();
                                        }
                                    }
                                    #endregion
                                }

                                for (int i = 0; i < dt_lengths.Rows.Count; ++i)
                                {
                                    #region length checks
                                    string MMID1 = Convert.ToString(dt_lengths.Rows[i][0]);
                                    string last_letter1 = MMID1.Substring(MMID1.Length - 1, 1);
                                    if (MMID1.Contains(".") == false && Functions.IsNumeric(last_letter1) == true)
                                    {
                                        if (dt_lengths.Rows[i][4] != DBNull.Value)
                                        {
                                            double new_parent_length = Convert.ToDouble(dt_lengths.Rows[i][5]);
                                            double original_length = Convert.ToDouble(dt_lengths.Rows[i][4]);
                                            string pipeid1 = Convert.ToString(dt_lengths.Rows[i][1]);
                                            string heat1 = Convert.ToString(dt_lengths.Rows[i][2]);
                                            double cumul_len_pups = 0;
                                            bool pup_found = false;
                                            for (int j = 0; j < dt_lengths.Rows.Count; ++j)
                                            {
                                                if (i != j)
                                                {
                                                    string MMID2 = Convert.ToString(dt_lengths.Rows[j][0]);
                                                    string str_new_len_pup = "0";

                                                    if (dt_lengths.Rows[j][5] != DBNull.Value)
                                                    {
                                                        str_new_len_pup = Convert.ToString(dt_lengths.Rows[j][5]);
                                                    }

                                                    string pipeid2 = Convert.ToString(dt_lengths.Rows[j][1]);
                                                    string heat2 = Convert.ToString(dt_lengths.Rows[j][2]);

                                                    if (MMID2.ToLower().Contains(MMID1.ToLower()) == true && MMID2.Contains(".") == true && pipeid1 == pipeid2 && heat1 == heat2)
                                                    {
                                                        pup_found = true;
                                                        if (Functions.IsNumeric(str_new_len_pup) == true)
                                                        {
                                                            double len_pup = Convert.ToDouble(str_new_len_pup);
                                                            cumul_len_pups = cumul_len_pups + len_pup;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        string last_letter = MMID2.Substring(MMID2.Length - 1, 1);
                                                        if (MMID2.ToLower().Contains(MMID1.ToLower()) == true && Functions.IsNumeric(last_letter) == false && pipeid1 == pipeid2 && heat1 == heat2)
                                                        {
                                                            pup_found = true;
                                                            if (Functions.IsNumeric(str_new_len_pup) == true)
                                                            {
                                                                double len_pup = Convert.ToDouble(str_new_len_pup);
                                                                cumul_len_pups = cumul_len_pups + len_pup;
                                                            }
                                                        }
                                                    }

                                                }
                                            }

                                            for (int j = 0; j < dt_lengths.Rows.Count; ++j)
                                            {
                                                if (i != j)
                                                {
                                                    string MMID2 = Convert.ToString(dt_lengths.Rows[j][0]);

                                                    string pipeid2 = Convert.ToString(dt_lengths.Rows[j][1]);
                                                    string heat2 = Convert.ToString(dt_lengths.Rows[j][2]);

                                                    if (MMID2.ToLower().Contains(MMID1.ToLower()) == true && MMID2.Contains(".") == true && pipeid1 == pipeid2 && heat1 == heat2)
                                                    {
                                                        pup_found = true;
                                                        dt_lengths.Rows[j][8] = Convert.ToString(original_length) + " vs " + Convert.ToString(cumul_len_pups + new_parent_length);
                                                    }
                                                    else
                                                    {
                                                        string last_letter = MMID2.Substring(MMID2.Length - 1, 1);
                                                        if (MMID2.ToLower().Contains(MMID1.ToLower()) == true && Functions.IsNumeric(last_letter) == false && pipeid1 == pipeid2 && heat1 == heat2)
                                                        {
                                                            pup_found = true;
                                                            dt_lengths.Rows[j][8] = Convert.ToString(original_length) + " vs " + Convert.ToString(cumul_len_pups + new_parent_length);
                                                        }
                                                    }
                                                }
                                            }

                                            double diference = Math.Round(Math.Round(cumul_len_pups, 2) + Math.Round(new_parent_length, 2) - Math.Round(original_length, 2), 2);

                                            if (Math.Abs(diference) > pup_length_tolerance)
                                            {
                                                if (pup_found == true)
                                                {
                                                    dt_errors.Rows.Add();
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = MMID1;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = Convert.ToString(original_length) + " vs " + Convert.ToString(cumul_len_pups + new_parent_length);
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_4.Text + Convert.ToString(Convert.ToInt32(dt_lengths.Rows[i][6]) + start_row);
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Pup length not matching original length";
                                                    ++nr_length_off;
                                                    dt_lengths.Rows[i][7] = false;
                                                }
                                                else
                                                {
                                                    dt_lengths.Rows[i][7] = true;
                                                }
                                            }
                                            else
                                            {
                                                dt_lengths.Rows[i][7] = true;
                                            }

                                            for (int j = 0; j < dt_lengths.Rows.Count; ++j)
                                            {
                                                if (i != j)
                                                {
                                                    string MMID2 = Convert.ToString(dt_lengths.Rows[j][0]);
                                                    string str_l2 = Convert.ToString(dt_lengths.Rows[j][5]);
                                                    string pipeid2 = Convert.ToString(dt_lengths.Rows[j][1]);
                                                    string heat2 = Convert.ToString(dt_lengths.Rows[j][2]);

                                                    if (MMID2.ToLower().Contains(MMID1.ToLower()) == true && MMID2.Contains(".") == true && pipeid1 == pipeid2 && heat1 == heat2)
                                                    {
                                                        dt_lengths.Rows[j][7] = dt_lengths.Rows[i][7];
                                                    }
                                                    else
                                                    {
                                                        string last_letter = MMID2.Substring(MMID2.Length - 1, 1);
                                                        if (MMID2.ToLower().Contains(MMID1.ToLower()) == true && Functions.IsNumeric(last_letter) == false && pipeid1 == pipeid2 && heat1 == heat2)
                                                        {
                                                            dt_lengths.Rows[j][7] = dt_lengths.Rows[i][7];
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    } //if (MMID1.Contains(".") == false) 
                                    #endregion
                                }

                                dataset1 = new DataSet();
                                if (dt_pipe_id_heat_duplicates != null) dt_pipe_id_heat_duplicates.TableName = "T11";
                                if (dt_lengths != null) dt_lengths.TableName = "T12";
                                dataset1.Tables.Add(dt_pipe_id_heat_duplicates);
                                dataset1.Tables.Add(dt_lengths);

                                DataRelation rel_len_mmid = new DataRelation("xxx", dt_pipe_id_heat_duplicates.Columns[0], dt_lengths.Columns[0], false);
                                dataset1.Relations.Add(rel_len_mmid);

                                for (int i = 0; i < dt_lengths.Rows.Count; ++i)
                                {
                                    string MMID1 = Convert.ToString(dt_lengths.Rows[i][0]);
                                    if (dt_lengths.Rows[i][6] == DBNull.Value)
                                    {
                                        dt_errors.Rows.Add();
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = MMID1;
                                        if (dt_lengths.Rows[i][8] != DBNull.Value) dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = Convert.ToString(dt_lengths.Rows[i][8]);
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_2.Text + Convert.ToString(Convert.ToInt32(dt_lengths.Rows[i][6]) + start_row);
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Length not matching the pipe manifest value";
                                        ++nr_duplicates_pipe;
                                    }
                                }

                                for (int i = 0; i < dt_pipe_id_heat_duplicates.Rows.Count; ++i)
                                {
                                    string MMID1 = Convert.ToString(dt_pipe_id_heat_duplicates.Rows[i][0]);
                                    if (dt_pipe_id_heat_duplicates.Rows[i].GetChildRows(rel_len_mmid).Length == 0)
                                    {
                                        dt_errors.Rows.Add();
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = MMID1;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = Convert.ToString(dt_pipe_id_heat_duplicates.Rows[i][1]) + " - " + Convert.ToString(dt_pipe_id_heat_duplicates.Rows[i][2]);
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_2.Text + Convert.ToString(Convert.ToInt32(dt_pipe_id_heat_duplicates.Rows[i][5]) + start_row);
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Pipe Id and Heat # Duplicate";
                                        ++nr_duplicates_pipe;
                                    }
                                }

                                dataset1.Relations.Remove(rel_len_mmid);
                                dataset1.Tables.Remove(dt_pipe_id_heat_duplicates);
                                dataset1.Tables.Remove(dt_lengths);


                                for (int i = 0; i < Wgen_main_form.dt_ground_tally.Rows.Count; ++i)
                                {
                                    string MMID1 = "";
                                    if (Wgen_main_form.dt_ground_tally.Rows[i][col1] != DBNull.Value)
                                    {
                                        MMID1 = Convert.ToString(Wgen_main_form.dt_ground_tally.Rows[i][col1]);
                                    }
                                    #region null values
                                    if (Wgen_main_form.dt_ground_tally.Rows[i][col1] == DBNull.Value)
                                    {
                                        dt_errors.Rows.Add();
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_1.Text + Convert.ToString(i + start_row);
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No MMID Specified";
                                        ++nr_null_values;
                                    }
                                    if (Wgen_main_form.dt_ground_tally.Rows[i][col2] == DBNull.Value)
                                    {
                                        dt_errors.Rows.Add();
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = "MMID: " + MMID1;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_2.Text + Convert.ToString(i + start_row);
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No Pipe ID Value Specified";
                                        ++nr_null_values;
                                    }
                                    if (Wgen_main_form.dt_ground_tally.Rows[i][col3] == DBNull.Value)
                                    {
                                        dt_errors.Rows.Add();
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = "MMID: " + MMID1;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_3.Text + Convert.ToString(i + start_row);
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No Heat Value Specified";
                                        ++nr_null_values;
                                    }

                                    if (Wgen_main_form.dt_ground_tally.Rows[i][col4] == DBNull.Value)
                                    {
                                        dt_errors.Rows.Add();
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = "MMID: " + MMID1;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_4.Text + Convert.ToString(i + start_row);
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No Original Value Specified";
                                        ++nr_null_values;
                                    }


                                    if (Wgen_main_form.dt_ground_tally.Rows[i][col6] == DBNull.Value)
                                    {
                                        dt_errors.Rows.Add();
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = "MMID: " + MMID1;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_6.Text + Convert.ToString(i + start_row);
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No Wall Thickness Specified";
                                        ++nr_null_values;
                                    }

                                    if (Wgen_main_form.dt_ground_tally.Rows[i][col7] == DBNull.Value)
                                    {
                                        dt_errors.Rows.Add();
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = "MMID: " + MMID1;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_7.Text + Convert.ToString(i + start_row);
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No Diameter Specified";
                                        ++nr_null_values;
                                    }

                                    if (Wgen_main_form.dt_ground_tally.Rows[i][col8] == DBNull.Value)
                                    {
                                        dt_errors.Rows.Add();
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = "MMID: " + MMID1;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_8.Text + Convert.ToString(i + start_row);
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No Grade Specified";
                                        ++nr_null_values;
                                    }

                                    if (Wgen_main_form.dt_ground_tally.Rows[i][col9] == DBNull.Value)
                                    {
                                        dt_errors.Rows.Add();
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = "MMID: " + MMID1;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_9.Text + Convert.ToString(i + start_row);
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No Coating Specified";
                                        ++nr_null_values;
                                    }

                                    if (Wgen_main_form.dt_ground_tally.Rows[i][col10] == DBNull.Value)
                                    {
                                        dt_errors.Rows.Add();
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = "MMID: " + MMID1;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_10.Text + Convert.ToString(i + start_row);
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No Manufacture Specified";
                                        ++nr_null_values;
                                    }
                                    #endregion
                                }

                                dt_errors = Functions.Sort_data_table(dt_errors, "Error");
                                transfer_errors_to_panel(dt_errors);
                                dt_export = Functions.creaza_error_export_table(dt_errors, sheet_name);

                                textBox_PT_no_rows.Text = Convert.ToString(Wgen_main_form.dt_ground_tally.Rows.Count);
                                textBox_PT_no_duplicates.Text = Convert.ToString(nr_duplicates_pipe);
                                textBox_PT_no_mmid_duplicates.Text = Convert.ToString(nr_duplicates_mmid);
                                textBox_PT_no_null.Text = Convert.ToString(nr_null_values);
                                textBox_PT_no_pipe_ID_not_found.Text = Convert.ToString(nr_pipe_id_missing);
                                textBox_PT_no_heat_not_found.Text = Convert.ToString(nr_heat_missing);
                                textBox_PT_no_length_matching.Text = Convert.ToString(nr_length_off);
                                textBox_PT_no_not_numeric.Text = Convert.ToString(nr_not_numeric);
                                textBox_PT_no_not_match.Text = Convert.ToString(nr_dj_missmatch);

                                button_pipe_tally_l.Visible = true;
                                button_pipe_tally_nl.Visible = false;
                            }
                            else
                            {
                                button_pipe_tally_l.Visible = false;
                                button_pipe_tally_nl.Visible = true;
                            }
                        }
                        set_enable_true();
                    }
                    else
                    {
                        button_pipe_tally_l.Visible = false;
                        button_pipe_tally_nl.Visible = true;
                    }
                }
                else
                {
                    button_pipe_tally_l.Visible = false;
                    button_pipe_tally_nl.Visible = true;
                }
            }
            else
            {
                button_pipe_tally_l.Visible = false;
                button_pipe_tally_nl.Visible = true;
            }
        }



        private void button_zoom_click(object ob, EventArgs e)
        {
            Control ctrl1 = ob as Control;
            if (dt_errors == null || dt_errors.Rows.Count == 0) return;
            if (ctrl1 != null)
            {
                int Y = ctrl1.Location.Y;
                int index1 = 0;
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
            dataGridView_error_pipe_tally.DataSource = null;
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




        private void button_export_errors_to_xl_Click(object sender, EventArgs e)
        {
            Functions.Transfer_datatable_to_new_excel_spreadsheet_named(dt_export, "PipeTallyErrors");
        }

        private void panel_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {

                ContextMenuStrip_load_build_pipe_tally.Show(Cursor.Position);
                ContextMenuStrip_load_build_pipe_tally.Visible = true;
            }
            else
            {
                ContextMenuStrip_load_build_pipe_tally.Visible = false;
            }
        }

        private void DataGridView_error_pipe_tally_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right && e.RowIndex >= 0)
            {
                dataGridView_error_pipe_tally.CurrentCell = dataGridView_error_pipe_tally.Rows[e.RowIndex].Cells[0];
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

            int index1 = dataGridView_error_pipe_tally.CurrentCell.RowIndex;
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



        public void radioButton_pipe_tally_CheckChanged(RadioButton radioButton_enlarged)
        {
            Font regularfont = new Font("Arial", 8.2f, FontStyle.Bold);


            Font englargedFont = new Font("Arial", 10f, FontStyle.Bold);


            Font regularHeader = new Font("Arial", 10f, FontStyle.Bold);

            Font englargedHeader = new Font("Arial", 12f, FontStyle.Bold);
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
                panel7.Location = new Point(3, 0);
                panel7.Size = new Size(723, 28);
            }
            else
            {
                panel7.Location = new Point(3, 3);
                panel7.Size = new Size(723, 25);
            }

            if (radioButton_enlarged.Checked == true)
            {
                panel_pipe_manifest.Location = new Point(3, 30);
                panel_pipe_manifest.Size = new Size(723, 35);
            }
            else
            {
                panel_pipe_manifest.Location = new Point(3, 27);
                panel_pipe_manifest.Size = new Size(723, 33);
            }

            if (radioButton_enlarged.Checked == true)
            {
                button_pipe_tally_l.Location = new Point(695, 3);
                button_pipe_tally_l.Size = new Size(24, 24);
            }
            else
            {
                button_pipe_tally_l.Location = new Point(697, 5);
                button_pipe_tally_l.Size = new Size(21, 21);
            }

            if (radioButton_enlarged.Checked == true)
            {
                button_pipe_tally_nl.Location = new Point(695, 3);
                button_pipe_tally_nl.Size = new Size(24, 24);
            }
            else
            {
                button_pipe_tally_nl.Location = new Point(697, 5);
                button_pipe_tally_nl.Size = new Size(21, 21);
            }

            if (radioButton_enlarged.Checked == true)
            {
                button_refresh_ws1.Location = new Point(3, 3);
                button_refresh_ws1.Size = new Size(24, 24);
            }
            else
            {
                button_refresh_ws1.Location = new Point(5, 5);
                button_refresh_ws1.Size = new Size(21, 21);
            }

            if (radioButton_enlarged.Checked == true)
            {
                button_load_pipe_tally.Location = new Point(494, 2);
                button_load_pipe_tally.Size = new Size(198, 30);

                button_load_pipe_tally.Font = englargedFont;
            }
            else
            {
                button_load_pipe_tally.Location = new Point(546, 2);
                button_load_pipe_tally.Size = new Size(145, 28);

                button_load_pipe_tally.Font = regularfont;
            }


            if (radioButton_enlarged.Checked == true)
            {
                comboBox_ws1.Location = new Point(32, 4);
                comboBox_ws1.Size = new Size(455, 25);

                comboBox_ws1.Font = englargedFont;
            }
            else
            {
                comboBox_ws1.Location = new Point(32, 4);
                comboBox_ws1.Size = new Size(508, 24);

                comboBox_ws1.Font = regularfont;
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
                panel6.Location = new Point(2, 69);
                panel6.Size = new Size(722, 29);
            }
            else
            {
                panel6.Location = new Point(2, 64);
                panel6.Size = new Size(722, 29);
            }

            if (radioButton_enlarged.Checked == true)
            {
                button_show_feature_codes.Location = new Point(525, -1);
                button_show_feature_codes.Size = new Size(198, 30);

                button_show_feature_codes.Font = englargedFont;
            }
            else
            {
                button_show_feature_codes.Location = new Point(548, -1);
                button_show_feature_codes.Size = new Size(173, 28);

                button_show_feature_codes.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                dataGridView_error_pipe_tally.Location = new Point(2, 98);
                dataGridView_error_pipe_tally.Size = new Size(723, 395);

                dataGridView_error_pipe_tally.DefaultCellStyle.Font = englargedFont;

                dataGridView_error_pipe_tally.ColumnHeadersDefaultCellStyle.Font = englargedHeader;

            }
            else
            {
                dataGridView_error_pipe_tally.Location = new Point(2, 91);
                dataGridView_error_pipe_tally.Size = new Size(723, 425);

                dataGridView_error_pipe_tally.DefaultCellStyle.Font = regularfont;

                dataGridView_error_pipe_tally.ColumnHeadersDefaultCellStyle.Font = regularHeader;

            }

            if (radioButton_enlarged.Checked == true)
            {
                panel_stats.Location = new Point(3, 495);
                panel_stats.Size = new Size(722, 178);
            }
            else
            {
                panel_stats.Location = new Point(3, 520);
                panel_stats.Size = new Size(722, 150);
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
                textBox_PM_Items.Size = new Size(305, 25);

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
                textBox_PT_no_rows.Location = new Point(315, 21);
                textBox_PT_no_rows.Size = new Size(40, 25);

                textBox_PT_no_rows.Font = englargedFont;
            }
            else
            {
                textBox_PT_no_rows.Location = new Point(311, 21);
                textBox_PT_no_rows.Size = new Size(37, 20);

                textBox_PT_no_rows.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_PM_missing_OD.Location = new Point(3, 47);
                textBox_PM_missing_OD.Size = new Size(305, 25);

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
                textBox_PT_no_duplicates.Location = new Point(315, 47);
                textBox_PT_no_duplicates.Size = new Size(40, 25);

                textBox_PT_no_duplicates.Font = englargedFont;
            }
            else
            {
                textBox_PT_no_duplicates.Location = new Point(311, 42);
                textBox_PT_no_duplicates.Size = new Size(37, 20);

                textBox_PT_no_duplicates.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_mmid_duplicates.Location = new Point(3, 73);
                textBox_mmid_duplicates.Size = new Size(305, 25);

                textBox_mmid_duplicates.Font = englargedFont;
            }
            else
            {
                textBox_mmid_duplicates.Location = new Point(3, 63);
                textBox_mmid_duplicates.Size = new Size(287, 20);

                textBox_mmid_duplicates.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_PT_no_mmid_duplicates.Location = new Point(315, 73);
                textBox_PT_no_mmid_duplicates.Size = new Size(40, 25);

                textBox_PT_no_mmid_duplicates.Font = englargedFont;
            }
            else
            {
                textBox_PT_no_mmid_duplicates.Location = new Point(311, 63);
                textBox_PT_no_mmid_duplicates.Size = new Size(37, 20);

                textBox_PT_no_mmid_duplicates.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_pipe_id_not_found.Location = new Point(3, 99);
                textBox_pipe_id_not_found.Size = new Size(305, 25);

                textBox_pipe_id_not_found.Font = englargedFont;
            }
            else
            {
                textBox_pipe_id_not_found.Location = new Point(3, 84);
                textBox_pipe_id_not_found.Size = new Size(287, 20);

                textBox_pipe_id_not_found.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_PT_no_pipe_ID_not_found.Location = new Point(315, 99);
                textBox_PT_no_pipe_ID_not_found.Size = new Size(40, 25);

                textBox_PT_no_pipe_ID_not_found.Font = englargedFont;
            }
            else
            {
                textBox_PT_no_pipe_ID_not_found.Location = new Point(311, 84);
                textBox_PT_no_pipe_ID_not_found.Size = new Size(37, 20);

                textBox_PT_no_pipe_ID_not_found.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_heat_not_found.Location = new Point(3, 125);
                textBox_heat_not_found.Size = new Size(305, 25);

                textBox_heat_not_found.Font = englargedFont;
            }
            else
            {
                textBox_heat_not_found.Location = new Point(3, 105);
                textBox_heat_not_found.Size = new Size(287, 20);

                textBox_heat_not_found.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_PT_no_heat_not_found.Location = new Point(315, 125);
                textBox_PT_no_heat_not_found.Size = new Size(40, 25);

                textBox_PT_no_heat_not_found.Font = englargedFont;
            }
            else
            {
                textBox_PT_no_heat_not_found.Location = new Point(311, 105);
                textBox_PT_no_heat_not_found.Size = new Size(37, 20);

                textBox_PT_no_heat_not_found.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_PM_null_value_items.Location = new Point(3, 151);
                textBox_PM_null_value_items.Size = new Size(305, 25);

                textBox_PM_null_value_items.Font = englargedFont;
            }
            else
            {
                textBox_PM_null_value_items.Location = new Point(3, 126);
                textBox_PM_null_value_items.Size = new Size(287, 20);

                textBox_PM_null_value_items.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_PT_no_null.Location = new Point(315, 151);
                textBox_PT_no_null.Size = new Size(40, 25);

                textBox_PT_no_null.Font = englargedFont;
            }
            else
            {
                textBox_PT_no_null.Location = new Point(311, 126);
                textBox_PT_no_null.Size = new Size(37, 20);

                textBox_PT_no_null.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_length_off_items.Location = new Point(360, 21);
                textBox_length_off_items.Size = new Size(305, 25);

                textBox_length_off_items.Font = englargedFont;
            }
            else
            {
                textBox_length_off_items.Location = new Point(362, 21);
                textBox_length_off_items.Size = new Size(287, 20);

                textBox_length_off_items.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_PT_no_length_matching.Location = new Point(670, 21);
                textBox_PT_no_length_matching.Size = new Size(40, 25);

                textBox_PT_no_length_matching.Font = englargedFont;
            }
            else
            {
                textBox_PT_no_length_matching.Location = new Point(662, 21);
                textBox_PT_no_length_matching.Size = new Size(37, 20);

                textBox_PT_no_length_matching.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_not_numeric.Location = new Point(360, 47);
                textBox_not_numeric.Size = new Size(305, 25);

                textBox_not_numeric.Font = englargedFont;
            }
            else
            {
                textBox_not_numeric.Location = new Point(362, 41);
                textBox_not_numeric.Size = new Size(287, 20);

                textBox_not_numeric.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_PT_no_not_numeric.Location = new Point(670, 47);
                textBox_PT_no_not_numeric.Size = new Size(40, 25);

                textBox_PT_no_not_numeric.Font = englargedFont;
            }
            else
            {
                textBox_PT_no_not_numeric.Location = new Point(662, 41);
                textBox_PT_no_not_numeric.Size = new Size(37, 20);

                textBox_PT_no_not_numeric.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_not_match_pipe.Location = new Point(360, 73);
                textBox_not_match_pipe.Size = new Size(305, 25);

                textBox_not_match_pipe.Font = englargedFont;
            }
            else
            {
                textBox_not_match_pipe.Location = new Point(362, 61);
                textBox_not_match_pipe.Size = new Size(287, 20);

                textBox_not_match_pipe.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_PT_no_not_match.Location = new Point(670, 73);
                textBox_PT_no_not_match.Size = new Size(40, 25);

                textBox_PT_no_not_match.Font = englargedFont;
            }
            else
            {
                textBox_PT_no_not_match.Location = new Point(662, 61);
                textBox_PT_no_not_match.Size = new Size(37, 20);

                textBox_PT_no_not_match.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                button_export_errors_to_xl.Location = new Point(505, 145);
                button_export_errors_to_xl.Size = new Size(214, 30);

                button_export_errors_to_xl.Font = englargedFont;
            }
            else
            {
                button_export_errors_to_xl.Location = new Point(557, 117);
                button_export_errors_to_xl.Size = new Size(161, 28);

                button_export_errors_to_xl.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                button_dismiss_errors.Location = new Point(382, 0);
                button_dismiss_errors.Size = new Size(124, 30);

                button_dismiss_errors.Font = englargedFont;
            }
            else
            {
                button_dismiss_errors.Location = new Point(433, 0);
                button_dismiss_errors.Size = new Size(104, 25);

                button_dismiss_errors.Font = regularfont;
            }

        }


        private void transfer_errors_to_panel(System.Data.DataTable dt1)
        {
            if (dt1.Rows.Count > 0)
            {
                System.Data.DataTable dt_display = dt1.Copy();
                dt_display.Columns.RemoveAt(3);

                if (dt_dismissed_errors != null && dt_dismissed_errors.Rows.Count > 0)
                {
                    dt_display.TableName = "err1";
                    dt_dismissed_errors.TableName = "t5";
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



                dataGridView_error_pipe_tally.DataSource = dt_display;
                dataGridView_error_pipe_tally.Columns[0].Width = 75;
                dataGridView_error_pipe_tally.Columns[1].Width = 300;
                dataGridView_error_pipe_tally.Columns[2].Width = 50;
                dataGridView_error_pipe_tally.Columns[3].Width = 300;
                dataGridView_error_pipe_tally.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_error_pipe_tally.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                dataGridView_error_pipe_tally.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_error_pipe_tally.DefaultCellStyle.ForeColor = Color.White;
                dataGridView_error_pipe_tally.EnableHeadersVisualStyles = false;
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

                foreach (DataGridViewCell cell1 in dataGridView_error_pipe_tally.SelectedCells)
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
                        dt_dismissed_errors.Rows[dt_dismissed_errors.Rows.Count - 1][0] = dataGridView_error_pipe_tally.Rows[lista1[i]].Cells[0].Value;
                        dt_dismissed_errors.Rows[dt_dismissed_errors.Rows.Count - 1][1] = dataGridView_error_pipe_tally.Rows[lista1[i]].Cells[1].Value;
                        dt_dismissed_errors.Rows[dt_dismissed_errors.Rows.Count - 1][2] = dataGridView_error_pipe_tally.Rows[lista1[i]].Cells[3].Value;
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
