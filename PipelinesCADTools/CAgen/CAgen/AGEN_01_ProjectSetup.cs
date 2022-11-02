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
    public partial class AGEN_Project_Setup : Form
    {

        _AGEN_mainform Ag = null;
        public bool is_loading = true;

        string sta_back = "Station Back";
        string sta_ahead = "Station Ahead";
        string rr_end_x = "Reroute End X";
        string rr_end_y = "Reroute End Y";
        string rr_end_z = "Reroute End Z";
        string version = "Version";
        string show_in_plan = "Show in plan";
        string Col_x = "X";
        string Col_y = "Y";
        string Col_z = "Z";
        string Col_3DSta = "3DSta";
        string Col_BackSta = "BackSta";
        string Col_AheadSta = "AheadSta";

        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();


            lista_butoane.Add(button_align_config_saveall);

            lista_butoane.Add(button_browse_select_output_folder);
            lista_butoane.Add(button_draw_stationing);
            lista_butoane.Add(button_eq_insert_block);


            lista_butoane.Add(button_insert_pi_blocks);
            lista_butoane.Add(button_load_config);

            lista_butoane.Add(button_load_new_CL_to_excel);
            lista_butoane.Add(button_mpkp_insert);

            lista_butoane.Add(button_set_project_folder);
            lista_butoane.Add(button_show_segment_list);

            lista_butoane.Add(comboBox_eq_as);
            lista_butoane.Add(comboBox_eq_block);
            lista_butoane.Add(comboBox_eq_bs);
            lista_butoane.Add(comboBox_eq_diff);
            lista_butoane.Add(comboBox_kpmp_units_precision);
            lista_butoane.Add(comboBox_mpkp_attribute);
            lista_butoane.Add(comboBox_mpkp_blocks);
            lista_butoane.Add(comboBox_pi_atr_defl);
            lista_butoane.Add(comboBox_pi_atr_sta);
            lista_butoane.Add(comboBox_pi_block);
            lista_butoane.Add(comboBox_segment_name);
            lista_butoane.Add(comboBox_text_styles);
            lista_butoane.Add(radioButton_canada);
            lista_butoane.Add(radioButton_Load_config);
            lista_butoane.Add(radioButton_new_config);
            lista_butoane.Add(radioButton_usa);

            lista_butoane.Add(textBox_client_name);
            lista_butoane.Add(textBox_project_name);

            lista_butoane.Add(textBox_kpmp_spacing);

            lista_butoane.Add(textBox_output_folder);
            lista_butoane.Add(textBox_pi_min_angle);
            lista_butoane.Add(textbox_project_database_folder);

            lista_butoane.Add(textBox_spacing_major);
            lista_butoane.Add(textBox_spacing_minor);
            lista_butoane.Add(textBox_start_station_CL);
            lista_butoane.Add(textBox_tic_major);

            lista_butoane.Add(textBox_tic_minor);
            lista_butoane.Add(button_display_tpage_load_cl_xl);




            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {

                bt1.Enabled = false;

            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();

            lista_butoane.Add(button_align_config_saveall);

            lista_butoane.Add(button_browse_select_output_folder);
            lista_butoane.Add(button_draw_stationing);
            lista_butoane.Add(button_eq_insert_block);

            lista_butoane.Add(button_insert_pi_blocks);
            lista_butoane.Add(button_load_config);

            lista_butoane.Add(button_load_new_CL_to_excel);
            lista_butoane.Add(button_mpkp_insert);

            lista_butoane.Add(button_set_project_folder);
            lista_butoane.Add(button_show_segment_list);

            lista_butoane.Add(comboBox_eq_as);
            lista_butoane.Add(comboBox_eq_block);
            lista_butoane.Add(comboBox_eq_bs);
            lista_butoane.Add(comboBox_eq_diff);
            lista_butoane.Add(comboBox_kpmp_units_precision);
            lista_butoane.Add(comboBox_mpkp_attribute);
            lista_butoane.Add(comboBox_mpkp_blocks);
            lista_butoane.Add(comboBox_pi_atr_defl);
            lista_butoane.Add(comboBox_pi_atr_sta);
            lista_butoane.Add(comboBox_pi_block);
            lista_butoane.Add(comboBox_segment_name);
            lista_butoane.Add(comboBox_text_styles);
            lista_butoane.Add(radioButton_canada);
            lista_butoane.Add(radioButton_Load_config);
            lista_butoane.Add(radioButton_new_config);
            lista_butoane.Add(radioButton_usa);


            lista_butoane.Add(textBox_client_name);
            lista_butoane.Add(textBox_project_name);

            lista_butoane.Add(textBox_kpmp_spacing);

            lista_butoane.Add(textBox_output_folder);
            lista_butoane.Add(textBox_pi_min_angle);
            lista_butoane.Add(textbox_project_database_folder);

            lista_butoane.Add(textBox_spacing_major);
            lista_butoane.Add(textBox_spacing_minor);
            lista_butoane.Add(textBox_start_station_CL);
            lista_butoane.Add(textBox_tic_major);

            lista_butoane.Add(button_display_tpage_load_cl_xl);
            lista_butoane.Add(textBox_tic_minor);
            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }



        public AGEN_Project_Setup()
        {
            InitializeComponent();

            button_load_config.Visible = false;


            _AGEN_mainform.config_path = "";
        }

        public void set_radioButton_canada(bool val)
        {
            radioButton_canada.Checked = val;
        }
        public void set_radioButton_usa(bool val)
        {
            radioButton_usa.Checked = val;
        }

        private void button_load_config_Click(object sender, EventArgs e)
        {
            Ag = this.MdiParent as _AGEN_mainform;
            using (OpenFileDialog fbd = new OpenFileDialog())
            {
                fbd.Multiselect = false;
                fbd.Filter = "Excel files (*.xlsx)|*.xlsx";

                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    is_loading = true;
                    _AGEN_mainform.tpage_processing.Show();
                    // Ag.WindowState = FormWindowState.Minimized;

                    _AGEN_mainform.lista_segments = null;
                    _AGEN_mainform.dt_settings_custom = null;
                    string File1 = fbd.FileName;

                    _AGEN_mainform.config_path = File1;

                    _AGEN_mainform.dt_layer_alias = null;
                    _AGEN_mainform.tpage_layer_alias.Set_layer_alias_label_to_red();

                    _AGEN_mainform.tpage_cust_draw.clear_combobox_custom();
                    _AGEN_mainform.tpage_cust_scan.clear_combobox_custom();
                    _AGEN_mainform.tpage_owner_scan.clear_combobox();
                    _AGEN_mainform.tpage_owner_draw.clear_combobox();
                    _AGEN_mainform.tpage_tblk_attrib.label_excel_to_red();

                    _AGEN_mainform.tpage_crossing_draw.clear_combobox();
                    _AGEN_mainform.ProjFolder = "";

                    #region Load_config_method
                    {
                        set_enable_false();
                        Load_existing_config_file(File1);
                        set_enable_true();
                    }
                    #endregion


                    _AGEN_mainform.tpage_sheetindex.Hide_labels_at_load_project();
                    _AGEN_mainform.tpage_processing.Hide();

                    if (_AGEN_mainform.COUNTRY == "CANADA")
                    {
                        radioButton_canada.Checked = true;
                        _AGEN_mainform.layer_centerline = "PNEW";
                        _AGEN_mainform.lw_cl = LineWeight.LineWeight050;
                        _AGEN_mainform.color_index_cl = 7;

                    }
                    else
                    {
                        radioButton_usa.Checked = true;
                        _AGEN_mainform.layer_centerline = "P_PL_CL";
                        _AGEN_mainform.lw_cl = LineWeight.ByLineWeightDefault;
                        _AGEN_mainform.color_index_cl = 1;


                    }
                    Ag.WindowState = FormWindowState.Normal;
                    is_loading = false;
                }
            }
        }
        private void Load_existing_config_file(string File1)
        {
            bool exista_extraVP = false;
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

                if (Excel1.Workbooks.Count == 0) Excel1.Visible = _AGEN_mainform.ExcelVisible;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(File1);

                string segment1 = "";




                try
                {
                    int no_worksheets = Workbook1.Worksheets.Count;
                    foreach (Microsoft.Office.Interop.Excel.Worksheet W1 in Workbook1.Worksheets)
                    {
                        try
                        {
                            #region main_worksheet

                            if (W1.Name == "main_cfg")
                            {

                                string b1 = Convert.ToString(W1.Range["B1"].Value2);
                                string b2 = Convert.ToString(W1.Range["B2"].Value);
                                string b3 = Convert.ToString(W1.Range["B3"].Value);
                                string b4 = Convert.ToString(W1.Range["B4"].Value);
                                string b5 = Convert.ToString(W1.Range["B5"].Value);
                                string b6 = Convert.ToString(W1.Range["B6"].Value);
                                string b7 = Convert.ToString(W1.Range["B7"].Value);
                                string b8 = Convert.ToString(W1.Range["B8"].Value);
                                string b9 = Convert.ToString(W1.Range["B9"].Value);
                                string b10 = Convert.ToString(W1.Range["B10"].Value);
                                string b11 = Convert.ToString(W1.Range["B11"].Value);
                                string b12 = Convert.ToString(W1.Range["B12"].Value);
                                string b13 = Convert.ToString(W1.Range["B13"].Value);
                                string b14 = Convert.ToString(W1.Range["B14"].Value);
                                string b15 = Convert.ToString(W1.Range["B15"].Value);
                                string b16 = Convert.ToString(W1.Range["B16"].Value);
                                string b17 = Convert.ToString(W1.Range["B17"].Value);
                                string b18 = Convert.ToString(W1.Range["B18"].Value);
                                string b19 = Convert.ToString(W1.Range["B19"].Value);
                                string b20 = Convert.ToString(W1.Range["B20"].Value);
                                string b21 = Convert.ToString(W1.Range["B21"].Value);
                                string b22 = Convert.ToString(W1.Range["B22"].Value);
                                string b23 = Convert.ToString(W1.Range["B23"].Value);
                                string b24 = Convert.ToString(W1.Range["B24"].Value);
                                string b25 = Convert.ToString(W1.Range["B25"].Value);
                                string b26 = Convert.ToString(W1.Range["B26"].Value);
                                string b27 = Convert.ToString(W1.Range["B27"].Value);
                                string b28 = Convert.ToString(W1.Range["B28"].Value);
                                string b29 = Convert.ToString(W1.Range["B29"].Value);
                                string b30 = Convert.ToString(W1.Range["B30"].Value);
                                string b31 = Convert.ToString(W1.Range["B31"].Value);
                                string b32 = Convert.ToString(W1.Range["B32"].Value);
                                string b33 = Convert.ToString(W1.Range["B33"].Value);
                                string b34 = Convert.ToString(W1.Range["B34"].Value);
                                string b35 = Convert.ToString(W1.Range["B35"].Value);
                                string b36 = Convert.ToString(W1.Range["B36"].Value);
                                string b37 = Convert.ToString(W1.Range["B37"].Value);
                                string b38 = Convert.ToString(W1.Range["B38"].Value);
                                string b39 = Convert.ToString(W1.Range["B39"].Value);
                                string b40 = Convert.ToString(W1.Range["B40"].Value);
                                string b41 = Convert.ToString(W1.Range["B41"].Value);
                                string b42 = Convert.ToString(W1.Range["B42"].Value);
                                string b43 = Convert.ToString(W1.Range["B43"].Value);
                                string b44 = Convert.ToString(W1.Range["B44"].Value);
                                string b45 = Convert.ToString(W1.Range["B45"].Value);
                                string b46 = Convert.ToString(W1.Range["B46"].Value);

                                string b47 = Convert.ToString(W1.Range["B47"].Value);
                                if (b47 != null && b47.Replace(" ", "") != "")
                                {
                                    _AGEN_mainform.lista_segments = new List<string>();
                                    _AGEN_mainform.lista_segments.Add(b47);
                                    segment1 = b47;
                                    if (segment1.Replace(" ", "") != "")
                                    {
                                        if (_AGEN_mainform.current_segment.ToLower() != segment1.ToLower())
                                        {
                                            _AGEN_mainform.current_segment = segment1;
                                        }
                                    }

                                    string b48 = Convert.ToString(W1.Range["B48"].Value);
                                    if (b48 != null && b48.Replace(" ", "") != "")
                                    {
                                        _AGEN_mainform.lista_segments.Add(b48);
                                        string b49 = Convert.ToString(W1.Range["B49"].Value);
                                        if (b49 != null && b49.Replace(" ", "") != "")
                                        {
                                            _AGEN_mainform.lista_segments.Add(b49);
                                            string b50 = Convert.ToString(W1.Range["B50"].Value);
                                            if (b50 != null && b50.Replace(" ", "") != "")
                                            {
                                                _AGEN_mainform.lista_segments.Add(b50);
                                                string b51 = Convert.ToString(W1.Range["B51"].Value);
                                                if (b51 != null && b51.Replace(" ", "") != "")
                                                {
                                                    _AGEN_mainform.lista_segments.Add(b51);
                                                    string b52 = Convert.ToString(W1.Range["B52"].Value);
                                                    if (b52 != null && b52.Replace(" ", "") != "")
                                                    {
                                                        _AGEN_mainform.lista_segments.Add(b52);
                                                        string b53 = Convert.ToString(W1.Range["B53"].Value);
                                                        if (b53 != null && b53.Replace(" ", "") != "")
                                                        {
                                                            _AGEN_mainform.lista_segments.Add(b53);
                                                            string b54 = Convert.ToString(W1.Range["B54"].Value);
                                                            if (b54 != null && b54.Replace(" ", "") != "")
                                                            {
                                                                _AGEN_mainform.lista_segments.Add(b54);
                                                                string b55 = Convert.ToString(W1.Range["B55"].Value);
                                                                if (b55 != null && b55.Replace(" ", "") != "")
                                                                {
                                                                    _AGEN_mainform.lista_segments.Add(b55);
                                                                    string b56 = Convert.ToString(W1.Range["B56"].Value);
                                                                    if (b56 != null && b56.Replace(" ", "") != "")
                                                                    {
                                                                        _AGEN_mainform.lista_segments.Add(b56);
                                                                        string b57 = Convert.ToString(W1.Range["B57"].Value);
                                                                        if (b57 != null && b57.Replace(" ", "") != "")
                                                                        {
                                                                            _AGEN_mainform.lista_segments.Add(b57);
                                                                            string b58 = Convert.ToString(W1.Range["B58"].Value);
                                                                            if (b58 != null && b58.Replace(" ", "") != "")
                                                                            {
                                                                                _AGEN_mainform.lista_segments.Add(b58);
                                                                                string b59 = Convert.ToString(W1.Range["B59"].Value);
                                                                                if (b59 != null && b59.Replace(" ", "") != "")
                                                                                {
                                                                                    _AGEN_mainform.lista_segments.Add(b59);
                                                                                    string b60 = Convert.ToString(W1.Range["B60"].Value);
                                                                                    if (b60 != null && b60.Replace(" ", "") != "")
                                                                                    {
                                                                                        _AGEN_mainform.lista_segments.Add(b60);
                                                                                        string b61 = Convert.ToString(W1.Range["B61"].Value);
                                                                                        if (b61 != null && b61.Replace(" ", "") != "")
                                                                                        {
                                                                                            _AGEN_mainform.lista_segments.Add(b61);
                                                                                            string b62 = Convert.ToString(W1.Range["B62"].Value);
                                                                                            if (b62 != null && b62.Replace(" ", "") != "")
                                                                                            {
                                                                                                _AGEN_mainform.lista_segments.Add(b62);
                                                                                                string b63 = Convert.ToString(W1.Range["B63"].Value);
                                                                                                if (b63 != null && b63.Replace(" ", "") != "")
                                                                                                {
                                                                                                    _AGEN_mainform.lista_segments.Add(b63);
                                                                                                    string b64 = Convert.ToString(W1.Range["B64"].Value);
                                                                                                    if (b64 != null)
                                                                                                    {
                                                                                                        _AGEN_mainform.lista_segments.Add(b64);
                                                                                                        string b65 = Convert.ToString(W1.Range["B65"].Value);
                                                                                                        if (b65 != null && b65.Replace(" ", "") != "")
                                                                                                        {
                                                                                                            _AGEN_mainform.lista_segments.Add(b65);
                                                                                                            string b66 = Convert.ToString(W1.Range["B66"].Value);
                                                                                                            if (b66 != null && b66.Replace(" ", "") != "")
                                                                                                            {
                                                                                                                _AGEN_mainform.lista_segments.Add(b66);
                                                                                                                string b67 = Convert.ToString(W1.Range["B67"].Value);
                                                                                                                if (b67 != null && b67.Replace(" ", "") != "")
                                                                                                                {
                                                                                                                    _AGEN_mainform.lista_segments.Add(b67);
                                                                                                                    string b68 = Convert.ToString(W1.Range["B68"].Value);
                                                                                                                    if (b68 != null && b68.Replace(" ", "") != "")
                                                                                                                    {
                                                                                                                        _AGEN_mainform.lista_segments.Add(b68);
                                                                                                                        string b69 = Convert.ToString(W1.Range["B69"].Value);
                                                                                                                        if (b69 != null && b69.Replace(" ", "") != "")
                                                                                                                        {
                                                                                                                            _AGEN_mainform.lista_segments.Add(b69);
                                                                                                                            string b70 = Convert.ToString(W1.Range["B70"].Value);
                                                                                                                            if (b70 != null && b70.Replace(" ", "") != "")
                                                                                                                            {
                                                                                                                                _AGEN_mainform.lista_segments.Add(b70);
                                                                                                                                string b71 = Convert.ToString(W1.Range["B71"].Value);
                                                                                                                                if (b71 != null && b71.Replace(" ", "") != "")
                                                                                                                                {
                                                                                                                                    _AGEN_mainform.lista_segments.Add(b71);
                                                                                                                                    string b72 = Convert.ToString(W1.Range["B72"].Value);
                                                                                                                                    if (b72 != null && b72.Replace(" ", "") != "")
                                                                                                                                    {
                                                                                                                                        _AGEN_mainform.lista_segments.Add(b72);
                                                                                                                                        string b73 = Convert.ToString(W1.Range["B73"].Value);
                                                                                                                                        if (b73 != null && b73.Replace(" ", "") != "")
                                                                                                                                        {
                                                                                                                                            _AGEN_mainform.lista_segments.Add(b73);
                                                                                                                                            string b74 = Convert.ToString(W1.Range["B74"].Value);
                                                                                                                                            if (b74 != null && b74.Replace(" ", "") != "")
                                                                                                                                            {
                                                                                                                                                _AGEN_mainform.lista_segments.Add(b74);
                                                                                                                                                string b75 = Convert.ToString(W1.Range["B75"].Value);
                                                                                                                                                if (b75 != null && b75.Replace(" ", "") != "")
                                                                                                                                                {
                                                                                                                                                    _AGEN_mainform.lista_segments.Add(b75);
                                                                                                                                                    string b76 = Convert.ToString(W1.Range["B76"].Value);
                                                                                                                                                    if (b76 != null && b76.Replace(" ", "") != "")
                                                                                                                                                    {
                                                                                                                                                        _AGEN_mainform.lista_segments.Add(b76);
                                                                                                                                                        string b77 = Convert.ToString(W1.Range["B77"].Value);
                                                                                                                                                        if (b77 != null && b77.Replace(" ", "") != "")
                                                                                                                                                        {
                                                                                                                                                            _AGEN_mainform.lista_segments.Add(b77);
                                                                                                                                                            string b78 = Convert.ToString(W1.Range["B78"].Value);
                                                                                                                                                            if (b78 != null && b78.Replace(" ", "") != "")
                                                                                                                                                            {
                                                                                                                                                                _AGEN_mainform.lista_segments.Add(b78);
                                                                                                                                                                string b79 = Convert.ToString(W1.Range["B79"].Value);
                                                                                                                                                                if (b79 != null && b79.Replace(" ", "") != "")
                                                                                                                                                                {
                                                                                                                                                                    _AGEN_mainform.lista_segments.Add(b79);
                                                                                                                                                                    string b80 = Convert.ToString(W1.Range["B80"].Value);
                                                                                                                                                                    if (b80 != null && b80.Replace(" ", "") != "")
                                                                                                                                                                    {
                                                                                                                                                                        _AGEN_mainform.lista_segments.Add(b80);
                                                                                                                                                                        string b81 = Convert.ToString(W1.Range["B81"].Value);
                                                                                                                                                                        if (b81 != null && b81.Replace(" ", "") != "")
                                                                                                                                                                        {
                                                                                                                                                                            _AGEN_mainform.lista_segments.Add(b81);
                                                                                                                                                                            string b82 = Convert.ToString(W1.Range["B82"].Value);
                                                                                                                                                                            if (b82 != null && b82.Replace(" ", "") != "")
                                                                                                                                                                            {
                                                                                                                                                                                _AGEN_mainform.lista_segments.Add(b82);
                                                                                                                                                                                string b83 = Convert.ToString(W1.Range["B83"].Value);
                                                                                                                                                                                if (b83 != null && b83.Replace(" ", "") != "")
                                                                                                                                                                                {
                                                                                                                                                                                    _AGEN_mainform.lista_segments.Add(b83);
                                                                                                                                                                                    string b84 = Convert.ToString(W1.Range["B84"].Value);
                                                                                                                                                                                    if (b84 != null && b84.Replace(" ", "") != "")
                                                                                                                                                                                    {
                                                                                                                                                                                        _AGEN_mainform.lista_segments.Add(b84);
                                                                                                                                                                                        string b85 = Convert.ToString(W1.Range["B85"].Value);
                                                                                                                                                                                        if (b85 != null && b85.Replace(" ", "") != "")
                                                                                                                                                                                        {
                                                                                                                                                                                            _AGEN_mainform.lista_segments.Add(b85);
                                                                                                                                                                                            string b86 = Convert.ToString(W1.Range["B86"].Value);
                                                                                                                                                                                            if (b86 != null && b86.Replace(" ", "") != "")
                                                                                                                                                                                            {
                                                                                                                                                                                                _AGEN_mainform.lista_segments.Add(b86);
                                                                                                                                                                                                string b87 = Convert.ToString(W1.Range["B87"].Value);
                                                                                                                                                                                                if (b87 != null && b87.Replace(" ", "") != "")
                                                                                                                                                                                                {
                                                                                                                                                                                                    _AGEN_mainform.lista_segments.Add(b87);
                                                                                                                                                                                                    string b88 = Convert.ToString(W1.Range["B88"].Value);
                                                                                                                                                                                                    if (b88 != null && b88.Replace(" ", "") != "")
                                                                                                                                                                                                    {
                                                                                                                                                                                                        _AGEN_mainform.lista_segments.Add(b88);
                                                                                                                                                                                                        string b89 = Convert.ToString(W1.Range["B89"].Value);
                                                                                                                                                                                                        if (b89 != null && b89.Replace(" ", "") != "")
                                                                                                                                                                                                        {
                                                                                                                                                                                                            _AGEN_mainform.lista_segments.Add(b89);
                                                                                                                                                                                                            string b90 = Convert.ToString(W1.Range["B90"].Value);
                                                                                                                                                                                                            if (b90 != null && b90.Replace(" ", "") != "")
                                                                                                                                                                                                            {
                                                                                                                                                                                                                _AGEN_mainform.lista_segments.Add(b90);
                                                                                                                                                                                                                string b91 = Convert.ToString(W1.Range["B91"].Value);
                                                                                                                                                                                                                if (b91 != null && b91.Replace(" ", "") != "")
                                                                                                                                                                                                                {
                                                                                                                                                                                                                    _AGEN_mainform.lista_segments.Add(b91);
                                                                                                                                                                                                                    string b92 = Convert.ToString(W1.Range["B92"].Value);
                                                                                                                                                                                                                    if (b92 != null && b92.Replace(" ", "") != "")
                                                                                                                                                                                                                    {
                                                                                                                                                                                                                        _AGEN_mainform.lista_segments.Add(b92);
                                                                                                                                                                                                                        string b93 = Convert.ToString(W1.Range["B93"].Value);
                                                                                                                                                                                                                        if (b93 != null && b93.Replace(" ", "") != "")
                                                                                                                                                                                                                        {
                                                                                                                                                                                                                            _AGEN_mainform.lista_segments.Add(b93);
                                                                                                                                                                                                                            string b94 = Convert.ToString(W1.Range["B94"].Value);
                                                                                                                                                                                                                            if (b94 != null && b94.Replace(" ", "") != "")
                                                                                                                                                                                                                            {
                                                                                                                                                                                                                                _AGEN_mainform.lista_segments.Add(b94);
                                                                                                                                                                                                                                string b95 = Convert.ToString(W1.Range["B95"].Value);
                                                                                                                                                                                                                                if (b95 != null && b95.Replace(" ", "") != "")
                                                                                                                                                                                                                                {
                                                                                                                                                                                                                                    _AGEN_mainform.lista_segments.Add(b95);
                                                                                                                                                                                                                                    string b96 = Convert.ToString(W1.Range["B96"].Value);
                                                                                                                                                                                                                                    if (b96 != null && b96.Replace(" ", "") != "")
                                                                                                                                                                                                                                    {
                                                                                                                                                                                                                                        _AGEN_mainform.lista_segments.Add(b96);
                                                                                                                                                                                                                                        string b97 = Convert.ToString(W1.Range["B97"].Value);
                                                                                                                                                                                                                                        if (b97 != null && b97.Replace(" ", "") != "")
                                                                                                                                                                                                                                        {
                                                                                                                                                                                                                                            _AGEN_mainform.lista_segments.Add(b97);
                                                                                                                                                                                                                                            string b98 = Convert.ToString(W1.Range["B98"].Value);
                                                                                                                                                                                                                                            if (b98 != null && b98.Replace(" ", "") != "")
                                                                                                                                                                                                                                            {
                                                                                                                                                                                                                                                _AGEN_mainform.lista_segments.Add(b98);
                                                                                                                                                                                                                                                string b99 = Convert.ToString(W1.Range["B99"].Value);
                                                                                                                                                                                                                                                if (b99 != null && b99.Replace(" ", "") != "")
                                                                                                                                                                                                                                                {
                                                                                                                                                                                                                                                    _AGEN_mainform.lista_segments.Add(b99);
                                                                                                                                                                                                                                                    string b100 = Convert.ToString(W1.Range["B100"].Value);
                                                                                                                                                                                                                                                    if (b100 != null && b100.Replace(" ", "") != "")
                                                                                                                                                                                                                                                    {
                                                                                                                                                                                                                                                        _AGEN_mainform.lista_segments.Add(b100);
                                                                                                                                                                                                                                                        string b101 = Convert.ToString(W1.Range["B101"].Value);
                                                                                                                                                                                                                                                        if (b101 != null && b101.Replace(" ", "") != "")
                                                                                                                                                                                                                                                        {
                                                                                                                                                                                                                                                            _AGEN_mainform.lista_segments.Add(b101);
                                                                                                                                                                                                                                                            string b102 = Convert.ToString(W1.Range["B102"].Value);
                                                                                                                                                                                                                                                            if (b102 != null && b102.Replace(" ", "") != "")
                                                                                                                                                                                                                                                            {
                                                                                                                                                                                                                                                                _AGEN_mainform.lista_segments.Add(b102);
                                                                                                                                                                                                                                                                string b103 = Convert.ToString(W1.Range["B103"].Value);
                                                                                                                                                                                                                                                                if (b103 != null && b103.Replace(" ", "") != "")
                                                                                                                                                                                                                                                                {
                                                                                                                                                                                                                                                                    _AGEN_mainform.lista_segments.Add(b103);
                                                                                                                                                                                                                                                                    string b104 = Convert.ToString(W1.Range["B104"].Value);
                                                                                                                                                                                                                                                                    if (b104 != null && b104.Replace(" ", "") != "")
                                                                                                                                                                                                                                                                    {
                                                                                                                                                                                                                                                                        _AGEN_mainform.lista_segments.Add(b104);
                                                                                                                                                                                                                                                                        string b105 = Convert.ToString(W1.Range["B105"].Value);
                                                                                                                                                                                                                                                                        if (b105 != null && b105.Replace(" ", "") != "")
                                                                                                                                                                                                                                                                        {
                                                                                                                                                                                                                                                                            _AGEN_mainform.lista_segments.Add(b105);
                                                                                                                                                                                                                                                                            string b106 = Convert.ToString(W1.Range["B106"].Value);
                                                                                                                                                                                                                                                                            if (b106 != null && b106.Replace(" ", "") != "")
                                                                                                                                                                                                                                                                            {
                                                                                                                                                                                                                                                                                _AGEN_mainform.lista_segments.Add(b106);
                                                                                                                                                                                                                                                                                string b107 = Convert.ToString(W1.Range["B107"].Value);
                                                                                                                                                                                                                                                                                if (b107 != null && b107.Replace(" ", "") != "")
                                                                                                                                                                                                                                                                                {
                                                                                                                                                                                                                                                                                    _AGEN_mainform.lista_segments.Add(b107);
                                                                                                                                                                                                                                                                                    string b108 = Convert.ToString(W1.Range["B108"].Value);
                                                                                                                                                                                                                                                                                    if (b108 != null && b108.Replace(" ", "") != "")
                                                                                                                                                                                                                                                                                    {
                                                                                                                                                                                                                                                                                        _AGEN_mainform.lista_segments.Add(b108);
                                                                                                                                                                                                                                                                                        string b109 = Convert.ToString(W1.Range["B109"].Value);
                                                                                                                                                                                                                                                                                        if (b109 != null && b109.Replace(" ", "") != "")
                                                                                                                                                                                                                                                                                        {
                                                                                                                                                                                                                                                                                            _AGEN_mainform.lista_segments.Add(b109);
                                                                                                                                                                                                                                                                                            string b110 = Convert.ToString(W1.Range["B110"].Value);
                                                                                                                                                                                                                                                                                            if (b110 != null && b110.Replace(" ", "") != "")
                                                                                                                                                                                                                                                                                            {
                                                                                                                                                                                                                                                                                                _AGEN_mainform.lista_segments.Add(b110);
                                                                                                                                                                                                                                                                                                string b111 = Convert.ToString(W1.Range["B111"].Value);
                                                                                                                                                                                                                                                                                                if (b111 != null && b111.Replace(" ", "") != "")
                                                                                                                                                                                                                                                                                                {
                                                                                                                                                                                                                                                                                                    _AGEN_mainform.lista_segments.Add(b111);
                                                                                                                                                                                                                                                                                                    string b112 = Convert.ToString(W1.Range["B112"].Value);
                                                                                                                                                                                                                                                                                                    if (b112 != null && b112.Replace(" ", "") != "")
                                                                                                                                                                                                                                                                                                    {
                                                                                                                                                                                                                                                                                                        _AGEN_mainform.lista_segments.Add(b112);
                                                                                                                                                                                                                                                                                                        string b113 = Convert.ToString(W1.Range["B113"].Value);
                                                                                                                                                                                                                                                                                                        if (b113 != null && b113.Replace(" ", "") != "")
                                                                                                                                                                                                                                                                                                        {
                                                                                                                                                                                                                                                                                                            _AGEN_mainform.lista_segments.Add(b113);
                                                                                                                                                                                                                                                                                                            string b114 = Convert.ToString(W1.Range["B114"].Value);
                                                                                                                                                                                                                                                                                                            if (b114 != null && b114.Replace(" ", "") != "")
                                                                                                                                                                                                                                                                                                            {
                                                                                                                                                                                                                                                                                                                _AGEN_mainform.lista_segments.Add(b114);
                                                                                                                                                                                                                                                                                                                string b115 = Convert.ToString(W1.Range["B115"].Value);
                                                                                                                                                                                                                                                                                                                if (b115 != null && b115.Replace(" ", "") != "")
                                                                                                                                                                                                                                                                                                                {
                                                                                                                                                                                                                                                                                                                    _AGEN_mainform.lista_segments.Add(b115);
                                                                                                                                                                                                                                                                                                                    string b116 = Convert.ToString(W1.Range["B116"].Value);
                                                                                                                                                                                                                                                                                                                    if (b116 != null && b116.Replace(" ", "") != "")
                                                                                                                                                                                                                                                                                                                    {
                                                                                                                                                                                                                                                                                                                        _AGEN_mainform.lista_segments.Add(b116);
                                                                                                                                                                                                                                                                                                                        string b117 = Convert.ToString(W1.Range["B117"].Value);
                                                                                                                                                                                                                                                                                                                        if (b117 != null && b117.Replace(" ", "") != "")
                                                                                                                                                                                                                                                                                                                        {
                                                                                                                                                                                                                                                                                                                            _AGEN_mainform.lista_segments.Add(b117);
                                                                                                                                                                                                                                                                                                                            string b118 = Convert.ToString(W1.Range["B118"].Value);
                                                                                                                                                                                                                                                                                                                            if (b118 != null && b118.Replace(" ", "") != "")
                                                                                                                                                                                                                                                                                                                            {
                                                                                                                                                                                                                                                                                                                                _AGEN_mainform.lista_segments.Add(b118);
                                                                                                                                                                                                                                                                                                                                string b119 = Convert.ToString(W1.Range["B119"].Value);
                                                                                                                                                                                                                                                                                                                                if (b119 != null && b119.Replace(" ", "") != "")
                                                                                                                                                                                                                                                                                                                                {
                                                                                                                                                                                                                                                                                                                                    _AGEN_mainform.lista_segments.Add(b119);
                                                                                                                                                                                                                                                                                                                                    string b120 = Convert.ToString(W1.Range["B120"].Value);
                                                                                                                                                                                                                                                                                                                                    if (b120 != null && b120.Replace(" ", "") != "")
                                                                                                                                                                                                                                                                                                                                    {
                                                                                                                                                                                                                                                                                                                                        _AGEN_mainform.lista_segments.Add(b120);
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
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    _AGEN_mainform.lista_segments = null;
                                }

                                transfer_segment_data();

                                if (b1 != null) textBox_client_name.Text = b1.ToString();
                                if (b2 != null) textBox_project_name.Text = b2.ToString();

                                string template1 = "";
                                string template2 = "";
                                

                                if (b3 != null)
                                {
                                    template1 = b3.ToString();

                                }


                                if (b4 != null)
                                {
                                    template2 = b4.ToString();

                                }

                                if (System.IO.File.Exists(template1) == true)
                                {
                                    _AGEN_mainform.template1 = template1;
                                }
                                else
                                {
                                    _AGEN_mainform.template1 = "";
                                }
                                if (System.IO.File.Exists(template2) == true)
                                {
                                    _AGEN_mainform.template2 = template2;
                                }
                                else
                                {
                                    _AGEN_mainform.template2 = "";
                                }

                                if (System.IO.File.Exists(template1) == true)
                                {
                                    _AGEN_mainform.tpage_viewport_settings.set_textBox_template_name(template1);
                                }
                                else if (System.IO.File.Exists(template2) == true)
                                {
                                    _AGEN_mainform.tpage_viewport_settings.set_textBox_template_name(template2);
                                }

                                if (b5 != null)
                                {
                                    string Output = b5.ToString();
                                    if (System.IO.Directory.Exists(Output) == true)
                                    {
                                        Set_output_folder_text_box(Output);
                                    }
                                }

                                if (b6 != null)
                                {
                                    string Pref1 = b6.ToString();
                                    if (Pref1 != null)
                                    {
                                        _AGEN_mainform.tpage_viewport_settings.Set_prefix_text_box(Pref1);
                                    }
                                }

                                if (b7 != null)
                                {
                                    _AGEN_mainform.tpage_viewport_settings.Set_suffix_text_box(b7);
                                }

                                if (b8 != null)
                                {
                                    string Startno = b8.ToString();
                                    if (Functions.IsNumeric(Startno) == true)
                                    {
                                        _AGEN_mainform.tpage_viewport_settings.Set_start_no_text_box(Startno);
                                    }
                                }

                                if (b9 != null)
                                {
                                    string Increment = b9.ToString();
                                    if (Functions.IsNumeric(Increment) == true)
                                    {
                                        _AGEN_mainform.tpage_viewport_settings.Set_increment_text_box(Increment);
                                    }
                                }

                                if (b10 != null)
                                {
                                    string country = b10.ToString();
                                    if (country.ToUpper() == "CANADA")
                                    {
                                        _AGEN_mainform.COUNTRY = "CANADA";
                                        _AGEN_mainform.units_of_measurement = "m";

                                        radioButton_canada.Checked = true;

                                    }
                                    else
                                    {
                                        _AGEN_mainform.COUNTRY = "USA";
                                        _AGEN_mainform.units_of_measurement = "f";
                                        radioButton_usa.Checked = true;

                                    }
                                }

                                if (b11 != null)
                                {
                                    string lr = b11.ToString();
                                    if (lr.ToUpper().Replace(" ", "") == "NO" || lr.ToUpper().Replace(" ", "") == "FALSE")
                                    {
                                        _AGEN_mainform.Left_to_Right = false;
                                        _AGEN_mainform.tpage_viewport_settings.set_radioButton_left_right(false);
                                    }
                                    else
                                    {
                                        _AGEN_mainform.Left_to_Right = true;
                                        _AGEN_mainform.tpage_viewport_settings.set_radioButton_left_right(true);
                                    }
                                }

                                if (b11 != null)
                                {
                                    string lr = b11.ToString();
                                    if (lr.ToUpper().Replace(" ", "") == "NO" || lr.ToUpper().Replace(" ", "") == "FALSE")
                                    {
                                        _AGEN_mainform.Left_to_Right = false;
                                        _AGEN_mainform.tpage_viewport_settings.set_radioButton_left_right(false);
                                    }
                                    else
                                    {
                                        _AGEN_mainform.Left_to_Right = true;
                                        _AGEN_mainform.tpage_viewport_settings.set_radioButton_left_right(true);
                                    }
                                }
                                if (b12 != null)
                                {
                                    _AGEN_mainform.version = b12.ToString();
                                }
                                else
                                {
                                    _AGEN_mainform.version = "";
                                }

                                if (b15 != null && b15 == "meters")
                                {
                                    _AGEN_mainform.units_of_measurement = "m";
                                    _AGEN_mainform.tpage_viewport_settings.Set_combobox_units_to_m();
                                }
                                else
                                {
                                    _AGEN_mainform.units_of_measurement = "f";
                                    _AGEN_mainform.tpage_viewport_settings.Set_combobox_units_to_ft();
                                }

                                if (b16 != null)
                                {
                                    _AGEN_mainform.ProjFolder = b16.ToString();
                                    string ProjFolder = _AGEN_mainform.ProjFolder;
                                    if (System.IO.Directory.Exists(ProjFolder) == true)
                                    {
                                        textbox_project_database_folder.Text = ProjFolder;

                                        // here is because of segment
                                        ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();


                                        if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                                        {
                                            ProjFolder = ProjFolder + "\\";
                                            _AGEN_mainform.ProjFolder = ProjFolder;
                                        }

                                        Microsoft.Office.Interop.Excel.Workbook Workbook2 = null;
                                        Microsoft.Office.Interop.Excel.Worksheet W2 = null;

                                        try
                                        {
                                            string fisier_si = ProjFolder + _AGEN_mainform.sheet_index_excel_name;
                                            if (System.IO.File.Exists(fisier_si) == true)
                                            {
                                                Workbook2 = Excel1.Workbooks.Open(fisier_si);
                                                W2 = Workbook2.Worksheets[1];
                                                _AGEN_mainform.dt_sheet_index = Functions.Build_Data_table_sheet_index_from_excel(W2, _AGEN_mainform.Start_row_Sheet_index + 1);
                                                _AGEN_mainform.tpage_sheetindex.set_dataGridView_sheet_index();
                                                Workbook2.Close();
                                            }

                                            string fisier_cl = ProjFolder + _AGEN_mainform.cl_excel_name;

                                            if (System.IO.File.Exists(fisier_cl) == true)
                                            {
                                                Load_centerline_and_station_equation(fisier_cl);
                                            }



                                            if (_AGEN_mainform.dt_centerline != null && _AGEN_mainform.dt_centerline.Rows.Count > 0)
                                            {
                                                Set_centerline_label_to_green();
                                            }
                                            else
                                            {
                                                Set_centerline_label_to_red();
                                            }
                                            _AGEN_mainform.tpage_st_eq.Populate_datagridview_with_equation_data();
                                        }
                                        catch (System.Exception ex)
                                        {
                                            System.Windows.Forms.MessageBox.Show(ex.Message);

                                        }
                                        finally
                                        {
                                            if (W2 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W2);
                                            if (Workbook2 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook2);

                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show("the Project database folder location is not specified\r\n" + ProjFolder + "\r\n operation aborted");

                                        return;
                                    }
                                }




                                if (b18 != null && b19 != null && b20 != null && b21 != null)
                                {
                                    if (b18 != null)
                                    {
                                        _AGEN_mainform.NA_name = b18.ToString();
                                    }

                                    if (b19 != null)
                                    {
                                        string Xna = b19.ToString();
                                        if (Functions.IsNumeric(Xna) == true)
                                        {
                                            _AGEN_mainform.NA_x = Convert.ToDouble(Xna);
                                        }
                                    }
                                    if (b20 != null)
                                    {
                                        string yna = b20.ToString();
                                        if (Functions.IsNumeric(yna) == true)
                                        {
                                            _AGEN_mainform.NA_y = Convert.ToDouble(yna);
                                        }
                                    }

                                    if (b21 != null)
                                    {
                                        string sc = b21.ToString();
                                        if (Functions.IsNumeric(sc) == true)
                                        {
                                            _AGEN_mainform.NA_scale = Convert.ToDouble(sc);
                                        }
                                    }

                                    _AGEN_mainform.Data_table_blocks = new System.Data.DataTable();
                                    _AGEN_mainform.Data_table_blocks.Columns.Add("TYPE", typeof(String));
                                    _AGEN_mainform.Data_table_blocks.Columns.Add("BLOCK_NAME", typeof(String));
                                    _AGEN_mainform.Data_table_blocks.Columns.Add("SCALE", typeof(double));
                                    _AGEN_mainform.Data_table_blocks.Columns.Add("X", typeof(double));
                                    _AGEN_mainform.Data_table_blocks.Columns.Add("Y", typeof(double));
                                    _AGEN_mainform.Data_table_blocks.Rows.Add();
                                    _AGEN_mainform.Data_table_blocks.Rows[0][0] = "North Arrow";
                                    _AGEN_mainform.Data_table_blocks.Rows[0][1] = _AGEN_mainform.NA_name;
                                    _AGEN_mainform.Data_table_blocks.Rows[0][2] = _AGEN_mainform.NA_scale;
                                    _AGEN_mainform.Data_table_blocks.Rows[0][3] = _AGEN_mainform.NA_x;
                                    _AGEN_mainform.Data_table_blocks.Rows[0][4] = _AGEN_mainform.NA_y;
                                    //old_code_commented >>> tabControl_work.SelectedTab = tabPage2;
                                    _AGEN_mainform.tpage_viewport_settings.set_dataGridView_north_arrow_blocks();
                                    //old_code_commented >>>    tabControl_work.SelectedTab = tabPage1;
                                }


                                if (b22 != null)
                                {
                                    _AGEN_mainform.tpage_viewport_settings.set_combobox_units_precision(b22);
                                    if (b22 == "0")
                                    {
                                        _AGEN_mainform.round1 = 0;
                                    }
                                    else if (b22 == "0.0")
                                    {
                                        _AGEN_mainform.round1 = 1;
                                    }
                                    else if (b22 == "0.00")
                                    {
                                        _AGEN_mainform.round1 = 2;
                                    }
                                    else if (b22 == "0.000")
                                    {
                                        _AGEN_mainform.round1 = 3;
                                    }
                                    else
                                    {
                                        _AGEN_mainform.round1 = 0;
                                    }
                                }

                                if (b23 != null)
                                {
                                    if (Convert.ToString(b23).Replace(" ", "") != "")
                                    {
                                        _AGEN_mainform.Matchline_BlockName_in_PaperSpace = Convert.ToString(b23);
                                    }
                                }

                                if (b36 != null)
                                {
                                    if (b36.ToLower() == "true" || b36.ToLower() == "yes")
                                    {
                                        _AGEN_mainform.Exista_viewport_main = true;
                                    }
                                }

                                if (b37 != null)
                                {
                                    if (b37.ToLower() == "true" || b37.ToLower() == "yes")
                                    {
                                        _AGEN_mainform.Exista_viewport_cross = true;
                                    }
                                }

                                if (b38 != null)
                                {
                                    if (b38.ToLower() == "true" || b38.ToLower() == "yes")
                                    {
                                        _AGEN_mainform.Exista_viewport_owner = true;
                                    }
                                }


                                if (b39 != null)
                                {
                                    if (b39.ToLower() == "true" || b39.ToLower() == "yes")
                                    {
                                        _AGEN_mainform.Exista_viewport_prof = true;
                                    }
                                }

                                if (b40 != null)
                                {
                                    if (b40.ToLower() == "3d")
                                    {
                                        _AGEN_mainform.tpage_sheetindex.set_radioButton_use3D_stations(true);
                                        _AGEN_mainform.Project_type = "3D";
                                    }
                                    else
                                    {
                                        _AGEN_mainform.tpage_sheetindex.set_radioButton_use2D_stations(true);
                                        _AGEN_mainform.Project_type = "2D";
                                    }
                                }

                                if (b41 != null)
                                {
                                    if (b41.ToLower() == "true" || b41.ToLower() == "yes")
                                    {
                                        _AGEN_mainform.Exista_viewport_mat = true;
                                    }
                                }

                                if (b42 != null)
                                {
                                    if (b42.ToLower() == "true" || b42.ToLower() == "yes")
                                    {
                                        _AGEN_mainform.Exista_viewport_prof_band = true;
                                    }
                                }


                                if (b43 != null)
                                {
                                    if (b43.ToLower() == "true" || b43.ToLower() == "yes")
                                    {
                                        _AGEN_mainform.Exista_viewport_tblk = true;
                                    }
                                }

                                Ag.Set_textBox_config_file_location(File1);

                            }
                            #endregion


                            #region Regular_band_data
                            else if (W1.Name == "Regular_band_data")
                            {

                                _AGEN_mainform.Data_Table_regular_bands = Functions.build_regular_band_data_table_from_excel(W1, 2);

                                string prop_band_name = _AGEN_mainform.tpage_viewport_settings.get_comboBox_viewport_target_areas(3);
                                string main_vp_name = _AGEN_mainform.tpage_viewport_settings.get_comboBox_viewport_target_areas(1);
                                string tblk_vp_name = _AGEN_mainform.tpage_viewport_settings.get_comboBox_viewport_target_areas(8);


                                if (_AGEN_mainform.Data_Table_regular_bands != null)
                                {
                                    if (_AGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                                    {
                                        for (int i = 0; i < _AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                                        {
                                            if (_AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] != DBNull.Value)
                                            {
                                                string bn = Convert.ToString(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"]);

                                                if (bn == main_vp_name)
                                                {
                                                    if (_AGEN_mainform.Data_Table_regular_bands.Rows[i]["Custom_scale"] != DBNull.Value)
                                                    {
                                                        string str_scale = Convert.ToString(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["Custom_scale"]);
                                                        if (Functions.IsNumeric(str_scale) == true)
                                                        {
                                                            _AGEN_mainform.Vw_scale = Convert.ToDouble(str_scale);




                                                            for (int k = 0; k < _AGEN_mainform.tpage_viewport_settings.Get_combobox_viewport_scale_count(); ++k)
                                                            {
                                                                string Scale1 = _AGEN_mainform.tpage_viewport_settings.Get_combobox_viewport_scale(k);

                                                                if (Scale1.Contains(":") == true)
                                                                {
                                                                    Scale1 = Scale1.Substring(2, Scale1.Length - 2);
                                                                    if (Functions.IsNumeric(Scale1) == true)
                                                                    {
                                                                        if (Math.Round(_AGEN_mainform.Vw_scale, 1) == Math.Round(1000 / Convert.ToDouble(Scale1), 1))
                                                                        {
                                                                            _AGEN_mainform.tpage_viewport_settings.Set_combobox_viewport_scale(k);
                                                                            k = _AGEN_mainform.tpage_viewport_settings.Get_combobox_viewport_scale_count();
                                                                        }
                                                                    }
                                                                }
                                                                string feet = "\u0022";

                                                                if (Scale1.Contains(feet + "=") == true && Scale1.Contains("'") == true)
                                                                {
                                                                    Scale1 = Scale1.Substring(3, Scale1.Length - 4);
                                                                    if (Functions.IsNumeric(Scale1) == true)
                                                                    {
                                                                        if (Math.Round(_AGEN_mainform.Vw_scale, 4) == Math.Round(1 / Convert.ToDouble(Scale1), 4))
                                                                        {
                                                                            _AGEN_mainform.tpage_viewport_settings.Set_combobox_viewport_scale(k);
                                                                            k = _AGEN_mainform.tpage_viewport_settings.Get_combobox_viewport_scale_count();
                                                                        }
                                                                    }
                                                                }
                                                            }

                                                            if (_AGEN_mainform.Vw_scale == 1 && _AGEN_mainform.COUNTRY == "USA")
                                                            {
                                                                _AGEN_mainform.tpage_viewport_settings.Set_combobox_viewport_scale(0);
                                                            }

                                                        }
                                                    }

                                                    if (_AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_x"] != DBNull.Value)
                                                    {
                                                        string str_val = Convert.ToString(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_x"]);
                                                        if (Functions.IsNumeric(str_val) == true)
                                                        {
                                                            _AGEN_mainform.Vw_ps_x = Convert.ToDouble(str_val);
                                                        }
                                                    }

                                                    if (_AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_y"] != DBNull.Value)
                                                    {
                                                        string str_val = Convert.ToString(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_y"]);
                                                        if (Functions.IsNumeric(str_val) == true)
                                                        {
                                                            _AGEN_mainform.Vw_ps_y = Convert.ToDouble(str_val);
                                                        }
                                                    }


                                                    if (_AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_width"] != DBNull.Value)
                                                    {
                                                        string str_val = Convert.ToString(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_width"]);
                                                        if (Functions.IsNumeric(str_val) == true)
                                                        {
                                                            _AGEN_mainform.Vw_width = Convert.ToDouble(str_val);
                                                        }
                                                    }

                                                    if (_AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_height"] != DBNull.Value)
                                                    {
                                                        string str_val = Convert.ToString(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_height"]);
                                                        if (Functions.IsNumeric(str_val) == true)
                                                        {
                                                            _AGEN_mainform.Vw_height = Convert.ToDouble(str_val);
                                                        }
                                                    }

                                                }




                                                if (bn == tblk_vp_name)
                                                {


                                                    if (_AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_x"] != DBNull.Value)
                                                    {
                                                        string str_val = Convert.ToString(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_x"]);
                                                        if (Functions.IsNumeric(str_val) == true)
                                                        {
                                                            _AGEN_mainform.Vw_ps_tblk_x = Convert.ToDouble(str_val);
                                                        }
                                                    }

                                                    if (_AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_y"] != DBNull.Value)
                                                    {
                                                        string str_val = Convert.ToString(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_y"]);
                                                        if (Functions.IsNumeric(str_val) == true)
                                                        {
                                                            _AGEN_mainform.Vw_ps_tblk_y = Convert.ToDouble(str_val);
                                                        }
                                                    }


                                                    if (_AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_width"] != DBNull.Value)
                                                    {
                                                        string str_val = Convert.ToString(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_width"]);
                                                        if (Functions.IsNumeric(str_val) == true)
                                                        {
                                                            _AGEN_mainform.Vw_tblk_width = Convert.ToDouble(str_val);
                                                        }
                                                    }

                                                    if (_AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_height"] != DBNull.Value)
                                                    {
                                                        string str_val = Convert.ToString(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_height"]);
                                                        if (Functions.IsNumeric(str_val) == true)
                                                        {
                                                            _AGEN_mainform.Vw_tblk_height = Convert.ToDouble(str_val);
                                                        }
                                                    }

                                                }



                                            }
                                        }
                                    }
                                }

                            }
                            #endregion
                            #region Custom_band_data
                            else if (W1.Name == "Custom_band_data")
                            {
                                try
                                {
                                    _AGEN_mainform.Data_Table_custom_bands = Functions.build_custom_band_data_table_from_excel(W1, 2);
                                    if (_AGEN_mainform.Data_Table_custom_bands != null && _AGEN_mainform.Data_Table_custom_bands.Rows.Count > 0 && _AGEN_mainform.Data_Table_custom_bands.Rows[0][0] != DBNull.Value)
                                    {
                                        _AGEN_mainform.first_custom_band = Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[0][0]);
                                    }

                                    // read and write parameters are loaded at changing index event!!!
                                }
                                catch (System.Exception ex)
                                {
                                    System.Windows.Forms.MessageBox.Show(ex.Message);
                                }
                            }
                            #endregion
                            #region Crossing_data_config
                            else if (W1.Name == "Crossing_data_config")
                            {
                                transfer_crossing_settings_from_excel(W1);
                            }
                            #endregion

                            #region Custom_data_config
                            else if (W1.Name == _AGEN_mainform.first_custom_band + "_cfg_" + segment1)
                            {
                                transfer_custom_band_settings_to_controls(W1);
                            }
                            #endregion

                            #region Ownership_data_config
                            else if (W1.Name == "Ownership_data_config_" + segment1)
                            {
                                transfer_ownership_band_settings_to_controls(W1);

                            }
                            #endregion

                            #region profile first segment pdc2
                            else if (W1.Name == "pdc2_" + segment1)
                            {
                                transfer_profile_settings_to_controls(W1);
                            }
                            #endregion

                            #region extra VP
                            else if (W1.Name == "ExtraVP_data")
                            {
                                _AGEN_mainform.Data_Table_extra_mainVP = Functions.build_extra_data_table_from_excel(W1, 2);
                                exista_extraVP = true;
                            }
                            #endregion

                        }
                        catch (System.Exception ex)
                        {
                            System.Windows.Forms.MessageBox.Show(ex.Message);
                        }
                    }

                    foreach (Microsoft.Office.Interop.Excel.Worksheet W1 in Workbook1.Worksheets)
                    {
                        try
                        {
                            #region build Custom_datatable_config
                            if (W1.Name.Contains("_cfg_" + segment1) == true)
                            {
                                build_dt_custom_settings(W1);
                            }
                            #endregion
                        }
                        catch (System.Exception ex)
                        {
                            System.Windows.Forms.MessageBox.Show(ex.Message);
                        }
                    }

                    if (exista_extraVP == false)
                    {
                        _AGEN_mainform.Data_Table_extra_mainVP = null;
                    }

                    Workbook1.Close();

                    if (Excel1.Workbooks.Count == 0)
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
                    if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);

                }

                display_checkboxes_into_generation_page();

                _AGEN_mainform.tpage_viewport_settings.creeaza_display_data_table(Functions.Creaza_lista_regular_vp_picked(), Functions.Creaza_lista_custom_vp_picked(), Functions.Creaza_lista_custom_vp_extra_picked());


                #region titleblock_attributes
                string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();

                if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                {
                    ProjF = ProjF + "\\";
                }
                string File2 = ProjF + _AGEN_mainform.block_attributes_excel_name;

                if (System.IO.File.Exists(File2) == true)
                {
                    Microsoft.Office.Interop.Excel.Workbook Workbook2 = Excel1.Workbooks.Open(File2);
                    Microsoft.Office.Interop.Excel.Worksheet W2 = Workbook2.Worksheets[1];
                    try
                    {
                        bool attr_loaded = Check_to_see_if_there_the_header_in_the_tblk_attributes(W2, _AGEN_mainform.Start_row_block_attributes);
                        if (attr_loaded == true)
                        {
                            _AGEN_mainform.tpage_tblk_attrib.Set_label_block_attributes_to_green();
                        }
                        Workbook2.Close();

                    }
                    catch (System.Exception ex)
                    {
                        System.Windows.Forms.MessageBox.Show(ex.Message);

                    }
                    finally
                    {
                        if (W2 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W2);
                        if (Workbook2 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook2);
                    }
                }

                #endregion
                if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);

            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        public void build_dt_custom_settings(Microsoft.Office.Interop.Excel.Worksheet W1)
        {
            string b1 = Convert.ToString(W1.Range["B1"].Value2);
            string b2 = Convert.ToString(W1.Range["B2"].Value);
            string b3 = Convert.ToString(W1.Range["B3"].Value);
            string b4 = Convert.ToString(W1.Range["B4"].Value);
            string b5 = Convert.ToString(W1.Range["B5"].Value);
            string b6 = Convert.ToString(W1.Range["B6"].Value);
            string b7 = Convert.ToString(W1.Range["B7"].Value);
            string b8 = Convert.ToString(W1.Range["B8"].Value);
            string b9 = Convert.ToString(W1.Range["B9"].Value);
            string b10 = Convert.ToString(W1.Range["B10"].Value);

            if (b1 == null) b1 = "";
            if (b2 == null) b2 = "";
            if (b3 == null) b3 = "";
            if (b4 == null) b4 = "";
            if (b5 == null) b5 = "";
            if (b6 == null) b6 = "";
            if (b7 == null) b7 = "";
            if (b8 == null) b8 = "";
            if (b9 == null) b9 = "";
            if (b10 == null) b10 = "";

            if (_AGEN_mainform.dt_settings_custom == null)
            {
                string values0 = "Band Excel File Name";
                string values1 = "OD Table";
                string values2 = "OD Field1";
                string values3 = "OD Field2";

                string values4 = "Custom Band Block";
                string values5 = "Block Tag Sta1";
                string values6 = "Block Tag Sta2";
                string values7 = "Block Tag Length";
                string values8 = "Block Tag Attribute 1";
                string values9 = "Block Tag Attribute 2";

                _AGEN_mainform.dt_settings_custom = new System.Data.DataTable();
                _AGEN_mainform.dt_settings_custom.Columns.Add(values0, typeof(string));
                _AGEN_mainform.dt_settings_custom.Columns.Add(values1, typeof(string));
                _AGEN_mainform.dt_settings_custom.Columns.Add(values2, typeof(string));
                _AGEN_mainform.dt_settings_custom.Columns.Add(values3, typeof(string));
                _AGEN_mainform.dt_settings_custom.Columns.Add(values4, typeof(string));
                _AGEN_mainform.dt_settings_custom.Columns.Add(values5, typeof(string));
                _AGEN_mainform.dt_settings_custom.Columns.Add(values6, typeof(string));
                _AGEN_mainform.dt_settings_custom.Columns.Add(values7, typeof(string));
                _AGEN_mainform.dt_settings_custom.Columns.Add(values8, typeof(string));
                _AGEN_mainform.dt_settings_custom.Columns.Add(values9, typeof(string));
            }

            _AGEN_mainform.dt_settings_custom.Rows.Add();

            if (b1 != "")
            {
                _AGEN_mainform.dt_settings_custom.Rows[_AGEN_mainform.dt_settings_custom.Rows.Count - 1][0] = b1;
            }
            if (b2 != "")
            {
                _AGEN_mainform.dt_settings_custom.Rows[_AGEN_mainform.dt_settings_custom.Rows.Count - 1][1] = b2;
            }
            if (b3 != "")
            {
                _AGEN_mainform.dt_settings_custom.Rows[_AGEN_mainform.dt_settings_custom.Rows.Count - 1][2] = b3;
            }
            if (b4 != "")
            {
                _AGEN_mainform.dt_settings_custom.Rows[_AGEN_mainform.dt_settings_custom.Rows.Count - 1][3] = b4;
            }
            if (b5 != "")
            {
                _AGEN_mainform.dt_settings_custom.Rows[_AGEN_mainform.dt_settings_custom.Rows.Count - 1][4] = b5;
            }
            if (b6 != "")
            {
                _AGEN_mainform.dt_settings_custom.Rows[_AGEN_mainform.dt_settings_custom.Rows.Count - 1][5] = b6;
            }
            if (b7 != "")
            {
                _AGEN_mainform.dt_settings_custom.Rows[_AGEN_mainform.dt_settings_custom.Rows.Count - 1][6] = b7;

            }
            if (b8 != "")
            {
                _AGEN_mainform.dt_settings_custom.Rows[_AGEN_mainform.dt_settings_custom.Rows.Count - 1][7] = b8;

            }
            if (b9 != "")
            {
                _AGEN_mainform.dt_settings_custom.Rows[_AGEN_mainform.dt_settings_custom.Rows.Count - 1][8] = b9;

            }
            if (b10 != "")
            {
                _AGEN_mainform.dt_settings_custom.Rows[_AGEN_mainform.dt_settings_custom.Rows.Count - 1][9] = b10;
            }


        }

        public void transfer_custom_band_settings_to_controls(Microsoft.Office.Interop.Excel.Worksheet W1)
        {
            string b1 = Convert.ToString(W1.Range["B1"].Value2);
            string b2 = Convert.ToString(W1.Range["B2"].Value);
            string b3 = Convert.ToString(W1.Range["B3"].Value);
            string b4 = Convert.ToString(W1.Range["B4"].Value);
            string b5 = Convert.ToString(W1.Range["B5"].Value);
            string b6 = Convert.ToString(W1.Range["B6"].Value);
            string b7 = Convert.ToString(W1.Range["B7"].Value);
            string b8 = Convert.ToString(W1.Range["B8"].Value);
            string b9 = Convert.ToString(W1.Range["B9"].Value);
            string b10 = Convert.ToString(W1.Range["B10"].Value);

            if (b1 == null) b1 = "";
            if (b2 == null) b2 = "";
            if (b3 == null) b3 = "";
            if (b4 == null) b4 = "";
            if (b5 == null) b5 = "";
            if (b6 == null) b6 = "";
            if (b7 == null) b7 = "";
            if (b8 == null) b8 = "";
            if (b9 == null) b9 = "";
            if (b10 == null) b10 = "";


            if (b1 != "")
            {
                _AGEN_mainform.tpage_cust_scan.set_comboBox_band_excel_name(b1);
                _AGEN_mainform.tpage_cust_draw.set_comboBox_custom_excel_name(b1);
            }
            if (b2 != "")
            {
                _AGEN_mainform.tpage_cust_scan.set_comboBox_custom_od_table(b2);

            }
            if (b3 != "")
            {
                _AGEN_mainform.tpage_cust_scan.set_comboBox_custom_field1_od(b3);

            }
            if (b4 != "")
            {
                _AGEN_mainform.tpage_cust_scan.set_comboBox_custom_field2_od(b4);

            }
            if (b5 != "")
            {
                _AGEN_mainform.tpage_cust_draw.set_comboBox_custom_block(b5);
            }
            if (b6 != "")
            {
                _AGEN_mainform.tpage_cust_draw.set_comboBox_custom_atr_sta1(b6);
            }
            if (b7 != "")
            {
                _AGEN_mainform.tpage_cust_draw.set_comboBox_custom_atr_sta2(b7);
            }
            if (b8 != "")
            {
                _AGEN_mainform.tpage_cust_draw.set_comboBox_custom_atr_distance(b8);
            }
            if (b9 != "")
            {
                _AGEN_mainform.tpage_cust_draw.set_comboBox_custom_atr_field1(b9);
            }
            if (b10 != "")
            {
                _AGEN_mainform.tpage_cust_draw.set_comboBox_custom_atr_field2(b10);
            }


        }



        public void transfer_ownership_band_settings_to_controls(Microsoft.Office.Interop.Excel.Worksheet W1)
        {

            string b1 = Convert.ToString(W1.Range["B1"].Value2);
            string b2 = Convert.ToString(W1.Range["B2"].Value);
            string b3 = Convert.ToString(W1.Range["B3"].Value);
            string b4 = Convert.ToString(W1.Range["B4"].Value);
            string b5 = Convert.ToString(W1.Range["B5"].Value);
            string b6 = Convert.ToString(W1.Range["B6"].Value);
            string b7 = Convert.ToString(W1.Range["B7"].Value);
            string b8 = Convert.ToString(W1.Range["B8"].Value);
            string b9 = Convert.ToString(W1.Range["B9"].Value);


            if (b1 == null) b1 = "";
            if (b2 == null) b2 = "";
            if (b3 == null) b3 = "";
            if (b4 == null) b4 = "";
            if (b5 == null) b5 = "";
            if (b6 == null) b6 = "";
            if (b7 == null) b7 = "";
            if (b8 == null) b8 = "";
            if (b9 == null) b9 = "";

            if (b1 != "")
            {
                _AGEN_mainform.tpage_owner_draw.set_comboBox_prop_block(b1);
            }

            if (b2 != "")
            {
                _AGEN_mainform.tpage_owner_draw.set_comboBox_prop_atr_sta1(b2);
            }

            if (b3 != "")
            {
                _AGEN_mainform.tpage_owner_draw.set_comboBox_prop_atr_sta2(b3);
            }

            if (b4 != "")
            {
                _AGEN_mainform.tpage_owner_draw.set_comboBox_prop_atr_distance(b4);
            }

            if (b5 != "")
            {
                _AGEN_mainform.tpage_owner_draw.set_comboBox_prop_atr_linelist(b5);
            }

            if (b6 != "")
            {
                _AGEN_mainform.tpage_owner_draw.set_comboBox_prop_atr_owner(b6);
            }

            if (b7 != "")
            {
                _AGEN_mainform.tpage_owner_scan.set_comboBox_prop_od_table(b7);
            }

            if (b8 != "")
            {
                _AGEN_mainform.tpage_owner_scan.set_comboBox_prop_owner_od(b8);
            }

            if (b9 != "")
            {
                _AGEN_mainform.tpage_owner_scan.set_comboBox_prop_linelist_od(b9);
            }

        }


        public void transfer_profile_settings_to_controls(Microsoft.Office.Interop.Excel.Worksheet W1)
        {


            string b7 = Convert.ToString(W1.Range["B7"].Value);
            string b8 = Convert.ToString(W1.Range["B8"].Value);
            string b9 = Convert.ToString(W1.Range["B9"].Value);
            string b10 = Convert.ToString(W1.Range["B10"].Value);


            string b14 = Convert.ToString(W1.Range["B14"].Value);
            string b15 = Convert.ToString(W1.Range["B15"].Value);
            string b16 = Convert.ToString(W1.Range["B16"].Value);
            string b17 = Convert.ToString(W1.Range["B17"].Value);
            string b18 = Convert.ToString(W1.Range["B18"].Value);
            string b19 = Convert.ToString(W1.Range["B19"].Value);
            string b20 = Convert.ToString(W1.Range["B20"].Value);
            string b21 = Convert.ToString(W1.Range["B21"].Value);
            string b22 = Convert.ToString(W1.Range["B22"].Value);
            string b23 = Convert.ToString(W1.Range["B23"].Value);
            string b24 = Convert.ToString(W1.Range["B24"].Value);
            string b25 = Convert.ToString(W1.Range["B25"].Value);
            string b26 = Convert.ToString(W1.Range["B26"].Value);


            if (b7 != null)
            {
                if (Functions.IsNumeric(b7) == true)
                {
                    _AGEN_mainform.tpage_profdraw.set_textBox_prof_Hex(b7);
                }
            }

            if (b8 != null)
            {
                if (Functions.IsNumeric(b8) == true)
                {
                    _AGEN_mainform.tpage_profdraw.set_textBox_prof_Vex(b8);
                }
            }

            if (b9 != null)
            {
                if (Functions.IsNumeric(b9) == true)
                {
                    _AGEN_mainform.tpage_profdraw.set_textBox_prof_Elev_bottom(b9);
                }
            }


            if (b10 != null)
            {
                if (Functions.IsNumeric(b10) == true)
                {
                    _AGEN_mainform.tpage_profdraw.set_textBox_prof_Elev_top(b10);
                }
            }

            if (b14 != null)
            {

                _AGEN_mainform.tpage_profdraw.set_comboBox_prof_textstyle(b14);

            }

            if (b15 != null)
            {
                if (Functions.IsNumeric(b15) == true)
                {
                    _AGEN_mainform.tpage_profdraw.set_textBox_prof_Hspacing(b15);
                }
            }


            if (b16 != null)
            {
                if (Functions.IsNumeric(b16) == true)
                {
                    _AGEN_mainform.tpage_profdraw.set_textBox_prof_Vspacing(b16);
                }
            }


            if (b17 != null)
            {

                _AGEN_mainform.tpage_profdraw.set_comboBox_prof_el_lbl_loc(b17);

            }

            if (b18 != null)
            {
                if (Functions.IsNumeric(b18) == true)
                {
                    _AGEN_mainform.tpage_profdraw.set_textBox_elev_round(Convert.ToString(Math.Abs(Convert.ToInt32(b18))));
                }
                else
                {
                    _AGEN_mainform.tpage_profdraw.set_textBox_elev_round("0");
                }
            }
            else
            {
                _AGEN_mainform.tpage_profdraw.set_textBox_elev_round("0");
            }

            if (b19 != null)
            {
                if (Functions.IsNumeric(b19) == true && Math.Abs(Convert.ToInt32(b19)) == 90)
                {
                    _AGEN_mainform.tpage_profdraw.set_checkBox_sta_at_90(true);
                }
                else
                {

                    _AGEN_mainform.tpage_profdraw.set_checkBox_sta_at_90(false);

                }
            }
            else
            {

                _AGEN_mainform.tpage_profdraw.set_checkBox_sta_at_90(false);

            }




            if (b21 != null)
            {
                if (b21.ToLower() == "true" || b21.ToLower() == "yes")
                {
                    _AGEN_mainform.tpage_profdraw.set_checkBox_draw_ver_at_start(true);
                }
                else
                {
                    _AGEN_mainform.tpage_profdraw.set_checkBox_draw_ver_at_start(false);
                }
            }
            else
            {
                _AGEN_mainform.tpage_profdraw.set_checkBox_draw_ver_at_start(false);
            }


            if (b22 != null)
            {
                if (b22.ToLower() == "true" || b22.ToLower() == "yes")
                {
                    _AGEN_mainform.tpage_profdraw.set_checkBox_set_zero_at_middle_of_profile(true);
                }
                else
                {
                    _AGEN_mainform.tpage_profdraw.set_checkBox_set_zero_at_middle_of_profile(false);
                }
            }
            else
            {
                _AGEN_mainform.tpage_profdraw.set_checkBox_set_zero_at_middle_of_profile(false);
            }

            if (b23 != null)
            {
                if (b23.ToLower() == "true" || b23.ToLower() == "yes")
                {
                    _AGEN_mainform.tpage_profdraw.set_checkBox_hydro_style(true);
                }
                else
                {
                    _AGEN_mainform.tpage_profdraw.set_checkBox_hydro_style(false);
                }
            }
            else
            {
                _AGEN_mainform.tpage_profdraw.set_checkBox_hydro_style(false);
            }

            if (b24 != null)
            {
                if (b24.ToLower() == "true" || b24.ToLower() == "yes")
                {
                    _AGEN_mainform.tpage_profdraw.set_checkBox_sta_at_90(true);
                }
                else
                {
                    _AGEN_mainform.tpage_profdraw.set_checkBox_sta_at_90(false);
                }
            }
            else
            {
                _AGEN_mainform.tpage_profdraw.set_checkBox_sta_at_90(false);
            }


            if (b26 != null)
            {
                if (Functions.IsNumeric(b26) == true)
                {
                    _AGEN_mainform.tpage_profdraw.set_textBox_overwrite_text_height(b26);
                    _AGEN_mainform.tpage_profdraw.set_checkBox_overwrite_text_height(true);
                }
                else
                {
                    _AGEN_mainform.tpage_profdraw.set_textBox_overwrite_text_height("");
                    _AGEN_mainform.tpage_profdraw.set_checkBox_overwrite_text_height(false);
                }
            }
            else
            {
                _AGEN_mainform.tpage_profdraw.set_textBox_overwrite_text_height("");
                _AGEN_mainform.tpage_profdraw.set_checkBox_overwrite_text_height(false);
            }

        }

        public void transfer_crossing_settings_from_excel(Microsoft.Office.Interop.Excel.Worksheet W1)
        {

            string b1 = Convert.ToString(W1.Range["B1"].Value2);
            string b2 = Convert.ToString(W1.Range["B2"].Value);
            string b3 = Convert.ToString(W1.Range["B3"].Value);
            string b4 = Convert.ToString(W1.Range["B4"].Value);
            string b5 = Convert.ToString(W1.Range["B5"].Value);
            string b6 = Convert.ToString(W1.Range["B6"].Value);
            string b7 = Convert.ToString(W1.Range["B7"].Value);
            string b8 = Convert.ToString(W1.Range["B8"].Value);
            string b9 = Convert.ToString(W1.Range["B9"].Value);
            string b10 = Convert.ToString(W1.Range["B10"].Value);
            string b11 = Convert.ToString(W1.Range["B11"].Value);
            string b12 = Convert.ToString(W1.Range["B12"].Value);
            string b13 = Convert.ToString(W1.Range["B13"].Value);
            string b14 = Convert.ToString(W1.Range["B14"].Value);
            string b15 = Convert.ToString(W1.Range["B15"].Value);
            string b16 = Convert.ToString(W1.Range["B16"].Value);
            string b17 = Convert.ToString(W1.Range["B17"].Value);
            double th = -1;

            if (b1 == null) b1 = "";
            if (b2 == null) b2 = "";
            if (b3 == null) b3 = "";
            if (b4 == null) b4 = "";
            if (b5 == null) b5 = "";

            bool boolb6 = false;
            if (b6 != null)
            {
                if (b6.ToLower() == "true" || b6.ToLower() == "yes") boolb6 = true;
            }
            bool boolb7 = false;
            if (b7 != null)
            {
                if (b7.ToLower() == "true" || b7.ToLower() == "yes") boolb7 = true;
            }

            if (b8 == null) b8 = "";
            if (b9 == null) b9 = "";

            bool boolb10 = false;
            if (b10 != null)
            {
                if (b10.ToLower() == "true" || b10.ToLower() == "yes") boolb10 = true;
            }
            bool boolb11 = false;
            if (b11 != null)
            {
                if (b11.ToLower() == "true" || b11.ToLower() == "yes") boolb11 = true;
            }
            bool boolb12 = false;
            if (b12 != null)
            {
                if (b12.ToLower() == "true" || b12.ToLower() == "yes") boolb12 = true;
            }

            bool boolb17 = false;
            if (b17 != null)
            {
                if (b17.ToLower() == "true" || b17.ToLower() == "yes") boolb17 = true;
            }


            if (b13 != null && Functions.IsNumeric(b13) == true)
            {
                _AGEN_mainform.XingDeltay1 = Convert.ToDouble(b13);
            }
            if (b14 != null && Functions.IsNumeric(b14) == true)
            {
                _AGEN_mainform.XingDeltay2 = Convert.ToDouble(b14);
            }
            if (b15 != null && Functions.IsNumeric(b15) == true)
            {
                _AGEN_mainform.XingDeltay3 = Convert.ToDouble(b15);
            }

            if (b16 != null && Functions.IsNumeric(b16) == true && Convert.ToDouble(b16) > 0)
            {
                th = Convert.ToDouble(b16);
            }

            _AGEN_mainform.tpage_crossing_draw.write_crossing_settings_to_controls(b1, b2, b3, b4, b5, boolb6, boolb7, b8, b9, boolb10, boolb11, boolb12, boolb17, th);
        }


        public void Build_sheet_index_dt_from_excel(string sheetname = "")
        {
            string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();

            if (System.IO.Directory.Exists(ProjFolder) == true)
            {

                if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                {
                    ProjFolder = ProjFolder + "\\";
                }

                bool excel_is_opened = false;
                string fisier_si = ProjFolder + _AGEN_mainform.sheet_index_excel_name;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;
                Microsoft.Office.Interop.Excel.Worksheet W1 = null;
                Microsoft.Office.Interop.Excel.Workbook Workbook3 = null;
                Microsoft.Office.Interop.Excel.Worksheet W3 = null;

                Microsoft.Office.Interop.Excel.Application Excel1 = null;

                try
                {
                    Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");

                    foreach (Microsoft.Office.Interop.Excel.Workbook Workbook2 in Excel1.Workbooks)
                    {
                        if (Workbook2.FullName.ToLower() == fisier_si.ToLower())
                        {
                            Workbook1 = Workbook2;

                            if (sheetname == "")
                            {
                                W1 = Workbook1.Worksheets[1];
                            }
                            else
                            {
                                foreach (Microsoft.Office.Interop.Excel.Worksheet W2 in Workbook1.Worksheets)
                                {
                                    if (W2.Name.ToLower() == sheetname.ToLower())
                                    {
                                        W1 = W2;
                                    }
                                }
                                if (W1 == null)
                                {
                                    W1 = Workbook1.Worksheets[1];
                                }
                            }

                            excel_is_opened = true;
                        }

                    }

                }
                catch (System.Exception ex)
                {
                    Excel1 = new Microsoft.Office.Interop.Excel.Application();

                }


                if (System.IO.File.Exists(fisier_si) == true)
                {
                    if (W1 == null)
                    {
                        if (Excel1.Workbooks.Count == 0) Excel1.Visible = _AGEN_mainform.ExcelVisible;
                        Workbook1 = Excel1.Workbooks.Open(fisier_si);

                        foreach (Microsoft.Office.Interop.Excel.Worksheet W2 in Workbook1.Worksheets)
                        {
                            if (W2.Name.ToLower() == sheetname.ToLower())
                            {
                                W1 = W2;
                            }
                        }
                        if (W1 == null)
                        {
                            W1 = Workbook1.Worksheets[1];
                        }

                    }
                }

                try
                {

                    if (System.IO.File.Exists(fisier_si) == true)
                    {


                        _AGEN_mainform.dt_sheet_index = Functions.Build_Data_table_sheet_index_from_excel(W1, _AGEN_mainform.Start_row_Sheet_index + 1);
                        _AGEN_mainform.tpage_sheetindex.set_dataGridView_sheet_index();
                        if (excel_is_opened == false)
                        {
                            Workbook1.Close();
                        }


                    }

                    string fisier_cl = ProjFolder + _AGEN_mainform.cl_excel_name;
                    Load_centerline_and_station_equation(fisier_cl);

                    if (_AGEN_mainform.dt_centerline.Rows.Count > 0)
                    {
                        Set_centerline_label_to_green();
                        string segment1 = _AGEN_mainform.tpage_setup.Get_segment_name1();
                        if (segment1 == "not defined") segment1 = "";
                        _AGEN_mainform.current_segment = segment1;
                    }
                    else
                    {
                        Set_centerline_label_to_red();
                    }
                    _AGEN_mainform.tpage_st_eq.Populate_datagridview_with_equation_data();


                    if (excel_is_opened == false)
                    {
                        if (Excel1.Workbooks.Count == 0)
                        {
                            Excel1.Quit();
                        }
                        else
                        {
                            Excel1.Visible = true;
                        }

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
                    if (W3 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W3);
                    if (Workbook3 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook3);
                    if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }
            }
            else
            {
                MessageBox.Show("the Project database folder location is not specified\r\n" + ProjFolder + "\r\n operation aborted");

                return;
            }
        }



        private bool Check_to_see_if_there_the_header_in_the_tblk_attributes(Microsoft.Office.Interop.Excel.Worksheet W1, int Start_row)
        {
            Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A7:XX8"];

            object[,] matrix1 = new object[2, 648];
            matrix1 = range1.Value2;
            int end1 = 0;
            for (int i = 1; i <= 648; ++i)
            {
                object val1 = matrix1[1, i];
                if (val1 == null)
                {
                    end1 = i - 1;
                    i = 649;
                }
            }

            if (end1 > 2)
            {
                return true;
            }
            else
            {
                return false;
            }
        }




        private void radioButton_Load_config_CheckedChanged(object sender, EventArgs e)
        {
            Ag = this.MdiParent as _AGEN_mainform;
            if (radioButton_Load_config.Checked == true)
            {


                button_load_config.Visible = true;
            }
            else
            {


                button_load_config.Visible = false;
            }
        }

        private void button_set_project_folder_Click(object sender, EventArgs e)
        {
            Ag = this.MdiParent as _AGEN_mainform;
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                if (System.IO.Directory.Exists(textBox_output_folder.Text) == true)
                {
                    fbd.SelectedPath = textBox_output_folder.Text;
                }


                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    textbox_project_database_folder.Text = fbd.SelectedPath.ToString();

                    DirectoryInfo di = new DirectoryInfo(textbox_project_database_folder.Text);




                    foreach (var File1 in di.GetFiles("*.xlsx", SearchOption.TopDirectoryOnly))
                    {
                        string nume_cu_path = File1.FullName;
                        string nume_fara_path = File1.Name;


                    }



                    button_align_config_saveall_boolean(true);
                    this.WindowState = FormWindowState.Normal;

                    Ag.WindowState = FormWindowState.Normal;

                    MessageBox.Show("done");


                }
            }
        }

        public string get_textBox_client_name_content()
        {
            return textBox_client_name.Text;
        }
        public string get_textBox_project_name_content()
        {
            return textBox_project_name.Text;
        }

        private void button_align_config_saveall_Click(object sender, EventArgs e)
        {
            Ag = this.MdiParent as _AGEN_mainform;
            button_align_config_saveall_boolean(true);

        }
        public void button_align_config_saveall_boolean(bool Close_dwt)
        {
            Ag = this.MdiParent as _AGEN_mainform;
            Functions.Kill_excel();

            string cfg1 = System.IO.Path.GetFileName(_AGEN_mainform.config_path);
            if (Functions.Get_if_workbook_is_open_in_Excel(cfg1) == true)
            {
                MessageBox.Show("Please close the " + cfg1 + " file");
                return;
            }

            _AGEN_mainform.tpage_processing.Show();
            //  Ag.WindowState = FormWindowState.Minimized;

            if (Close_dwt == true) close_template();


            try
            {
                System.Data.DataTable Data_table_config = new System.Data.DataTable();
                Data_table_config.Columns.Add("A", typeof(string));
                Data_table_config.Columns.Add("B", typeof(string));

                if (_AGEN_mainform.lista_segments != null && _AGEN_mainform.lista_segments.Count > 1)
                {
                    for (int i = 1; i < _AGEN_mainform.lista_segments.Count; ++i)
                    {
                        Data_table_config.Columns.Add("seg" + i.ToString(), typeof(string));
                    }
                }


                for (int i = 0; i <= 45; ++i)
                {
                    Data_table_config.Rows.Add();
                }

                Data_table_config.Rows[0][0] = "Client Name";
                Data_table_config.Rows[0][1] = textBox_client_name.Text;


                Data_table_config.Rows[1][0] = "Project Name";
                Data_table_config.Rows[1][1] = textBox_project_name.Text;

                Data_table_config.Rows[2][0] = "Template1";
                Data_table_config.Rows[2][1] = _AGEN_mainform.tpage_viewport_settings.get_template_name_from_text_box();

                Data_table_config.Rows[3][0] = "Template2";
                Data_table_config.Rows[3][1] = _AGEN_mainform.tpage_viewport_settings.get_template_name_from_text_box();

                Data_table_config.Rows[4][0] = "Output folder";

                string out1 = get_output_folder_from_text_box();
                if (out1.Length > 0)
                {
                    if (out1.Substring(out1.Length - 1, 1) != "\\")
                    {
                        out1 = out1 + "\\";
                    }
                }

                Data_table_config.Rows[4][1] = out1;

                Data_table_config.Rows[5][0] = "Prefix File Name";
                Data_table_config.Rows[5][1] = _AGEN_mainform.tpage_viewport_settings.get_prefix_name_from_text_box();

                Data_table_config.Rows[6][0] = "Suffix File Name";
                Data_table_config.Rows[6][1] = _AGEN_mainform.tpage_viewport_settings.get_suffix_name_from_text_box();

                Data_table_config.Rows[7][0] = "Start numbering";
                Data_table_config.Rows[7][1] = _AGEN_mainform.tpage_viewport_settings.get_start_number_from_text_box();

                Data_table_config.Rows[8][0] = "Increment";
                Data_table_config.Rows[8][1] = _AGEN_mainform.tpage_viewport_settings.get_increment_from_text_box();

                Data_table_config.Rows[9][0] = "Country";

                if (radioButton_canada.Checked == true)
                {
                    _AGEN_mainform.COUNTRY = "CANADA";
                    _AGEN_mainform.tpage_sheetindex.set_radioButton_use3D_stations(true);
                    _AGEN_mainform.Project_type = "3D";
                }
                Data_table_config.Rows[9][1] = _AGEN_mainform.COUNTRY;

                Data_table_config.Rows[10][0] = "Left to Right";
                if (_AGEN_mainform.Left_to_Right == true)
                {
                    Data_table_config.Rows[10][1] = "yes";
                }
                else
                {
                    Data_table_config.Rows[10][1] = "no";
                }

                Data_table_config.Rows[11][0] = "Version";
                Data_table_config.Rows[11][1] = _AGEN_mainform.version;

                Data_table_config.Rows[12][0] = "Empty";
                Data_table_config.Rows[12][1] = "";

                Data_table_config.Rows[13][0] = "Empty";
                Data_table_config.Rows[13][1] = "";


                Data_table_config.Rows[14][0] = "Units";


                if (_AGEN_mainform.units_of_measurement == "m")
                {
                    Data_table_config.Rows[14][1] = "meters";
                }
                else
                {
                    Data_table_config.Rows[14][1] = "feet";
                }

                Data_table_config.Rows[15][0] = "Project database folder location";
                string out3 = textbox_project_database_folder.Text;
                if (out3.Length > 0)
                {
                    if (out3.Substring(out3.Length - 1, 1) != "\\")
                    {
                        out3 = out3 + "\\";
                    }
                }
                Data_table_config.Rows[15][1] = out3;
                _AGEN_mainform.ProjFolder = out3;
                Data_table_config.Rows[16][0] = "Empty";
                Data_table_config.Rows[16][1] = "";



                Data_table_config.Rows[17][0] = "North Arrow Block name";
                Data_table_config.Rows[17][1] = "";

                Data_table_config.Rows[18][0] = "North Arrow PS X";
                Data_table_config.Rows[18][1] = "0";

                Data_table_config.Rows[19][0] = "North Arrow PS Y";
                Data_table_config.Rows[19][1] = "0";

                Data_table_config.Rows[20][0] = "North Arrow scale";
                Data_table_config.Rows[20][1] = "1";

                if (_AGEN_mainform.Data_table_blocks != null)
                {
                    if (_AGEN_mainform.Data_table_blocks.Rows.Count > 0)
                    {
                        for (int i = 0; i < _AGEN_mainform.Data_table_blocks.Rows.Count; ++i)
                        {
                            if (_AGEN_mainform.Data_table_blocks.Rows[i][0] != DBNull.Value)
                            {
                                if (_AGEN_mainform.Data_table_blocks.Rows[i]["TYPE"].ToString() == "North Arrow")
                                {

                                    _AGEN_mainform.NA_scale = Convert.ToDouble(_AGEN_mainform.Data_table_blocks.Rows[i]["SCALE"]);
                                    Data_table_config.Rows[17][1] = _AGEN_mainform.Data_table_blocks.Rows[i]["BLOCK_NAME"].ToString();
                                    Data_table_config.Rows[20][1] = _AGEN_mainform.Data_table_blocks.Rows[i]["SCALE"].ToString();
                                    Data_table_config.Rows[18][1] = _AGEN_mainform.Data_table_blocks.Rows[i][_AGEN_mainform.Col_x].ToString();
                                    Data_table_config.Rows[19][1] = _AGEN_mainform.Data_table_blocks.Rows[i][_AGEN_mainform.Col_y].ToString();
                                }
                            }
                        }
                    }
                }
                Data_table_config.Rows[21][0] = "Units precision";
                Data_table_config.Rows[21][1] = _AGEN_mainform.tpage_viewport_settings.get_combobox_units_precision();


                Data_table_config.Rows[22][0] = "Matchline Block Name (in PaperSpace)";
                Data_table_config.Rows[22][1] = _AGEN_mainform.Matchline_BlockName_in_PaperSpace;

                Data_table_config.Rows[23][0] = "Empty";
                Data_table_config.Rows[23][1] = "";

                Data_table_config.Rows[24][0] = "Empty";
                Data_table_config.Rows[24][1] = "";

                Data_table_config.Rows[25][0] = "Empty";
                Data_table_config.Rows[25][1] = "";

                Data_table_config.Rows[26][0] = "Empty";
                Data_table_config.Rows[26][1] = "";

                Data_table_config.Rows[27][0] = "Empty";
                Data_table_config.Rows[27][1] = "";

                Data_table_config.Rows[28][0] = "Empty";
                Data_table_config.Rows[28][1] = "";

                Data_table_config.Rows[29][0] = "Empty";
                Data_table_config.Rows[29][1] = "";

                Data_table_config.Rows[30][0] = "Empty";
                Data_table_config.Rows[30][1] = "";

                Data_table_config.Rows[31][0] = "Empty";
                Data_table_config.Rows[31][1] = "";

                Data_table_config.Rows[32][0] = "Empty";
                Data_table_config.Rows[32][1] = "";

                Data_table_config.Rows[33][0] = "Empty";
                Data_table_config.Rows[33][1] = "";

                Data_table_config.Rows[34][0] = "Empty";
                Data_table_config.Rows[34][1] = "";

                Data_table_config.Rows[35][0] = "Main viewport picked";

                if (_AGEN_mainform.Vw_height > 0 && _AGEN_mainform.Vw_width > 0)
                {
                    _AGEN_mainform.Exista_viewport_main = true;

                }
                else
                {
                    _AGEN_mainform.Exista_viewport_main = false;
                }
                Data_table_config.Rows[35][1] = _AGEN_mainform.Exista_viewport_main.ToString();

                Data_table_config.Rows[36][0] = "Crossing viewport picked";

                if (_AGEN_mainform.Vw_cross_height > 0)
                {
                    _AGEN_mainform.Exista_viewport_cross = true;
                }
                else
                {
                    _AGEN_mainform.Exista_viewport_cross = false;
                }

                Data_table_config.Rows[36][1] = _AGEN_mainform.Exista_viewport_cross.ToString();

                Data_table_config.Rows[37][0] = "Ownership viewport picked";

                if (_AGEN_mainform.Vw_prop_height > 0)
                {
                    _AGEN_mainform.Exista_viewport_owner = true;
                }
                else
                {
                    _AGEN_mainform.Exista_viewport_owner = false;
                }
                Data_table_config.Rows[37][1] = _AGEN_mainform.Exista_viewport_owner.ToString();

                Data_table_config.Rows[38][0] = "Profile viewports picked";

                if (_AGEN_mainform.Vw_prof_height > 0)
                {
                    _AGEN_mainform.Exista_viewport_prof = true;
                }
                else
                {
                    _AGEN_mainform.Exista_viewport_prof = false;
                }
                Data_table_config.Rows[38][1] = _AGEN_mainform.Exista_viewport_prof.ToString();


                Data_table_config.Rows[39][0] = "Project Station Values";
                if (_AGEN_mainform.Project_type == "3D")
                {
                    Data_table_config.Rows[39][1] = "3D";
                }
                else
                {
                    Data_table_config.Rows[39][1] = "2D";
                }

                Data_table_config.Rows[40][0] = "Material viewports picked";

                if (_AGEN_mainform.Vw_mat_height > 0)
                {
                    _AGEN_mainform.Exista_viewport_mat = true;
                }
                else
                {
                    _AGEN_mainform.Exista_viewport_mat = false;
                }
                Data_table_config.Rows[40][1] = _AGEN_mainform.Exista_viewport_mat.ToString();


                Data_table_config.Rows[41][0] = "Profile band viewport picked";
                if (_AGEN_mainform.Vw_profband_height > 0)
                {
                    _AGEN_mainform.Exista_viewport_prof_band = true;
                }
                else
                {
                    _AGEN_mainform.Exista_viewport_prof_band = false;
                }
                Data_table_config.Rows[41][1] = _AGEN_mainform.Exista_viewport_prof_band.ToString();

                Data_table_config.Rows[42][0] = "TBLK band viewport picked";
                if (_AGEN_mainform.Vw_tblk_height > 0)
                {
                    _AGEN_mainform.Exista_viewport_tblk = true;
                }
                else
                {
                    _AGEN_mainform.Exista_viewport_tblk = false;
                }
                Data_table_config.Rows[42][1] = _AGEN_mainform.Exista_viewport_tblk.ToString(); ;

                Data_table_config.Rows[43][0] = "Empty";
                Data_table_config.Rows[43][1] = "";

                Data_table_config.Rows[44][0] = "Empty";
                Data_table_config.Rows[44][1] = "";

                Data_table_config.Rows[45][0] = "Empty";
                Data_table_config.Rows[45][1] = "";

                int first_number = 46;
                int last_number = 99;

                string ProjFolder = textbox_project_database_folder.Text;
                if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                {
                    ProjFolder = ProjFolder + "\\";
                }

                if (System.IO.Directory.Exists(ProjFolder) == false)
                {
                    _AGEN_mainform.tpage_processing.Hide();
                    Ag.WindowState = FormWindowState.Normal;
                    MessageBox.Show("No project folder");
                    return;
                }

                if (_AGEN_mainform.lista_segments != null && _AGEN_mainform.lista_segments.Count > 0)
                {
                    for (int i = 0; i < _AGEN_mainform.lista_segments.Count; ++i)
                    {
                        if (System.IO.Directory.Exists(ProjFolder) == true)
                        {
                            string folder_sgm = ProjFolder + _AGEN_mainform.lista_segments[i];
                            if (System.IO.Directory.Exists(folder_sgm) == false)
                            {
                                System.IO.Directory.CreateDirectory(folder_sgm);
                            }
                        }
                        Data_table_config.Rows.Add();
                        Data_table_config.Rows[Data_table_config.Rows.Count - 1][0] = "Segment " + (i + 1).ToString();
                        Data_table_config.Rows[Data_table_config.Rows.Count - 1][1] = _AGEN_mainform.lista_segments[i];
                        ++first_number;
                    }
                }
                if (first_number < last_number)
                {
                    for (int k = first_number; k <= last_number; ++k)
                    {
                        Data_table_config.Rows.Add();
                        Data_table_config.Rows[Data_table_config.Rows.Count - 1][0] = "";
                        Data_table_config.Rows[Data_table_config.Rows.Count - 1][1] = "";
                    }
                }



                string Scale1 = _AGEN_mainform.tpage_viewport_settings.Get_combobox_viewport_scale_text();

                if (Functions.IsNumeric(Scale1) == true)
                {
                    _AGEN_mainform.Vw_scale = Convert.ToDouble(Scale1);
                }
                else
                {
                    if (Scale1.Contains(":") == true)
                    {
                        Scale1 = Scale1.Replace("1:", "");
                        if (Functions.IsNumeric(Scale1) == true)
                        {
                            _AGEN_mainform.Vw_scale = 1000 / Convert.ToDouble(Scale1);
                        }
                    }
                    else
                    {
                        string inch = "\u0022";

                        if (Scale1.Contains(inch + "=") == true && Scale1.Contains("'") == true)
                        {
                            Scale1 = Scale1.Replace("1" + inch + "=", "");
                            Scale1 = Scale1.Substring(0, Scale1.Length - 1);
                        }

                        inch = "\u0094";

                        if (Scale1.Contains(inch + "=") == true && Scale1.Contains("'") == true)
                        {
                            Scale1 = Scale1.Replace("1" + inch + "=", "");
                            Scale1 = Scale1.Substring(0, Scale1.Length - 1);

                        }

                        if (Functions.IsNumeric(Scale1) == true)
                        {
                            _AGEN_mainform.Vw_scale = 1 / Convert.ToDouble(Scale1);
                        }
                    }
                }


                if (System.IO.File.Exists(_AGEN_mainform.config_path) == true)
                {
                    update_config_file(Data_table_config);
                }
                else
                {
                    save_new_config_file(Data_table_config);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            _AGEN_mainform.tpage_processing.Hide();


            Ag.WindowState = FormWindowState.Normal;
        }

        private void close_template()
        {
            Ag = this.MdiParent as _AGEN_mainform;

            try
            {



                string strTemplatePath = _AGEN_mainform.tpage_viewport_settings.get_template_name_from_text_box();

                DocumentCollection DocumentManager1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
                _AGEN_mainform.tpage_processing.Show();
                // Ag.WindowState = FormWindowState.Minimized;
                foreach (Document Doc in DocumentManager1)
                {
                    if (Doc.Name == strTemplatePath)
                    {

                        Doc.CloseAndDiscard();



                    }

                }
                if (DocumentManager1.Count == 0)
                {
                    string Template1 = "acad.dwt";
                    DocumentManager1.Add(Template1);
                }



            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            _AGEN_mainform.Template_is_open = false;
            _AGEN_mainform.tpage_processing.Hide();
            Ag.WindowState = FormWindowState.Normal;


        }


        public string get_output_folder_from_text_box()
        {
            return textBox_output_folder.Text;
        }




        private void save_new_config_file(System.Data.DataTable Data_table_config)
        {
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

                if (Excel1.Workbooks.Count == 0) Excel1.Visible = _AGEN_mainform.ExcelVisible;

                Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Add();
                Microsoft.Office.Interop.Excel._Worksheet W1 = Workbook1.Worksheets[1];
                W1.Name = "main_cfg";


                try
                {



                    SaveFileDialog Save_dlg = new SaveFileDialog();
                    Save_dlg.Filter = "Excel file|*.xlsx";


                    if (Save_dlg.ShowDialog() == DialogResult.OK)
                    {

                        if (System.IO.File.Exists(Save_dlg.FileName) == false)
                        {
                            string path1 = Save_dlg.FileName;


                            W1.Cells.NumberFormat = "@";

                            int maxRows = Data_table_config.Rows.Count, maxCols = Data_table_config.Columns.Count;
                            Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[1, 1], W1.Cells[maxRows, maxCols]];

                            object[,] values = new object[maxRows, maxCols];
                            for (int row = 0; row < maxRows; row++)
                            {
                                for (int col = 0; col < maxCols; col++)
                                {
                                    if (Data_table_config.Rows[row][col] != DBNull.Value)
                                    {
                                        values[row, col] = Data_table_config.Rows[row][col];
                                    }
                                }
                            }
                            range1.Value2 = values;

                            range1.Columns.AutoFit();

                            transfera_regular_band_to_excel(Workbook1);

                            Ag.Set_textBox_config_file_location(path1);



                            Workbook1.SaveAs(path1);
                        }
                        else
                        {
                            MessageBox.Show("File exists\r\nOperation aborted\r\nSpecify another name.....");
                            return;
                        }

                        Workbook1.Close();
                        if (Excel1.Workbooks.Count == 0)
                        {
                            Excel1.Quit();
                        }
                        else
                        {
                            Excel1.Visible = true;
                        }
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
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }

        }


        private void update_config_file(System.Data.DataTable Data_table_config)
        {
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

                if (Excel1.Workbooks.Count == 0) Excel1.Visible = _AGEN_mainform.ExcelVisible;



                Excel1.Visible = _AGEN_mainform.ExcelVisible;



                if (System.IO.File.Exists(_AGEN_mainform.config_path) == true)
                {

                    Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(_AGEN_mainform.config_path);
                    Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];
                    W1.Name = "main_cfg";
                    try
                    {
                        W1.Cells.NumberFormat = "@";
                        int maxRows = Data_table_config.Rows.Count;
                        int maxCols = Data_table_config.Columns.Count;
                        Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[1, 1], W1.Cells[maxRows, maxCols]];

                        object[,] values = new object[maxRows, maxCols];
                        for (int row = 0; row < maxRows; row++)
                        {
                            for (int col = 0; col < maxCols; col++)
                            {
                                if (Data_table_config.Rows[row][col] != DBNull.Value)
                                {
                                    values[row, col] = Data_table_config.Rows[row][col];
                                }
                            }
                        }
                        range1.Value2 = values;
                        transfera_regular_band_to_excel(Workbook1);


                        Workbook1.Save();
                        Workbook1.Close();
                        if (Excel1.Workbooks.Count == 0)
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

        }




        public string Get_project_database_folder()
        {

            string ProjFolder = textbox_project_database_folder.Text;
            if (System.IO.Directory.Exists(ProjFolder) == false)
            {
                MessageBox.Show("project folder not found or\r\n\"main_cfg\" tab not found", "agen", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return "";
            }

            if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
            {
                ProjFolder = ProjFolder + "\\";
            }

            string Segment1 = _AGEN_mainform.tpage_setup.Get_segment_name1();
            if (Segment1 == "not defined")
            {
                Segment1 = "";
            }
            if (Segment1 != "")
            {
                if (System.IO.Directory.Exists(ProjFolder + Segment1) == false)
                {
                    System.IO.Directory.CreateDirectory(ProjFolder + Segment1);
                }
                Segment1 = Segment1 + "\\";
            }

            return ProjFolder + Segment1;
        }



        public string Get_project_database_folder_without_segment()
        {
            string ProjFolder = textbox_project_database_folder.Text;
            if (System.IO.Directory.Exists(ProjFolder) == false)
            {
                return "";
            }

            if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
            {
                ProjFolder = ProjFolder + "\\";
            }

            return ProjFolder;
        }


        public string Get_client_name()
        {
            return textBox_client_name.Text;
        }

        public string Get_project_name()
        {
            return textBox_project_name.Text;
        }

        public string Get_segment_name1()
        {
            if (comboBox_segment_name.Text != "")
            {
                return comboBox_segment_name.Text;
            }
            return "not defined";
        }

        public void Fill_combobox_segments()
        {
            comboBox_segment_name.Items.Clear();
            if (_AGEN_mainform.lista_segments != null && _AGEN_mainform.lista_segments.Count > 0)
            {
                try
                {
                    for (int i = 0; i < _AGEN_mainform.lista_segments.Count; ++i)
                    {
                        comboBox_segment_name.Items.Add(_AGEN_mainform.lista_segments[i]);
                    }
                    comboBox_segment_name.SelectedIndex = 0;
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }


        public System.Data.DataTable Load_existing_property(string File1)
        {

            if (System.IO.File.Exists(File1) == false)
            {
                MessageBox.Show("the property data file does not exist");
                return null;
            }

            System.Data.DataTable dt2 = new System.Data.DataTable();

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

                if (Excel1.Workbooks.Count == 0) Excel1.Visible = _AGEN_mainform.ExcelVisible;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(File1);
                Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];

                try
                {
                    dt2 = Functions.Build_Data_table_property_from_excel(W1, _AGEN_mainform.Start_row_property + 1, _AGEN_mainform.tpage_sheetindex.get_radioButton_use3D_stations());

                    Workbook1.Save();
                    Workbook1.Close();
                    if (Excel1.Workbooks.Count == 0)
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
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }
            return dt2;

        }
        public System.Data.DataTable Load_existing_sheet_index(string File1, string tabname = "danpopescu")
        {

            if (System.IO.File.Exists(File1) == false)
            {
                MessageBox.Show("the sheet index data file does not exist");
                return null;
            }


            System.Data.DataTable dt2 = new System.Data.DataTable();

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

                if (Excel1.Workbooks.Count == 0) Excel1.Visible = _AGEN_mainform.ExcelVisible;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(File1);
                Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];

                if (tabname != "danpopescu")
                {
                    for (int i = 1; i <= Workbook1.Worksheets.Count; ++i)
                    {
                        if (Workbook1.Worksheets[i].Name == tabname)
                        {
                            W1 = Workbook1.Worksheets[i];
                        }
                    }
                }
                try
                {
                    dt2 = Functions.Build_Data_table_sheet_index_from_excel(W1, _AGEN_mainform.Start_row_Sheet_index + 1);

                    Workbook1.Close();
                    if (Excel1.Workbooks.Count == 0)
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
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }
            return dt2;

        }


        private void button_browse_select_output_folder_Click(object sender, EventArgs e)
        {
            Ag = this.MdiParent as _AGEN_mainform;
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    textBox_output_folder.Text = fbd.SelectedPath.ToString();
                }

            }
        }







        public void Set_output_folder_text_box(string output_f)
        {
            textBox_output_folder.Text = output_f;
        }



        public void Set_centerline_label_to_red()
        {
            label_cl_loaded.Text = "CL not loaded";
            label_cl_loaded.ForeColor = Color.Red;
        }

        public void Set_centerline_label_to_green()
        {
            label_cl_loaded.Text = "CL loaded";
            label_cl_loaded.ForeColor = Color.LimeGreen;
        }


        public List<int> create_band_list_indexes_for_generation(Point3d inspt, double band_spacing, string layer_no_plot)
        {

            List<int> lista1 = new List<int>();
            if (_AGEN_mainform.dt_sheet_index != null)
            {
                if (_AGEN_mainform.dt_sheet_index.Rows.Count > 0)
                {
                    for (int j = 0; j < _AGEN_mainform.dt_sheet_index.Rows.Count; ++j)
                    {
                        lista1.Add(j);
                    }

                    ObjectId[] Empty_array = null;
                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;

                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {

                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;
                        foreach (ObjectId id1 in BTrecord)
                        {
                            Polyline rect1 = Trans1.GetObject(id1, OpenMode.ForRead) as Polyline;
                            if (rect1 != null)
                            {
                                if (rect1.Closed == true)
                                {
                                    if (rect1.Layer == layer_no_plot)
                                    {
                                        if (rect1.ColorIndex == 3)
                                        {
                                            if (rect1.NumberOfVertices == 4)
                                            {
                                                Point3d pt1 = rect1.GetPointAtParameter(0);
                                                Point3d pt2 = rect1.GetPointAtParameter(1);
                                                if (Math.Round(pt1.DistanceTo(pt2), 0) == Math.Round(_AGEN_mainform.Vw_width, 0))
                                                {
                                                    Point3d ptm = new Point3d((pt1.X + pt2.X) / 2, (pt1.Y + pt2.Y) / 2, 0);

                                                    double no_rand = (inspt.Y - ptm.Y) / band_spacing;
                                                    if (Math.Abs(Math.Round(no_rand, 0) - no_rand) < 0.01 && Math.Abs(inspt.X - ptm.X) < 0.01)
                                                    {
                                                        int nr_rand = Convert.ToInt32(Math.Round(no_rand, 0));
                                                        if (lista1.Contains(nr_rand) == true)
                                                        {
                                                            lista1.Remove(nr_rand);
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
                    Editor1.SetImpliedSelection(Empty_array);

                }



            }
            return lista1;
        }

        public List<int> create_band_list_of_dwg(string start1, string end1)
        {
            List<int> lista1 = new List<int>();
            if (_AGEN_mainform.dt_sheet_index != null)
            {
                if (_AGEN_mainform.dt_sheet_index.Rows.Count > 0)
                {
                    bool adauga = false;
                    for (int i = 0; i < _AGEN_mainform.dt_sheet_index.Rows.Count; ++i)
                    {
                        string nume1 = Convert.ToString(_AGEN_mainform.dt_sheet_index.Rows[i]["DwgNo"]);
                        if (nume1.ToUpper() == start1.ToUpper() || start1 == "" && end1 == "")
                        {
                            adauga = true;
                        }

                        if (adauga == true)
                        {
                            lista1.Add(i);
                        }
                        if (nume1.ToUpper() == end1.ToUpper())
                        {
                            adauga = false;
                        }

                    }
                }
            }
            return lista1;
        }

        public void display_checkboxes_into_generation_page()
        {
            if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count > 0)
            {
                _AGEN_mainform.tpage_sheet_gen.show_panel_custom_bands();

                for (int i = 0; i < _AGEN_mainform.Data_Table_custom_bands.Rows.Count; ++i)
                {
                    if (i == 0)
                    {
                        _AGEN_mainform.tpage_sheet_gen.show_checkBox1();
                        _AGEN_mainform.tpage_sheet_gen.set_txt_checkBox1(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"].ToString());
                    }
                    else if (i == 1)
                    {
                        _AGEN_mainform.tpage_sheet_gen.show_checkBox2();
                        _AGEN_mainform.tpage_sheet_gen.set_txt_checkBox2(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"].ToString());
                    }
                    else if (i == 2)
                    {
                        _AGEN_mainform.tpage_sheet_gen.show_checkBox3();
                        _AGEN_mainform.tpage_sheet_gen.set_txt_checkBox3(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"].ToString());
                    }
                    else if (i == 3)
                    {
                        _AGEN_mainform.tpage_sheet_gen.show_checkBox4();
                        _AGEN_mainform.tpage_sheet_gen.set_txt_checkBox4(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"].ToString());
                    }
                    else if (i == 4)
                    {
                        _AGEN_mainform.tpage_sheet_gen.show_checkBox5();
                        _AGEN_mainform.tpage_sheet_gen.set_txt_checkBox5(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"].ToString());
                    }
                    else if (i == 5)
                    {
                        _AGEN_mainform.tpage_sheet_gen.show_checkBox6();
                        _AGEN_mainform.tpage_sheet_gen.set_txt_checkBox6(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"].ToString());
                    }
                    else if (i == 6)
                    {
                        _AGEN_mainform.tpage_sheet_gen.show_checkBox7();
                        _AGEN_mainform.tpage_sheet_gen.set_txt_checkBox7(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"].ToString());
                    }
                    else if (i == 7)
                    {
                        _AGEN_mainform.tpage_sheet_gen.show_checkBox8();
                        _AGEN_mainform.tpage_sheet_gen.set_txt_checkBox8(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"].ToString());
                    }
                    else if (i == 8)
                    {
                        _AGEN_mainform.tpage_sheet_gen.show_checkBox9();
                        _AGEN_mainform.tpage_sheet_gen.set_txt_checkBox9(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"].ToString());
                    }
                    else if (i == 9)
                    {
                        _AGEN_mainform.tpage_sheet_gen.show_checkBox10();
                        _AGEN_mainform.tpage_sheet_gen.set_txt_checkBox10(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"].ToString());
                    }
                    else if (i == 10)
                    {
                        _AGEN_mainform.tpage_sheet_gen.show_checkBox11();
                        _AGEN_mainform.tpage_sheet_gen.set_txt_checkBox11(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"].ToString());
                    }
                    else if (i == 11)
                    {
                        _AGEN_mainform.tpage_sheet_gen.show_checkBox12();
                        _AGEN_mainform.tpage_sheet_gen.set_txt_checkBox12(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"].ToString());
                    }
                    else if (i == 12)
                    {
                        _AGEN_mainform.tpage_sheet_gen.show_checkBox13();
                        _AGEN_mainform.tpage_sheet_gen.set_txt_checkBox13(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"].ToString());
                    }
                    else if (i == 13)
                    {
                        _AGEN_mainform.tpage_sheet_gen.show_checkBox14();
                        _AGEN_mainform.tpage_sheet_gen.set_txt_checkBox14(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"].ToString());
                    }
                    else if (i == 14)
                    {
                        _AGEN_mainform.tpage_sheet_gen.show_checkBox15();
                        _AGEN_mainform.tpage_sheet_gen.set_txt_checkBox15(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"].ToString());
                    }
                    else if (i == 15)
                    {
                        _AGEN_mainform.tpage_sheet_gen.show_checkBox16();
                        _AGEN_mainform.tpage_sheet_gen.set_txt_checkBox16(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"].ToString());
                    }
                }
                if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count < 16)
                {
                    for (int i = _AGEN_mainform.Data_Table_custom_bands.Rows.Count; i < 17; ++i)
                    {
                        if (i == 1) _AGEN_mainform.tpage_sheet_gen.Hide_checkBox2();
                        else if (i == 2) _AGEN_mainform.tpage_sheet_gen.Hide_checkBox3();
                        else if (i == 3) _AGEN_mainform.tpage_sheet_gen.Hide_checkBox4();
                        else if (i == 4) _AGEN_mainform.tpage_sheet_gen.Hide_checkBox5();
                        else if (i == 5) _AGEN_mainform.tpage_sheet_gen.Hide_checkBox6();
                        else if (i == 6) _AGEN_mainform.tpage_sheet_gen.Hide_checkBox7();
                        else if (i == 7) _AGEN_mainform.tpage_sheet_gen.Hide_checkBox8();
                        else if (i == 8) _AGEN_mainform.tpage_sheet_gen.Hide_checkBox9();
                        else if (i == 9) _AGEN_mainform.tpage_sheet_gen.Hide_checkBox10();
                        else if (i == 10) _AGEN_mainform.tpage_sheet_gen.Hide_checkBox11();
                        else if (i == 11) _AGEN_mainform.tpage_sheet_gen.Hide_checkBox12();
                        else if (i == 12) _AGEN_mainform.tpage_sheet_gen.Hide_checkBox13();
                        else if (i == 13) _AGEN_mainform.tpage_sheet_gen.Hide_checkBox14();
                        else if (i == 14) _AGEN_mainform.tpage_sheet_gen.Hide_checkBox15();
                        else if (i == 15) _AGEN_mainform.tpage_sheet_gen.Hide_checkBox16();
                    }

                }


            }
            else
            {
                _AGEN_mainform.tpage_sheet_gen.Hide_panel_custom_bands();
            }


            if (_AGEN_mainform.Data_Table_extra_mainVP != null && _AGEN_mainform.Data_Table_extra_mainVP.Rows.Count > 0)
            {
                _AGEN_mainform.tpage_sheet_gen.show_panel_extra_bands();

                if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count > 0)
                {
                    _AGEN_mainform.tpage_sheet_gen.show_checkBox_extra1();
                }
                else
                {
                    _AGEN_mainform.tpage_sheet_gen.Hide_checkBox_extra1();
                }

                if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count > 1)
                {
                    _AGEN_mainform.tpage_sheet_gen.show_checkBox_extra2();
                }
                else
                {
                    _AGEN_mainform.tpage_sheet_gen.Hide_checkBox_extra2();
                }

                if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count > 2)
                {
                    _AGEN_mainform.tpage_sheet_gen.show_checkBox_extra3();
                }
                else
                {
                    _AGEN_mainform.tpage_sheet_gen.Hide_checkBox_extra3();
                }

                if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count > 3)
                {
                    _AGEN_mainform.tpage_sheet_gen.show_checkBox_extra4();
                }
                else
                {
                    _AGEN_mainform.tpage_sheet_gen.Hide_checkBox_extra4();
                }

                if (_AGEN_mainform.Data_Table_extra_mainVP.Rows.Count > 4)
                {
                    _AGEN_mainform.tpage_sheet_gen.show_checkBox_extra5();
                }
                else
                {
                    _AGEN_mainform.tpage_sheet_gen.Hide_checkBox_extra5();
                }

            }
            else
            {
                _AGEN_mainform.tpage_sheet_gen.Hide_panel_extra_bands();
            }

        }

        public void transfera_custom_band_to_excel(Microsoft.Office.Interop.Excel.Workbook Workbook1)
        {
            if (_AGEN_mainform.Data_Table_custom_bands != null)
            {
                if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count > 0)
                {
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
                }
            }
        }
        public void transfera_regular_band_to_excel(Microsoft.Office.Interop.Excel.Workbook Workbook1)
        {
            if (_AGEN_mainform.Data_Table_regular_bands != null)
            {
                if (_AGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                {
                    for (int i = 0; i < _AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                    {
                        if (_AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] != DBNull.Value)
                        {
                            string bn = Convert.ToString(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"]);
                            if (bn == _AGEN_mainform.nume_main_vp || bn == _AGEN_mainform.nume_banda_prof || bn == _AGEN_mainform.nume_banda_profband)
                            {
                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["Custom_scale"] = _AGEN_mainform.Vw_scale;
                            }

                        }
                    }


                    Microsoft.Office.Interop.Excel.Worksheet W1 = null;

                    foreach (Microsoft.Office.Interop.Excel.Worksheet wsh1 in Workbook1.Worksheets)
                    {
                        if (wsh1.Name == "Regular_band_data")
                        {
                            W1 = wsh1;
                        }
                    }

                    if (W1 == null)
                    {
                        W1 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[Workbook1.Worksheets.Count], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                        W1.Name = "Regular_band_data";

                    }

                    W1.Columns["A:XX"].Delete();
                    W1.Range["A:S"].ColumnWidth = 18;
                    W1.Range["C:K"].ColumnWidth = 2;
                    int maxRows = _AGEN_mainform.Data_Table_regular_bands.Rows.Count;
                    int maxCols = _AGEN_mainform.Data_Table_regular_bands.Columns.Count;

                    Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[2, 1], W1.Cells[maxRows + 1, maxCols]];
                    object[,] values1 = new object[maxRows, maxCols];

                    for (int i = 0; i < maxRows; ++i)
                    {
                        for (int j = 0; j < maxCols; ++j)
                        {
                            if (_AGEN_mainform.Data_Table_regular_bands.Rows[i][j] != DBNull.Value)
                            {
                                values1[i, j] = _AGEN_mainform.Data_Table_regular_bands.Rows[i][j];
                            }
                        }
                    }

                    for (int i = 0; i < _AGEN_mainform.Data_Table_regular_bands.Columns.Count; ++i)
                    {
                        W1.Cells[1, i + 1].value2 = _AGEN_mainform.Data_Table_regular_bands.Columns[i].ColumnName;
                    }

                    range1.Cells.NumberFormat = "@";
                    range1.Value2 = values1;

                    Functions.Color_border_range_inside(range1, 0);

                }
            }

        }

        public void transfera_band_settings_to_config_excel(System.Data.DataTable dt_dwg_data_ownership, System.Data.DataTable dt_dwg_data_crossing)
        {
            if (_AGEN_mainform.Data_Table_regular_bands != null)
            {
                if (_AGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                {

                    if (System.IO.File.Exists(_AGEN_mainform.config_path) == false)
                    {
                        MessageBox.Show("config file not found");
                        return;
                    }

                    Microsoft.Office.Interop.Excel.Application Excel1 = null;

                    try
                    {
                        Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    }
                    catch (System.Exception ex)
                    {
                        Excel1 = new Microsoft.Office.Interop.Excel.Application();

                    }

                    if (Excel1.Workbooks.Count == 0) Excel1.Visible = _AGEN_mainform.ExcelVisible;
                    Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(_AGEN_mainform.config_path);

                    Microsoft.Office.Interop.Excel.Worksheet W1 = null;
                    Microsoft.Office.Interop.Excel.Worksheet W2 = null;
                    Microsoft.Office.Interop.Excel.Worksheet W3 = null;


                    string segment1 = _AGEN_mainform.tpage_setup.Get_segment_name1();
                    if (segment1 == "not defined") segment1 = "";
                    try
                    {
                        foreach (Microsoft.Office.Interop.Excel.Worksheet wsh1 in Workbook1.Worksheets)
                        {
                            if (wsh1.Name == "Regular_band_data")
                            {
                                W1 = wsh1;
                            }
                            if (wsh1.Name == "Ownership_dwg_data_" + segment1)
                            {
                                W2 = wsh1;
                            }
                            if (wsh1.Name == "Crossing_dwg_data_" + segment1)
                            {
                                W3 = wsh1;
                            }

                        }

                        if (W1 == null)
                        {
                            W1 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[Workbook1.Worksheets.Count], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                            W1.Name = "Regular_band_data";

                        }

                        W1.Columns["A:XX"].Delete();
                        int maxRows = _AGEN_mainform.Data_Table_regular_bands.Rows.Count;
                        int maxCols = _AGEN_mainform.Data_Table_regular_bands.Columns.Count;

                        Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[2, 1], W1.Cells[maxRows + 1, maxCols]];
                        object[,] values1 = new object[maxRows, maxCols];

                        for (int i = 0; i < maxRows; ++i)
                        {
                            for (int j = 0; j < maxCols; ++j)
                            {
                                if (_AGEN_mainform.Data_Table_regular_bands.Rows[i][j] != DBNull.Value)
                                {
                                    values1[i, j] = _AGEN_mainform.Data_Table_regular_bands.Rows[i][j];
                                }
                            }
                        }

                        for (int i = 0; i < _AGEN_mainform.Data_Table_regular_bands.Columns.Count; ++i)
                        {
                            W1.Cells[1, i + 1].value2 = _AGEN_mainform.Data_Table_regular_bands.Columns[i].ColumnName;
                        }

                        range1.Cells.NumberFormat = "@";
                        range1.Value2 = values1;

                        Functions.Color_border_range_inside(range1, 0);

                        bool exista_dwg_data_for_config = false;

                        if (dt_dwg_data_ownership != null && dt_dwg_data_ownership.Rows.Count > 0)
                        {
                            exista_dwg_data_for_config = true;
                        }

                        if (dt_dwg_data_crossing != null && dt_dwg_data_crossing.Rows.Count > 0)
                        {
                            exista_dwg_data_for_config = true;
                        }


                        if (exista_dwg_data_for_config == true)
                        {

                            if (dt_dwg_data_ownership != null && dt_dwg_data_ownership.Rows.Count > 0)
                            {
                                if (W2 == null)
                                {
                                    W2 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[Workbook1.Worksheets.Count], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                                    W2.Name = "Ownership_dwg_data_" + segment1;

                                }

                                W2.Columns["A:XX"].Delete();
                                maxRows = dt_dwg_data_ownership.Rows.Count;
                                maxCols = dt_dwg_data_ownership.Columns.Count;

                                range1 = W2.Range[W2.Cells[2, 1], W2.Cells[maxRows + 1, maxCols]];
                                values1 = new object[maxRows, maxCols];

                                for (int i = 0; i < maxRows; ++i)
                                {
                                    for (int j = 0; j < maxCols; ++j)
                                    {
                                        if (dt_dwg_data_ownership.Rows[i][j] != DBNull.Value)
                                        {
                                            values1[i, j] = dt_dwg_data_ownership.Rows[i][j];
                                        }
                                    }
                                }

                                for (int i = 0; i < dt_dwg_data_ownership.Columns.Count; ++i)
                                {
                                    W2.Cells[1, i + 1].value2 = dt_dwg_data_ownership.Columns[i].ColumnName;
                                }

                                range1.Cells.NumberFormat = "@";
                                range1.Value2 = values1;

                                Functions.Color_border_range_inside(range1, 0);
                            }

                            if (dt_dwg_data_crossing != null && dt_dwg_data_crossing.Rows.Count > 0)
                            {
                                if (W3 == null)
                                {
                                    W3 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[Workbook1.Worksheets.Count], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                                    W3.Name = "Crossing_dwg_data_" + segment1;

                                }

                                W3.Columns["A:XX"].Delete();
                                maxRows = dt_dwg_data_crossing.Rows.Count;
                                maxCols = dt_dwg_data_crossing.Columns.Count;

                                range1 = W3.Range[W3.Cells[2, 1], W3.Cells[maxRows + 1, maxCols]];
                                values1 = new object[maxRows, maxCols];

                                for (int i = 0; i < maxRows; ++i)
                                {
                                    for (int j = 0; j < maxCols; ++j)
                                    {
                                        if (dt_dwg_data_crossing.Rows[i][j] != DBNull.Value)
                                        {
                                            values1[i, j] = dt_dwg_data_crossing.Rows[i][j];
                                        }
                                    }
                                }

                                for (int i = 0; i < dt_dwg_data_crossing.Columns.Count; ++i)
                                {
                                    W3.Cells[1, i + 1].value2 = dt_dwg_data_crossing.Columns[i].ColumnName;
                                }

                                range1.Cells.NumberFormat = "@";
                                range1.Value2 = values1;

                                Functions.Color_border_range_inside(range1, 0);
                            }



                        }




                        Workbook1.Save();
                        Workbook1.Close();
                        if (Excel1.Workbooks.Count == 0)
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
                        if (W2 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W2);
                        if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                        if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                    }


                }
            }

        }

        public System.Data.DataTable Load_existing_station_equations(string File1)
        {

            if (System.IO.File.Exists(File1) == false)
            {
                MessageBox.Show("the station equations data file does not exist");
                return null;
            }


            _AGEN_mainform.dt_station_equation = Functions.Creaza_station_equation_datatable_structure();

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

                if (Excel1.Workbooks.Count == 0) Excel1.Visible = _AGEN_mainform.ExcelVisible;

                Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(File1);
                Microsoft.Office.Interop.Excel.Worksheet W1 = null;

                if (Workbook1.Sheets.Count >= 2)
                {
                    try
                    {
                        W1 = Workbook1.Worksheets[2];
                        if (Convert.ToString(W1.Range["C1"].Value2) == "Station Equations")
                        {
                            _AGEN_mainform.dt_station_equation = Functions.Build_Data_table_station_Equation_from_excel(W1, _AGEN_mainform.Start_row_station_equation + 1);
                        }

                        Workbook1.Close();
                        if (Excel1.Workbooks.Count == 0)
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


            return _AGEN_mainform.dt_station_equation;

        }


        public void Load_centerline_and_station_equation(string File1)
        {

            _AGEN_mainform.dt_station_equation = Functions.Creaza_station_equation_datatable_structure();
            try
            {
                Microsoft.Office.Interop.Excel.Application Excel1 = null;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;
                Microsoft.Office.Interop.Excel.Worksheet W1 = null;
                Microsoft.Office.Interop.Excel.Worksheet W2 = null;
                bool excel_is_opened = false;

                try
                {
                    Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    foreach (Microsoft.Office.Interop.Excel.Workbook Workbook2 in Excel1.Workbooks)
                    {
                        if (Workbook2.FullName.ToLower() == File1.ToLower())
                        {
                            Workbook1 = Workbook2;
                            W1 = Workbook1.Worksheets[1];
                            excel_is_opened = true;
                        }

                    }
                }
                catch (System.Exception ex)
                {
                    Excel1 = new Microsoft.Office.Interop.Excel.Application();

                }

                if (W1 == null)
                {
                    if (Excel1.Workbooks.Count == 0) Excel1.Visible = _AGEN_mainform.ExcelVisible;
                    Workbook1 = Excel1.Workbooks.Open(File1);
                    W1 = Workbook1.Worksheets[1];
                }

                try
                {
                    bool CSF = false;
                    if (_AGEN_mainform.COUNTRY == "CANADA") CSF = true;
                    _AGEN_mainform.dt_centerline = Functions.Build_Data_table_centerline_from_excel(W1, _AGEN_mainform.Start_row_CL + 1, CSF);
                    System.Data.DataTable dt_cl = _AGEN_mainform.dt_centerline;
                    System.Data.DataTable dt_seq = null;
                    if (_AGEN_mainform.dt_centerline.Rows.Count > 0)
                    {
                        Set_centerline_label_to_green();
                        if (Workbook1.Sheets.Count > 1)
                        {
                            if (Workbook1.Worksheets[2].Name == "St_eq")
                            {
                                W2 = Workbook1.Worksheets[2];
                                _AGEN_mainform.dt_station_equation = Functions.Build_Data_table_station_Equation_from_excel(W2, _AGEN_mainform.Start_row_station_equation + 1);
                            }
                        }

                        if (_AGEN_mainform.COUNTRY == "CANADA")
                        {
                            if (dt_cl != null && dt_cl.Rows.Count > 0)
                            {
                                for (int i = 0; i < dt_cl.Rows.Count; ++i)
                                {
                                    if (dt_cl.Rows[i][Col_BackSta] != DBNull.Value && dt_cl.Rows[i][Col_AheadSta] != DBNull.Value)
                                    {
                                        if (dt_seq == null)
                                        {
                                            dt_seq = Functions.Creaza_station_equation_datatable_structure();
                                        }

                                        System.Data.DataRow row1 = dt_seq.NewRow();
                                        row1[sta_back] = dt_cl.Rows[i][Col_BackSta];
                                        row1[sta_ahead] = dt_cl.Rows[i][Col_AheadSta];
                                        row1[rr_end_x] = dt_cl.Rows[i][Col_x];
                                        row1[rr_end_y] = dt_cl.Rows[i][Col_y];
                                        row1[rr_end_z] = dt_cl.Rows[i][Col_z];
                                        row1[version] = _AGEN_mainform.current_segment;
                                        row1[show_in_plan] = "YES";
                                        dt_seq.Rows.Add(row1);
                                    }
                                }
                            }
                            _AGEN_mainform.dt_station_equation = dt_seq;
                        }

                    }
                    else
                    {
                        Set_centerline_label_to_red();
                    }

                    if (excel_is_opened == false)
                    {
                        Workbook1.Close();
                        if (Excel1.Workbooks.Count == 0)
                        {
                            Excel1.Quit();
                        }
                        else
                        {
                            Excel1.Visible = true;
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);
                }
                finally
                {
                    if (W1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W1);
                    if (W2 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W2);
                    if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                    if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
            //Functions.Transfer_datatable_to_new_excel_spreadsheet(_AGEN_mainform.dt_centerline, "cl1");

        }







        private System.Data.DataTable Load_profile_poly_from_excel_config(Microsoft.Office.Interop.Excel.Worksheet W1)
        {
            System.Data.DataTable dt1 = new System.Data.DataTable();

            dt1.Columns.Add("x", typeof(double));
            dt1.Columns.Add("y", typeof(double));
            dt1.Columns.Add("sta", typeof(double));

            Microsoft.Office.Interop.Excel.Range range2 = W1.Range["A2:A30001"];

            object[,] matrix2 = new object[1, 30000];
            matrix2 = range2.Value2;



            int row_end1 = 1;
            for (int i = 1; i <= 30000; ++i)
            {
                object val1 = matrix2[i, 1];

                if (val1 == null)
                {
                    row_end1 = i;
                    i = 30001;
                }
            }


            if (row_end1 == 1)
            {
                MessageBox.Show("no profile drafted\r\nOperation aborted");
                return null;
            }


            Microsoft.Office.Interop.Excel.Range range3 = W1.Range[W1.Cells[1, 1], W1.Cells[row_end1, 3]];

            object[,] matrix3 = new object[row_end1, 3];
            matrix3 = range3.Value2;



            for (int i = 2; i <= row_end1; ++i)
            {
                dt1.Rows.Add();

                for (int j = 1; j <= 3; ++j)
                {
                    object val1 = matrix3[i, j];
                    if (val1 != null)
                    {
                        string continut = val1.ToString();
                        if (Functions.IsNumeric(continut) == true)
                        {
                            dt1.Rows[dt1.Rows.Count - 1][j - 1] = Convert.ToDouble(continut);
                        }
                    }
                }

            }


            //Functions.Transfer_datatable_to_new_excel_spreadsheet(dt1);


            return dt1;
        }







        private void button_load_CL_to_excel_Click(object sender, EventArgs e)
        {
            Ag = this.MdiParent as _AGEN_mainform;

            if (System.IO.File.Exists(_AGEN_mainform.config_path) == true)
            {

                string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                {
                    ProjF = ProjF + "\\";
                }
                string fisier_cl = ProjF + _AGEN_mainform.cl_excel_name;
                if (System.IO.File.Exists(fisier_cl) == true)
                {
                    if (MessageBox.Show("all data from centerline.xls will be overwriten... \r\nDo you want to continue?", "agen", MessageBoxButtons.YesNo) == DialogResult.No)
                    {
                        return;
                    }
                }
            }
            else
            {
                MessageBox.Show("Please save your project before you load the centerline");
                Set_centerline_label_to_red();
                return;
            }

            Functions.Kill_excel();

            if (Functions.Get_if_workbook_is_open_in_Excel("centerline.xlsx") == true)
            {
                MessageBox.Show("Please close the centerline file");
                Set_centerline_label_to_red();
                return;
            }

            Point3d pt0 = new Point3d();
            double poly_length = 0;

            _AGEN_mainform.tpage_processing.Show();
            Ag.WindowState = FormWindowState.Minimized;
            set_enable_false();

            ObjectId[] Empty_array = null;

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_optionsCL = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect Centerline:");
                        Prompt_optionsCL.SetRejectMessage("\nYou did not selected a polyline (2d or 3d)");
                        Prompt_optionsCL.AddAllowedClass(typeof(Polyline), true);
                        Prompt_optionsCL.AddAllowedClass(typeof(Polyline3d), true);

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_CL = Editor1.GetEntity(Prompt_optionsCL);
                        if (Rezultat_CL.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            _AGEN_mainform.tpage_processing.Hide();
                            set_enable_true();
                            Editor1.SetImpliedSelection(Empty_array);
                            Editor1.WriteMessage("\nCommand:");
                            Ag.WindowState = FormWindowState.Normal;
                            Set_centerline_label_to_red();
                            return;
                        }

                        Ag.WindowState = FormWindowState.Normal;
                        Curve Curba1 = Trans1.GetObject(Rezultat_CL.ObjectId, OpenMode.ForRead) as Curve;

                        if (Curba1 == null)
                        {
                            MessageBox.Show("you did not select a polyline or a polyline3d");
                            _AGEN_mainform.tpage_processing.Hide();
                            set_enable_true();
                            Editor1.SetImpliedSelection(Empty_array);
                            Editor1.WriteMessage("\nCommand:");
                            Ag.WindowState = FormWindowState.Normal;
                            Set_centerline_label_to_red();
                            return;
                        }


                        if (Curba1 is Polyline)
                        {
                            _AGEN_mainform.Poly2D = (Polyline)Curba1;
                            _AGEN_mainform.Poly3D = null;
                            _AGEN_mainform.tpage_sheetindex.set_radioButton_use3D_stations(false);
                            _AGEN_mainform.Project_type = "2D";
                            poly_length = _AGEN_mainform.Poly2D.Length;
                            pt0 = _AGEN_mainform.Poly2D.StartPoint;
                        }

                        else if (Curba1 is Polyline3d)
                        {
                            _AGEN_mainform.Poly3D = (Polyline3d)Curba1;
                            _AGEN_mainform.Poly2D = Functions.Build_2dpoly_from_3d(_AGEN_mainform.Poly3D);
                            _AGEN_mainform.tpage_sheetindex.set_radioButton_use3D_stations(true);
                            _AGEN_mainform.Project_type = "3D";
                            poly_length = _AGEN_mainform.Poly3D.Length;
                            pt0 = _AGEN_mainform.Poly3D.StartPoint;
                        }

                        else
                        {
                            MessageBox.Show("you did not select a polyline or a polyline3d");
                            _AGEN_mainform.tpage_processing.Hide();
                            set_enable_true();
                            Editor1.SetImpliedSelection(Empty_array);
                            Editor1.WriteMessage("\nCommand:");
                            Ag.WindowState = FormWindowState.Normal;
                            Set_centerline_label_to_red();
                            return;
                        }


                        BlockTable BlockTable1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);

                        Autodesk.AutoCAD.DatabaseServices.LayerTable LayerTable1 = (Autodesk.AutoCAD.DatabaseServices.LayerTable)Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);


                        _AGEN_mainform.dt_centerline = Functions.Creaza_centerline_datatable_structure();



                        for (int i = 0; i < _AGEN_mainform.Poly2D.NumberOfVertices; ++i)
                        {

                            double x2 = _AGEN_mainform.Poly2D.GetPointAtParameter(i).X;
                            double y2 = _AGEN_mainform.Poly2D.GetPointAtParameter(i).Y;
                            double z2 = _AGEN_mainform.Poly2D.GetPointAtParameter(i).Z;
                            double bulge = _AGEN_mainform.Poly2D.GetBulgeAt(i);

                            if (_AGEN_mainform.Poly3D != null)
                            {
                                z2 = _AGEN_mainform.Poly3D.GetPointAtParameter(i).Z;
                            }

                            _AGEN_mainform.dt_centerline.Rows.Add();
                            _AGEN_mainform.dt_centerline.Rows[_AGEN_mainform.dt_centerline.Rows.Count - 1][_AGEN_mainform.Col_x] = x2;
                            _AGEN_mainform.dt_centerline.Rows[_AGEN_mainform.dt_centerline.Rows.Count - 1][_AGEN_mainform.Col_y] = y2;
                            _AGEN_mainform.dt_centerline.Rows[_AGEN_mainform.dt_centerline.Rows.Count - 1][_AGEN_mainform.Col_z] = z2;
                            _AGEN_mainform.dt_centerline.Rows[_AGEN_mainform.dt_centerline.Rows.Count - 1][_AGEN_mainform.Col_2DSta] = _AGEN_mainform.Poly2D.GetDistanceAtParameter(i);

                            if (bulge != 0) _AGEN_mainform.dt_centerline.Rows[_AGEN_mainform.dt_centerline.Rows.Count - 1][_AGEN_mainform.Col_MMid] = bulge;


                            if (_AGEN_mainform.Poly3D != null)
                            {
                                _AGEN_mainform.dt_centerline.Rows[_AGEN_mainform.dt_centerline.Rows.Count - 1][_AGEN_mainform.Col_3DSta] = _AGEN_mainform.Poly3D.GetDistanceAtParameter(i);
                                z2 = _AGEN_mainform.Poly3D.GetPointAtParameter(i).Z;
                                _AGEN_mainform.dt_centerline.Rows[_AGEN_mainform.dt_centerline.Rows.Count - 1][_AGEN_mainform.Col_z] = z2;
                            }


                            if (i > 0 && i < _AGEN_mainform.Poly2D.NumberOfVertices - 1)
                            {
                                double x1 = _AGEN_mainform.Poly2D.GetPointAtParameter(i - 1).X;
                                double y1 = _AGEN_mainform.Poly2D.GetPointAtParameter(i - 1).Y;
                                double x3 = _AGEN_mainform.Poly2D.GetPointAtParameter(i + 1).X;
                                double y3 = _AGEN_mainform.Poly2D.GetPointAtParameter(i + 1).Y;

                                string Deflexia = Functions.Get_deflection_angle_dms(x1, y1, x2, y2, x3, y3);
                                double Defl1 = 180 * Functions.Get_deflection_angle_rad(x1, y1, x2, y2, x3, y3) / Math.PI;

                                if (bulge == 0) _AGEN_mainform.dt_centerline.Rows[_AGEN_mainform.dt_centerline.Rows.Count - 1][_AGEN_mainform.Col_DeflAng] = Defl1;
                                if (bulge == 0) _AGEN_mainform.dt_centerline.Rows[_AGEN_mainform.dt_centerline.Rows.Count - 1][_AGEN_mainform.Col_DeflAngDMS] = Deflexia;



                            }




                            if (i < _AGEN_mainform.Poly2D.NumberOfVertices - 1)
                            {
                                double x3 = _AGEN_mainform.Poly2D.GetPointAtParameter(i + 1).X;
                                double y3 = _AGEN_mainform.Poly2D.GetPointAtParameter(i + 1).Y;
                                if (bulge == 0) _AGEN_mainform.dt_centerline.Rows[_AGEN_mainform.dt_centerline.Rows.Count - 1][_AGEN_mainform.Col_Bearing] = Functions.Get_Quadrant_bearing(Functions.GET_Bearing_rad(x2, y2, x3, y3));
                                if (bulge == 0) _AGEN_mainform.dt_centerline.Rows[_AGEN_mainform.dt_centerline.Rows.Count - 1][_AGEN_mainform.Col_Distance] = Math.Pow((x2 - x3) * (x2 - x3) + (y2 - y3) * (y2 - y3), 0.5);

                            }


                        }
                        Trans1.Commit();
                    }
                }

                string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                {
                    ProjFolder = ProjFolder + "\\";
                }
                if (System.IO.Directory.Exists(ProjFolder) == true)
                {
                    double start1 = 0;


                    string fisier_cl = ProjFolder + _AGEN_mainform.cl_excel_name;

                    string Start_sta_cl = textBox_start_station_CL.Text;

                    Set_centerline_label_to_green();

                    if (Functions.IsNumeric(Start_sta_cl.Replace("+", "")) == true)
                    {
                        _AGEN_mainform.dt_station_equation = Functions.Creaza_station_equation_datatable_structure();

                        start1 = Convert.ToDouble(Start_sta_cl.Replace("+", ""));
                        if (start1 != 0)
                        {
                            _AGEN_mainform.tpage_st_eq.set_textBox_start_station_CL(Start_sta_cl);
                            _AGEN_mainform.dt_station_equation.Rows.Add();
                            _AGEN_mainform.dt_station_equation.Rows[_AGEN_mainform.dt_station_equation.Rows.Count - 1]["Station Back"] = 0;
                            _AGEN_mainform.dt_station_equation.Rows[_AGEN_mainform.dt_station_equation.Rows.Count - 1]["Station Ahead"] = start1;
                            _AGEN_mainform.dt_station_equation.Rows[_AGEN_mainform.dt_station_equation.Rows.Count - 1]["Reroute Start X"] = pt0.X;
                            _AGEN_mainform.dt_station_equation.Rows[_AGEN_mainform.dt_station_equation.Rows.Count - 1]["Reroute Start Y"] = pt0.Y;
                            _AGEN_mainform.dt_station_equation.Rows[_AGEN_mainform.dt_station_equation.Rows.Count - 1]["Reroute Start Z"] = pt0.Z;
                            _AGEN_mainform.dt_station_equation.Rows[_AGEN_mainform.dt_station_equation.Rows.Count - 1]["Reroute End X"] = pt0.X;
                            _AGEN_mainform.dt_station_equation.Rows[_AGEN_mainform.dt_station_equation.Rows.Count - 1]["Reroute End Y"] = pt0.Y;
                            _AGEN_mainform.dt_station_equation.Rows[_AGEN_mainform.dt_station_equation.Rows.Count - 1]["Reroute End Z"] = pt0.Z;
                        }
                    }

                    Functions.create_backup(fisier_cl);

                    if (start1 == 0)
                    {

                        Populate_centerline_file(fisier_cl, true, false);
                    }
                    else
                    {
                        Populate_centerline_file(fisier_cl, false, false);


                    }
                }
                else
                {
                    MessageBox.Show("no project folder!/r/noperation aborted");
                }
            }

            catch (System.Exception ex)
            {
                Set_centerline_label_to_red();
                _AGEN_mainform.tpage_processing.Hide();
                set_enable_true();
                MessageBox.Show(ex.Message);
            }

            if (System.IO.File.Exists(_AGEN_mainform.config_path) == true)
            {
                _AGEN_mainform.tpage_setup.button_align_config_saveall_boolean(true);
            }


            _AGEN_mainform.tpage_processing.Hide();
            _AGEN_mainform.tpage_st_eq.Populate_datagridview_with_equation_data();


            set_enable_true();
        }

        public void Populate_centerline_file(string File1, bool delete_steq, bool CSF)
        {
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

                if (Excel1.Workbooks.Count == 0) Excel1.Visible = _AGEN_mainform.ExcelVisible;




                Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;

                if (System.IO.File.Exists(File1) == false)
                {
                    Workbook1 = Excel1.Workbooks.Add();
                }

                else
                {
                    Workbook1 = Excel1.Workbooks.Open(File1);
                }
                Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];

                try
                {
                    string segment1 = _AGEN_mainform.tpage_setup.Get_segment_name1();
                    if (segment1 == "not defined") segment1 = "";
                    Functions.Transfer_to_worksheet_Data_table(W1, _AGEN_mainform.dt_centerline, _AGEN_mainform.Start_row_CL, "General");
                    Functions.Create_header_centerline_file(W1, _AGEN_mainform.tpage_setup.Get_client_name(), _AGEN_mainform.tpage_setup.Get_project_name(), segment1, CSF, _AGEN_mainform.version);


                    if (delete_steq == true)
                    {
                        _AGEN_mainform.dt_station_equation = Functions.Creaza_station_equation_datatable_structure();

                        if (Workbook1.Worksheets.Count > 1)
                        {
                            Workbook1.Worksheets[2].Columns["A:XX"].Delete();
                        }

                    }

                    else
                    {
                        if (_AGEN_mainform.dt_station_equation != null)
                        {
                            if (_AGEN_mainform.dt_station_equation.Rows.Count > 0)
                            {
                                if (Workbook1.Worksheets.Count == 1)
                                {
                                    Microsoft.Office.Interop.Excel.Worksheet W3 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[1], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(W3);
                                }

                                Microsoft.Office.Interop.Excel.Worksheet W2 = Workbook1.Worksheets[2];
                                W2.Name = "St_eq";

                                try
                                {
                                    Functions.Transfer_to_worksheet_Data_table(W2, _AGEN_mainform.dt_station_equation, _AGEN_mainform.Start_row_station_equation, "General");
                                    Functions.Create_header_station_eq(W2, _AGEN_mainform.tpage_setup.Get_client_name(), _AGEN_mainform.tpage_setup.Get_project_name(), segment1);

                                }
                                catch (System.Exception ex)
                                {
                                    System.Windows.Forms.MessageBox.Show(ex.Message);

                                }
                                finally
                                {
                                    if (W2 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W2);

                                }
                            }
                        }

                    }

                    if (System.IO.File.Exists(File1) == false)
                    {
                        Workbook1.SaveAs(File1);
                    }

                    else
                    {
                        Workbook1.Save();
                    }


                    Workbook1.Close();
                    if (Excel1.Workbooks.Count == 0)
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
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }

        }



        public void Add_to_centerline_file_station_equations(String File1, System.Data.DataTable Data_table_station_equation)
        {
            try
            {
                if (System.IO.File.Exists(File1) == false)
                {
                    MessageBox.Show("the centerline file does not exists\r\n" + File1);
                    return;
                }


                Microsoft.Office.Interop.Excel.Application Excel1 = null;

                try
                {
                    Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                }
                catch (System.Exception ex)
                {
                    Excel1 = new Microsoft.Office.Interop.Excel.Application();

                }

                if (Excel1.Workbooks.Count == 0) Excel1.Visible = _AGEN_mainform.ExcelVisible;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(File1);

                if (Workbook1.Worksheets.Count == 1)
                {
                    Microsoft.Office.Interop.Excel.Worksheet W3 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[1], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(W3);
                }

                Microsoft.Office.Interop.Excel.Worksheet W2 = Workbook1.Worksheets[2];
                W2.Name = "St_eq";

                try
                {
                    string segment1 = _AGEN_mainform.tpage_setup.Get_segment_name1();
                    if (segment1 == "not defined") segment1 = "";
                    Functions.Transfer_to_worksheet_Data_table(W2, Data_table_station_equation, _AGEN_mainform.Start_row_station_equation, "General");
                    Functions.Create_header_station_eq(W2, _AGEN_mainform.tpage_setup.Get_client_name(), _AGEN_mainform.tpage_setup.Get_project_name(), segment1);
                    Workbook1.Save();
                    Workbook1.Close();
                    if (Excel1.Workbooks.Count == 0)
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
                    if (W2 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W2);
                    if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                    if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }

        }


        public Polyline create_new_centerline(List<ObjectId> lista1)
        {
            string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();

            if (System.IO.Directory.Exists(ProjFolder) == false)
            {
                MessageBox.Show("No project Loaded");
                return null;
            }

            if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
            {
                ProjFolder = ProjFolder + "\\";
            }

            string fisier_cl = ProjFolder + _AGEN_mainform.cl_excel_name;

            if (System.IO.File.Exists(fisier_cl) == false)
            {
                MessageBox.Show("No centerline file found");
                return null;
            }

            if (_AGEN_mainform.dt_centerline == null || _AGEN_mainform.dt_centerline.Rows.Count == 0)
            {
                MessageBox.Show("No centerline data");
                return null;
            }



            if (lista1 == null)
            {
                MessageBox.Show("No reroute data");
                return null;
            }

            if (lista1.Count == 0)
            {
                MessageBox.Show("No reroute data");
                return null;
            }

            ObjectId[] Empty_array = null;

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Editor1.SetImpliedSelection(Empty_array);

            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                Polyline3d Poly3D = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);
                Polyline Poly2D = Functions.Build_2d_poly_for_scanning(_AGEN_mainform.dt_centerline);

                BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);


                Functions.Creaza_layer(_AGEN_mainform.layer_centerline, _AGEN_mainform.color_index_cl, true);
                Functions.set_layer_lineweight(_AGEN_mainform.layer_centerline, _AGEN_mainform.lw_cl);
                Functions.Creaza_layer("ORIGINAL_ROUTE", 4, false);

                List<ObjectId> lista_del = new List<ObjectId>();
                lista_del.Add(Poly3D.ObjectId);

                for (int n = 0; n < lista1.Count; ++n)
                {
                    Entity curba1 = Trans1.GetObject(lista1[n], OpenMode.ForRead) as Entity;
                    if (curba1 != null)
                    {
                        Polyline polyR = null;
                        Polyline3d polyR3d = null;

                        if (curba1 is Polyline3d)
                        {
                            polyR = Functions.Build_2dpoly_from_3d(curba1 as Polyline3d) as Polyline;
                            polyR3d = curba1 as Polyline3d;
                        }

                        if (curba1 is Polyline)
                        {
                            polyR = curba1 as Polyline;
                            polyR3d = Functions.Build_3d_poly_from2D_poly(polyR);
                            lista_del.Add(polyR3d.ObjectId);
                        }

                        if (polyR != null)
                        {
                            Polyline3d new3D = new Polyline3d();
                            new3D.Layer = _AGEN_mainform.layer_centerline;
                            new3D.ColorIndex = 256;
                            BTrecord.AppendEntity(new3D);
                            Trans1.AddNewlyCreatedDBObject(new3D, true);

                            Polyline new2D = new Polyline();

                            Polyline original2D = new Polyline();


                            Polyline3d original3D = new Polyline3d();

                            original3D.Layer = "ORIGINAL_ROUTE";
                            original3D.ColorIndex = 256;
                            BTrecord.AppendEntity(original3D);
                            Trans1.AddNewlyCreatedDBObject(original3D, true);


                            Point3d pt_on_poly1 = Poly2D.GetClosestPointTo(polyR.StartPoint, Vector3d.ZAxis, false);

                            double dist1 = pt_on_poly1.DistanceTo(polyR.StartPoint);

                            Point3d pt_on_poly2 = Poly2D.GetClosestPointTo(polyR.EndPoint, Vector3d.ZAxis, false);

                            double dist2 = pt_on_poly2.DistanceTo(polyR.EndPoint);

                            if (dist1 < 0.01 && dist2 < 0.01)
                            {
                                double param1 = Poly2D.GetParameterAtPoint(pt_on_poly1);
                                double param2 = Poly2D.GetParameterAtPoint(pt_on_poly2);

                                bool reversed = false;

                                if (param1 > param2)
                                {
                                    double temp = param1;
                                    param1 = param2;
                                    param2 = temp;

                                    Point3d ptt = pt_on_poly1;
                                    pt_on_poly1 = pt_on_poly2;
                                    pt_on_poly2 = ptt;
                                    reversed = true;
                                }

                                int q = 0;

                                if (Math.Round(param1, 4) != Math.Ceiling(param1))
                                {
                                    original2D.AddVertexAt(q, new Point2d(Poly2D.GetPointAtParameter(param1).X, Poly2D.GetPointAtParameter(param1).Y), 0,
                                                               Poly2D.GetStartWidthAt(Convert.ToInt32(Math.Floor(param1))),
                                                               Poly2D.GetEndWidthAt(Convert.ToInt32(Math.Floor(param1))));
                                    q = q + 1;

                                    PolylineVertex3d Vertex_new = new PolylineVertex3d(Poly3D.GetPointAtParameter(param1));
                                    original3D.AppendVertex(Vertex_new);
                                    Trans1.AddNewlyCreatedDBObject(Vertex_new, true);

                                }

                                for (int i = Convert.ToInt32(Math.Ceiling(param1)); i <= Convert.ToInt32(Math.Floor(param2)); ++i)
                                {
                                    original2D.AddVertexAt(q, Poly2D.GetPoint2dAt(i), Poly2D.GetBulgeAt(i), Poly2D.GetStartWidthAt(i), Poly2D.GetEndWidthAt(i));
                                    q = q + 1;

                                    PolylineVertex3d Vertex_new = new PolylineVertex3d(Poly3D.GetPointAtParameter(i));
                                    original3D.AppendVertex(Vertex_new);
                                    Trans1.AddNewlyCreatedDBObject(Vertex_new, true);
                                }

                                if (Math.Round(param2, 4) != Math.Floor(param2))
                                {
                                    original2D.AddVertexAt(q, new Point2d(Poly2D.GetPointAtParameter(param2).X, Poly2D.GetPointAtParameter(param2).Y), 0,
                                                               Poly2D.GetStartWidthAt(Convert.ToInt32(Math.Floor(param2))),
                                                               Poly2D.GetEndWidthAt(Convert.ToInt32(Math.Floor(param2))));
                                    q = q + 1;

                                    PolylineVertex3d Vertex_new = new PolylineVertex3d(Poly3D.GetPointAtParameter(param2));
                                    original3D.AppendVertex(Vertex_new);
                                    Trans1.AddNewlyCreatedDBObject(Vertex_new, true);

                                }

                                if (_AGEN_mainform.Project_type == "2D")
                                {
                                    original2D.Layer = "ORIGINAL_ROUTE";
                                    original2D.ColorIndex = 256;
                                    BTrecord.AppendEntity(original2D);
                                    Trans1.AddNewlyCreatedDBObject(original2D, true);
                                    lista_del.Add(original3D.ObjectId);
                                }

                                int k = 0;

                                for (int i = 0; i <= param1; ++i)
                                {
                                    new2D.AddVertexAt(k, Poly2D.GetPoint2dAt(i), Poly2D.GetBulgeAt(i), Poly2D.GetStartWidthAt(i), Poly2D.GetEndWidthAt(i));
                                    k = k + 1;
                                    PolylineVertex3d Vertex_new = new PolylineVertex3d(Poly3D.GetPointAtParameter(i));
                                    new3D.AppendVertex(Vertex_new);
                                    Trans1.AddNewlyCreatedDBObject(Vertex_new, true);

                                }

                                if (Math.Round(param1, 4) > Math.Floor(param1))
                                {
                                    new2D.AddVertexAt(k, new Point2d(Poly2D.GetPointAtParameter(param1).X, Poly2D.GetPointAtParameter(param1).Y), 0,
                                                               Poly2D.GetStartWidthAt(Convert.ToInt32(Math.Floor(param1))),
                                                               Poly2D.GetEndWidthAt(Convert.ToInt32(Math.Floor(param1))));
                                    k = k + 1;

                                    PolylineVertex3d Vertex_new = new PolylineVertex3d(Poly3D.GetPointAtParameter(param1));
                                    new3D.AppendVertex(Vertex_new);
                                    Trans1.AddNewlyCreatedDBObject(Vertex_new, true);

                                }

                                if (reversed == false)
                                {
                                    for (int i = 1; i < polyR.EndParam; ++i)
                                    {
                                        new2D.AddVertexAt(k, polyR.GetPoint2dAt(i), polyR.GetBulgeAt(i), polyR.GetStartWidthAt(i), polyR.GetEndWidthAt(i));
                                        k = k + 1;

                                        Point3d pt_r = polyR3d.GetPointAtParameter(i);

                                        if (curba1 is Polyline)
                                        {
                                            pt_r = new Point3d(pt_r.X, pt_r.Y, Poly3D.GetPointAtParameter(param1).Z);
                                        }
                                        PolylineVertex3d Vertex_new = new PolylineVertex3d(pt_r);
                                        new3D.AppendVertex(Vertex_new);
                                        Trans1.AddNewlyCreatedDBObject(Vertex_new, true);

                                    }
                                }
                                else
                                {
                                    for (int i = Convert.ToInt32(polyR.EndParam) - 1; i > 0; --i)
                                    {
                                        new2D.AddVertexAt(k, polyR.GetPoint2dAt(i), polyR.GetBulgeAt(i), polyR.GetStartWidthAt(i), polyR.GetEndWidthAt(i));
                                        k = k + 1;

                                        Point3d pt_r = polyR3d.GetPointAtParameter(i);

                                        if (curba1 is Polyline)
                                        {
                                            pt_r = new Point3d(pt_r.X, pt_r.Y, Poly3D.GetPointAtParameter(param1).Z);
                                        }
                                        PolylineVertex3d Vertex_new = new PolylineVertex3d(pt_r);
                                        new3D.AppendVertex(Vertex_new);
                                        Trans1.AddNewlyCreatedDBObject(Vertex_new, true);

                                    }

                                }


                                if (Math.Round(param2, 4) < Math.Ceiling(param2))
                                {
                                    new2D.AddVertexAt(k, new Point2d(Poly2D.GetPointAtParameter(param2).X, Poly2D.GetPointAtParameter(param2).Y), 0,
                                                                Poly2D.GetStartWidthAt(Convert.ToInt32(Math.Floor(param2))),
                                                                Poly2D.GetEndWidthAt(Convert.ToInt32(Math.Floor(param2))));

                                    k = k + 1;

                                    PolylineVertex3d Vertex_new = new PolylineVertex3d(Poly3D.GetPointAtParameter(param2));
                                    new3D.AppendVertex(Vertex_new);
                                    Trans1.AddNewlyCreatedDBObject(Vertex_new, true);

                                }

                                for (int i = Convert.ToInt32(Math.Ceiling(param2)); i <= Poly2D.EndParam; ++i)
                                {
                                    new2D.AddVertexAt(k, Poly2D.GetPoint2dAt(i), Poly2D.GetBulgeAt(i), Poly2D.GetStartWidthAt(i), Poly2D.GetEndWidthAt(i));
                                    k = k + 1;
                                    PolylineVertex3d Vertex_new = new PolylineVertex3d(Poly3D.GetPointAtParameter(i));
                                    new3D.AppendVertex(Vertex_new);
                                    Trans1.AddNewlyCreatedDBObject(Vertex_new, true);

                                }

                                Poly2D = new2D;
                                Poly3D = new3D;
                                lista_del.Add(new3D.ObjectId);
                            }
                        }
                    }
                }

                for (int n = 0; n < lista_del.Count - 1; ++n)
                {

                    Entity ent1 = Trans1.GetObject(lista_del[n], OpenMode.ForWrite) as Entity;
                    ent1.Erase();

                }

                _AGEN_mainform.Poly2D = Poly2D;
                _AGEN_mainform.Poly3D = Poly3D;

                if (_AGEN_mainform.Project_type == "3D")
                {


                }
                else
                {

                    Poly2D.Layer = _AGEN_mainform.layer_centerline;
                    Poly2D.ColorIndex = 256;
                    BTrecord.AppendEntity(Poly2D);
                    Trans1.AddNewlyCreatedDBObject(Poly2D, true);
                    Entity ent1 = Trans1.GetObject(lista_del[lista_del.Count - 1], OpenMode.ForWrite) as Entity;
                    ent1.Erase();
                }

                _AGEN_mainform.dt_centerline = Functions.Creaza_centerline_datatable_structure();
                string Col_EqSta = "EqSta";

                for (int i = 0; i < Poly2D.NumberOfVertices; ++i)
                {

                    double x2 = Poly2D.GetPointAtParameter(i).X;
                    double y2 = Poly2D.GetPointAtParameter(i).Y;
                    double z2 = Poly3D.GetPointAtParameter(i).Z;

                    double Dist2d = Poly2D.GetDistanceAtParameter(i);
                    double Dist3d = Poly3D.GetDistanceAtParameter(i);

                    _AGEN_mainform.dt_centerline.Rows.Add();
                    _AGEN_mainform.dt_centerline.Rows[_AGEN_mainform.dt_centerline.Rows.Count - 1][_AGEN_mainform.Col_x] = x2;
                    _AGEN_mainform.dt_centerline.Rows[_AGEN_mainform.dt_centerline.Rows.Count - 1][_AGEN_mainform.Col_y] = y2;
                    _AGEN_mainform.dt_centerline.Rows[_AGEN_mainform.dt_centerline.Rows.Count - 1][_AGEN_mainform.Col_z] = z2;
                    _AGEN_mainform.dt_centerline.Rows[_AGEN_mainform.dt_centerline.Rows.Count - 1][_AGEN_mainform.Col_2DSta] = Dist2d;
                    _AGEN_mainform.dt_centerline.Rows[_AGEN_mainform.dt_centerline.Rows.Count - 1][Col_EqSta] = DBNull.Value;

                    if (_AGEN_mainform.Project_type == "3D")
                    {
                        _AGEN_mainform.dt_centerline.Rows[_AGEN_mainform.dt_centerline.Rows.Count - 1][_AGEN_mainform.Col_3DSta] = Dist3d;

                    }

                    if (_AGEN_mainform.dt_station_equation != null)
                    {
                        if (_AGEN_mainform.dt_station_equation.Rows.Count > 0)
                        {


                            if (_AGEN_mainform.Project_type == "3D")
                            {
                                _AGEN_mainform.dt_centerline.Rows[_AGEN_mainform.dt_centerline.Rows.Count - 1][Col_EqSta] =
                                    Math.Round(Functions.Station_equation_of(Dist3d, _AGEN_mainform.dt_station_equation), _AGEN_mainform.round1);
                            }
                            else
                            {
                                _AGEN_mainform.dt_centerline.Rows[_AGEN_mainform.dt_centerline.Rows.Count - 1][Col_EqSta] =
                                     Math.Round(Functions.Station_equation_of(Dist2d, _AGEN_mainform.dt_station_equation), _AGEN_mainform.round1);
                            }
                        }
                    }



                    if (i > 0 && i < Poly2D.NumberOfVertices - 1)
                    {
                        double x1 = Poly2D.GetPointAtParameter(i - 1).X;
                        double y1 = Poly2D.GetPointAtParameter(i - 1).Y;
                        double x3 = Poly2D.GetPointAtParameter(i + 1).X;
                        double y3 = Poly2D.GetPointAtParameter(i + 1).Y;

                        string Deflexia = Functions.Get_deflection_angle_dms(x1, y1, x2, y2, x3, y3);
                        double Defl1 = 180 * Functions.Get_deflection_angle_rad(x1, y1, x2, y2, x3, y3) / Math.PI;

                        _AGEN_mainform.dt_centerline.Rows[_AGEN_mainform.dt_centerline.Rows.Count - 1][_AGEN_mainform.Col_DeflAng] = Defl1;
                        _AGEN_mainform.dt_centerline.Rows[_AGEN_mainform.dt_centerline.Rows.Count - 1][_AGEN_mainform.Col_DeflAngDMS] = Deflexia;
                    }

                    if (i < Poly2D.NumberOfVertices - 1)
                    {
                        double x3 = Poly2D.GetPointAtParameter(i + 1).X;
                        double y3 = Poly2D.GetPointAtParameter(i + 1).Y;
                        _AGEN_mainform.dt_centerline.Rows[_AGEN_mainform.dt_centerline.Rows.Count - 1][_AGEN_mainform.Col_Bearing] = Functions.Get_Quadrant_bearing(Functions.GET_Bearing_rad(x2, y2, x3, y3));
                        _AGEN_mainform.dt_centerline.Rows[_AGEN_mainform.dt_centerline.Rows.Count - 1][_AGEN_mainform.Col_Distance] = Math.Pow((x2 - x3) * (x2 - x3) + (y2 - y3) * (y2 - y3), 0.5);
                    }

                }


                Trans1.Commit();
                return Poly2D;
            }
        }

        public static void Create_csf_cl_od_table()
        {
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        List<string> List1 = new List<string>();
                        List<string> List2 = new List<string>();
                        List<Autodesk.Gis.Map.Constants.DataType> List3 = new List<Autodesk.Gis.Map.Constants.DataType>();

                        List1.Add("Version");
                        List2.Add("Excel file");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add("Segment");
                        List2.Add("Name of the sgment");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add("Note1");
                        List2.Add("Notes");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        Functions.Get_object_data_table("Agen_csf_cl", "Generated by AGEN", List1, List2, List3);

                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        private void button_refresh_text_styles_Click(object sender, EventArgs e)
        {
            try
            {
                set_enable_false();
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Functions.Incarca_existing_textstyles_to_combobox(comboBox_text_styles);

                        if (comboBox_text_styles.Items.Contains("Agen_Stationing") == true)
                        {
                            comboBox_text_styles.SelectedIndex = comboBox_text_styles.Items.IndexOf("Agen_Stationing");
                        }
                        else
                        {
                            comboBox_text_styles.SelectedIndex = 0;
                        }


                        Trans1.Dispose();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            set_enable_true();

        }


        private void sort_canadian_station_eq(Polyline Poly2D)
        {
            if (_AGEN_mainform.dt_station_equation != null && _AGEN_mainform.dt_station_equation.Rows.Count > 0)
            {

                if (_AGEN_mainform.dt_station_equation.Columns.Contains("measured") == false) _AGEN_mainform.dt_station_equation.Columns.Add("measured", typeof(double));

                for (int k = 0; k < _AGEN_mainform.dt_station_equation.Rows.Count; ++k)
                {

                    if (_AGEN_mainform.dt_station_equation.Rows[k][rr_end_x] != DBNull.Value && _AGEN_mainform.dt_station_equation.Rows[k][rr_end_y] != DBNull.Value)
                    {

                        double x = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[k][rr_end_x]);
                        double y = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[k][rr_end_y]);

                        Point3d pt_eq = new Point3d(x, y, 0);
                        Point3d pt_on_poly = Poly2D.GetClosestPointTo(pt_eq, Vector3d.ZAxis, false);


                        double dist1 = Math.Pow(Math.Pow(pt_eq.X - pt_on_poly.X, 2) + Math.Pow(pt_eq.Y - pt_on_poly.Y, 2), 0.5);
                        if (dist1 < 0.001)
                        {
                            double meas1 = Poly2D.GetDistAtPoint(pt_on_poly);
                            _AGEN_mainform.dt_station_equation.Rows[k]["measured"] = meas1;
                        }
                        else
                        {
                            MessageBox.Show("Equatin point " + pt_eq.ToString() + " is " + Math.Round(dist1, 3).ToString() + " apart from the centerline");
                            set_enable_true();
                            return;
                        }
                    }



                }

                _AGEN_mainform.dt_station_equation = Functions.Sort_data_table(_AGEN_mainform.dt_station_equation, "measured");
            }
        }

        private void button_draw_stationing_Click(object sender, EventArgs e)
        {
            Ag = this.MdiParent as _AGEN_mainform;

            string od_name = "Agen_stationing";
            Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
            #region object data stationing
            List<object> Lista_val_CL = new List<object>();
            List<Autodesk.Gis.Map.Constants.DataType> Lista_type_CL = new List<Autodesk.Gis.Map.Constants.DataType>();

            Lista_val_CL.Add(comboBox_segment_name.Text);
            Lista_type_CL.Add(Autodesk.Gis.Map.Constants.DataType.Character);

            Lista_val_CL.Add(System.DateTime.Today.Year + "-" + System.DateTime.Today.Month + "-" + System.DateTime.Today.Day + " at " + System.DateTime.Now.Hour + ":" + System.DateTime.Now.Minute + " by " + Environment.UserName.ToUpper());
            Lista_type_CL.Add(Autodesk.Gis.Map.Constants.DataType.Character);


            #endregion


            if (checkBox_create_major_labels.Checked == false &&
                checkBox_create_minor_labels.Checked == false &&
                checkBox_draw_major_lines.Checked == false &&
                checkBox_draw_minor_lines.Checked == false)
            {
                MessageBox.Show("Please select at least an option!\r\nAll checkboxes are unchecked!");
                return;
            }

            if (Functions.IsNumeric(textBox_spacing_major.Text) == false && checkBox_draw_major_lines.Checked == true)
            {
                MessageBox.Show("spacing major issue");
                return;
            }


            double spacing_major = Math.Abs(Convert.ToDouble(textBox_spacing_major.Text));

            if (spacing_major == 0 && checkBox_draw_major_lines.Checked == true)
            {
                MessageBox.Show("spacing major issue");
                return;
            }

            if (Functions.IsNumeric(textBox_spacing_minor.Text) == false && checkBox_draw_minor_lines.Checked == true)
            {
                MessageBox.Show("spacing minor issue");
                return;
            }


            double spacing_minor = Math.Abs(Convert.ToDouble(textBox_spacing_minor.Text));

            if (spacing_minor == 0 && checkBox_draw_minor_lines.Checked == true)
            {
                MessageBox.Show("spacing minor issue");
                return;
            }



            if (Functions.IsNumeric(textBox_tic_major.Text) == false && checkBox_create_major_labels.Checked == true)
            {
                MessageBox.Show("tick major issue");
                return;
            }


            double tick_major = Math.Abs(Convert.ToDouble(textBox_tic_major.Text));

            if (tick_major == 0 && checkBox_create_major_labels.Checked == true)
            {
                MessageBox.Show("tick major issue");
                return;
            }

            if (Functions.IsNumeric(textBox_tic_minor.Text) == false && checkBox_create_minor_labels.Checked == true)
            {
                MessageBox.Show("tick minor issue");
                return;
            }


            double tick_minor = Math.Abs(Convert.ToDouble(textBox_tic_minor.Text));

            if (tick_minor == 0 && checkBox_create_minor_labels.Checked == true)
            {
                MessageBox.Show("tick minor issue");
                return;
            }

         
            double texth = Functions.Get_text_height_from_textstyle(comboBox_text_styles.Text);
            double gap1 = texth;


            if (texth == 0)
            {
                texth = 8;
                gap1 = texth;
            }



            Functions.Kill_excel();



            if (Functions.Get_if_workbook_is_open_in_Excel("centerline.xlsx") == true)
            {
                MessageBox.Show("Please close the centerline file");
                return;
            }

            string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();

            if (System.IO.Directory.Exists(ProjFolder) == false)
            {
                MessageBox.Show("No project Loaded");
                return;
            }

            if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
            {
                ProjFolder = ProjFolder + "\\";
            }

            string fisier_cl = ProjFolder + _AGEN_mainform.cl_excel_name;


            if (System.IO.File.Exists(fisier_cl) == false)
            {
                MessageBox.Show("No centerline file found");
                return;
            }

            _AGEN_mainform.tpage_processing.Show();

            set_enable_false();

            Load_centerline_and_station_equation(fisier_cl);

            if (_AGEN_mainform.dt_centerline != null)
            {
                if (_AGEN_mainform.dt_centerline.Rows.Count > 0)
                {

                    _AGEN_mainform.layer_stationing = _AGEN_mainform.layer_stationing_original + "_" + _AGEN_mainform.current_segment;

                    try
                    {

                        Functions.Create_stationing_od_table();

                        delete_entities_with_OD(_AGEN_mainform.layer_stationing, od_name);
                        delete_entities_with_OD(_AGEN_mainform.layer_centerline, od_name);

                        Functions.Creaza_layer(_AGEN_mainform.layer_stationing, 2, true);
                        Functions.Creaza_layer(_AGEN_mainform.layer_centerline, _AGEN_mainform.color_index_cl, true);
                        Functions.set_layer_lineweight(_AGEN_mainform.layer_centerline, _AGEN_mainform.lw_cl);

                        Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                        Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                        using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                        {

                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                            {
                                Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                                Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);


                                Polyline3d Poly3D = null;
                                bool is_equated = false;

                                if (_AGEN_mainform.Project_type == "3D")
                                {
                                    Poly3D = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);
                                    Poly3D.Layer = _AGEN_mainform.layer_centerline;
                                    Poly3D.ColorIndex = 256;

                                }
                                else
                                {
                                    Poly3D = null;
                                }



                                Polyline Poly2D = Functions.Build_2d_poly_for_scanning(_AGEN_mainform.dt_centerline);
                                Poly2D.Layer = _AGEN_mainform.layer_centerline;
                                Poly2D.ColorIndex = 256;




                                #region USA
                                if (_AGEN_mainform.COUNTRY == "USA")
                                {
                                    double start1 = 0;




                                    if (_AGEN_mainform.dt_station_equation != null)
                                    {
                                        if (_AGEN_mainform.dt_station_equation.Rows.Count > 0)
                                        {
                                            if (_AGEN_mainform.dt_station_equation.Rows[0]["Station Back"] != DBNull.Value && _AGEN_mainform.dt_station_equation.Rows[0]["Station Ahead"] != DBNull.Value)
                                            {
                                                double SB = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[0]["Station Back"]);
                                                double SA = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[0]["Station Ahead"]);
                                                if (SB == 0)
                                                {
                                                    start1 = SA;
                                                }
                                            }


                                            double param1p = 0;
                                            double SAp = 0;

                                            for (int i = 0; i < _AGEN_mainform.dt_station_equation.Rows.Count; ++i)
                                            {
                                                if (_AGEN_mainform.dt_station_equation.Rows[i]["Station Back"] != DBNull.Value &&
                                                    _AGEN_mainform.dt_station_equation.Rows[i]["Station Ahead"] != DBNull.Value &&
                                                    _AGEN_mainform.dt_station_equation.Rows[i]["Reroute End X"] != DBNull.Value &&
                                                    _AGEN_mainform.dt_station_equation.Rows[i]["Reroute End Y"] != DBNull.Value &&
                                                    _AGEN_mainform.dt_station_equation.Rows[i]["Reroute End Z"] != DBNull.Value)
                                                {
                                                    double SB1 = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i]["Station Back"]);
                                                    double SA = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i]["Station Ahead"]);

                                                    if (i == 0)
                                                    {
                                                        double start_eq = 0;
                                                        if (SB1 == 0) start_eq = SA;
                                                        add_zero_plus_zero_zero_stationing(Poly2D, od_name, start_eq, gap1, texth, tick_major, 0);
                                                    }

                                                    Point3d pt_end = new Point3d(Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i]["Reroute End X"]),
                                                        Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i]["Reroute End Y"]),
                                                        Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i]["Reroute End Z"]));
                                                    double param1 = Poly2D.GetParameterAtPoint(Poly2D.GetClosestPointTo(new Point3d(pt_end.X, pt_end.Y, Poly2D.Elevation), Vector3d.ZAxis, false));

                                                    if (_AGEN_mainform.Project_type == "3D")
                                                    {
                                                        if (param1 > Poly3D.EndParam) param1 = Poly3D.EndParam;
                                                    }
                                                    else
                                                    {
                                                        if (param1 > Poly2D.EndParam) param1 = Poly2D.EndParam;
                                                    }

                                                    if (i == _AGEN_mainform.dt_station_equation.Rows.Count - 1)
                                                    {
                                                        double dist1 = -1;
                                                        double dist2 = -1;

                                                        if (_AGEN_mainform.Project_type == "3D")
                                                        {
                                                            dist1 = Poly3D.GetDistanceAtParameter(param1);
                                                            dist2 = Poly3D.Length;
                                                        }

                                                        else
                                                        {
                                                            dist1 = Poly2D.GetDistanceAtParameter(param1);
                                                            dist2 = Poly2D.Length;
                                                        }

                                                        double end_eq = SA + dist2 - dist1;
                                                        add_last_stationing(Poly2D, od_name, end_eq, gap1, texth, tick_major, 0);
                                                    }

                                                    if (_AGEN_mainform.Project_type == "3D")
                                                    {
                                                        _AGEN_mainform.dt_station_equation.Rows[i]["Station Back"] = SAp + Poly3D.GetDistanceAtParameter(param1) - Poly3D.GetDistanceAtParameter(param1p);
                                                    }
                                                    else
                                                    {
                                                        _AGEN_mainform.dt_station_equation.Rows[i]["Station Back"] = SAp + Poly2D.GetDistanceAtParameter(param1) - Poly2D.GetDistanceAtParameter(param1p);
                                                    }


                                                    if (SB1 != 0)
                                                    {
                                                        is_equated = true;

                                                        Polyline Poly2D_eq = new Polyline();
                                                        Polyline3d Poly3D_eq = new Polyline3d();
                                                        if (_AGEN_mainform.Project_type == "3D")
                                                        {
                                                            Poly3D_eq = new Polyline3d();
                                                            BTrecord.AppendEntity(Poly3D_eq);
                                                            Trans1.AddNewlyCreatedDBObject(Poly3D_eq, true);

                                                            if (Math.Ceiling(param1p) != param1p)
                                                            {
                                                                PolylineVertex3d Vertex_last = new PolylineVertex3d(new Point3d(Poly3D.GetPointAtParameter(param1p).X, Poly3D.GetPointAtParameter(param1p).Y, Poly3D.GetPointAtParameter(param1p).Z));
                                                                Poly3D_eq.AppendVertex(Vertex_last);
                                                                Trans1.AddNewlyCreatedDBObject(Vertex_last, true);
                                                            }

                                                            for (int k = Convert.ToInt32(Math.Ceiling(param1p)); k <= Math.Floor(param1); ++k)
                                                            {
                                                                PolylineVertex3d Vertex_eq = new PolylineVertex3d(new Point3d(Poly3D.GetPointAtParameter(k).X, Poly3D.GetPointAtParameter(k).Y, Poly3D.GetPointAtParameter(k).Z));
                                                                Poly3D_eq.AppendVertex(Vertex_eq);
                                                                Trans1.AddNewlyCreatedDBObject(Vertex_eq, true);
                                                            }
                                                            if (Math.Floor(param1) != param1)
                                                            {
                                                                PolylineVertex3d Vertex_last = new PolylineVertex3d(new Point3d(Poly3D.GetPointAtParameter(param1).X, Poly3D.GetPointAtParameter(param1).Y, Poly3D.GetPointAtParameter(param1).Z));
                                                                Poly3D_eq.AppendVertex(Vertex_last);
                                                                Trans1.AddNewlyCreatedDBObject(Vertex_last, true);
                                                            }
                                                        }
                                                        else
                                                        {
                                                            Poly2D_eq = new Polyline();
                                                            Poly3D_eq = null;
                                                            Point2d p2d = new Point2d(Poly2D.GetPointAtParameter(param1p).X, Poly2D.GetPointAtParameter(param1p).Y);
                                                            int r = 0;
                                                            double b_start = 0;



                                                            Poly2D_eq.AddVertexAt(r, p2d, b_start, 0, 0);
                                                            ++r;
                                                            for (int k = Convert.ToInt32(Math.Ceiling(param1p)); k <= Math.Floor(param1); ++k)
                                                            {
                                                                p2d = new Point2d(Poly2D.GetPointAtParameter(k).X, Poly2D.GetPointAtParameter(k).Y);

                                                                Poly2D_eq.AddVertexAt(r, p2d, Poly2D.GetBulgeAt(k), 0, 0);
                                                                ++r;
                                                            }

                                                            if (Math.Floor(param1) != param1)
                                                            {
                                                                p2d = new Point2d(Poly2D.GetPointAtParameter(param1).X, Poly2D.GetPointAtParameter(param1).Y);
                                                                double b_end = 0;
                                                                Poly2D_eq.AddVertexAt(r, p2d, b_end, 0, 0);
                                                                ++r;
                                                            }

                                                        }




                                                        create_stationing(Poly3D_eq, Poly2D_eq, od_name, SAp, gap1, texth, spacing_major, spacing_minor, tick_major, tick_minor);
                                                        param1p = param1;
                                                    }
                                                    SAp = SA;
                                                }
                                            }

                                            double last_param = -1;

                                            if (_AGEN_mainform.Project_type == "3D")
                                            {
                                                last_param = Poly3D.EndParam;
                                            }
                                            else
                                            {
                                                last_param = Poly2D.EndParam;
                                            }

                                            if (param1p < last_param)
                                            {

                                                Polyline Poly2D_eq = new Polyline();
                                                Polyline3d Poly3D_eq = new Polyline3d();

                                                if (_AGEN_mainform.Project_type == "3D")
                                                {
                                                    BTrecord.AppendEntity(Poly3D_eq);
                                                    Trans1.AddNewlyCreatedDBObject(Poly3D_eq, true);
                                                    if (Math.Ceiling(param1p) != param1p)
                                                    {
                                                        PolylineVertex3d Vertex_last = new PolylineVertex3d(new Point3d(Poly3D.GetPointAtParameter(param1p).X, Poly3D.GetPointAtParameter(param1p).Y, Poly3D.GetPointAtParameter(param1p).Z));
                                                        Poly3D_eq.AppendVertex(Vertex_last);
                                                        Trans1.AddNewlyCreatedDBObject(Vertex_last, true);
                                                    }
                                                    for (int k = Convert.ToInt32(Math.Ceiling(param1p)); k <= last_param; ++k)
                                                    {
                                                        PolylineVertex3d Vertex_eq = new PolylineVertex3d(new Point3d(Poly3D.GetPointAtParameter(k).X, Poly3D.GetPointAtParameter(k).Y, Poly3D.GetPointAtParameter(k).Z));
                                                        Poly3D_eq.AppendVertex(Vertex_eq);
                                                        Trans1.AddNewlyCreatedDBObject(Vertex_eq, true);
                                                    }
                                                }
                                                else
                                                {

                                                    Poly3D_eq = null;

                                                    int r = 0;
                                                    if (Math.Ceiling(param1p) != param1p)
                                                    {
                                                        Point2d p2d = new Point2d(Poly2D.GetPointAtParameter(param1p).X, Poly2D.GetPointAtParameter(param1p).Y);
                                                        double b_start = 0;
                                                        Poly2D_eq.AddVertexAt(r, p2d, b_start, 0, 0);
                                                        ++r;

                                                    }

                                                    for (int k = Convert.ToInt32(Math.Ceiling(param1p)); k <= last_param; ++k)
                                                    {
                                                        Point2d p2d = new Point2d(Poly2D.GetPointAtParameter(k).X, Poly2D.GetPointAtParameter(k).Y);
                                                        Poly2D_eq.AddVertexAt(r, p2d, Poly2D.GetBulgeAt(k), 0, 0);
                                                        ++r;

                                                    }


                                                }

                                                is_equated = true;
                                                create_stationing(Poly3D_eq, Poly2D_eq, od_name, SAp, gap1, texth, spacing_major, spacing_minor, tick_major, tick_minor);
                                            }


                                        }
                                    }

                                    if (is_equated == false)
                                    {

                                        create_stationing(Poly3D, Poly2D, od_name, start1, gap1, texth, spacing_major, spacing_minor, tick_major, tick_minor);
                                        add_zero_plus_zero_zero_stationing(Poly2D, od_name, 0, gap1, texth, tick_major, 0);

                                        if (_AGEN_mainform.Project_type == "3D")
                                        {
                                            add_last_stationing(Poly2D, od_name, Poly3D.Length - 0.000001, gap1, texth, tick_major, 0);

                                        }
                                        else
                                        {
                                            add_last_stationing(Poly2D, od_name, Poly2D.Length - 0.000001, gap1, texth, tick_major, 0);

                                        }

                                    }


                                }
                                #endregion
                                if (_AGEN_mainform.Project_type == "3D")
                                {

                                }
                                else
                                {

                                    BTrecord.AppendEntity(Poly2D);
                                    Trans1.AddNewlyCreatedDBObject(Poly2D, true);
                                    Functions.Populate_object_data_table_from_objectid(Tables1, Poly2D.ObjectId, od_name, Lista_val_CL, Lista_type_CL);
                                }

                                #region CANADA
                                if (_AGEN_mainform.COUNTRY == "CANADA")
                                {
                                    if (_AGEN_mainform.dt_centerline != null)
                                    {
                                        if (_AGEN_mainform.dt_centerline.Rows.Count > 1)
                                        {

                                            sort_canadian_station_eq(Poly2D);

                                            string Col_x = "X";
                                            string Col_y = "Y";
                                            string Col_3DSta = "3DSta";
                                            string Col_back = "BackSta";
                                            string Col_ahead = "AheadSta";


                                            Create_csf_cl_od_table();

                                            List<object> Lista_val = new List<object>();
                                            List<Autodesk.Gis.Map.Constants.DataType> Lista_type = new List<Autodesk.Gis.Map.Constants.DataType>();

                                            Lista_val.Add(_AGEN_mainform.version);
                                            Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                            string segment1 = _AGEN_mainform.tpage_setup.Get_segment_name1();
                                            if (segment1 == "not defined") segment1 = "";
                                            Lista_val.Add(segment1);
                                            Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                            Lista_val.Add(System.DateTime.Today.Month + "/" + System.DateTime.Today.Day + "/" + System.DateTime.Today.Year + " at " + System.DateTime.Now.Hour + ":" + System.DateTime.Now.Minute + " by " + Environment.UserName.ToUpper());
                                            Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                            Functions.Populate_object_data_table_from_objectid(Tables1, Poly3D.ObjectId, "Agen_csf_cl", Lista_val, Lista_type);


                                            if ((_AGEN_mainform.dt_centerline.Rows[0][Col_3DSta] != DBNull.Value ||
                                                _AGEN_mainform.dt_centerline.Rows[0][Col_ahead] != DBNull.Value) &&
                                                (Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[0][Col_3DSta])) == true ||
                                                Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[0][Col_ahead])) == true))
                                            {
                                                double sta0 = -1;
                                                if (_AGEN_mainform.dt_centerline.Rows[0][Col_3DSta] != DBNull.Value &&
                                                   Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[0][Col_3DSta])) == true)
                                                {
                                                    sta0 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[0][Col_3DSta]);
                                                }

                                                add_first_station_for_canada(Poly2D, od_name, sta0, gap1, texth, tick_major, 2);
                                            }

                                            if ((_AGEN_mainform.dt_centerline.Rows[_AGEN_mainform.dt_centerline.Rows.Count - 1][Col_3DSta] != DBNull.Value ||
                                                _AGEN_mainform.dt_centerline.Rows[_AGEN_mainform.dt_centerline.Rows.Count - 1][Col_ahead] != DBNull.Value) &&
                                                (Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[_AGEN_mainform.dt_centerline.Rows.Count - 1][Col_3DSta])) == true ||
                                                Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[_AGEN_mainform.dt_centerline.Rows.Count - 1][Col_ahead])) == true))
                                            {
                                                double end0 = -1;
                                                if (_AGEN_mainform.dt_centerline.Rows[_AGEN_mainform.dt_centerline.Rows.Count - 1][Col_3DSta] != DBNull.Value &&
                                                   Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[_AGEN_mainform.dt_centerline.Rows.Count - 1][Col_3DSta])) == true)
                                                {
                                                    end0 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[_AGEN_mainform.dt_centerline.Rows.Count - 1][Col_3DSta]);
                                                }
                                                add_last_stationing(Poly2D, od_name, end0, gap1, texth, tick_major, 2);
                                            }

                                            for (int i = 1; i < _AGEN_mainform.dt_centerline.Rows.Count; ++i)
                                            {
                                                if (_AGEN_mainform.dt_centerline.Rows[i - 1][Col_x] != DBNull.Value &&
                                                        _AGEN_mainform.dt_centerline.Rows[i - 1][Col_y] != DBNull.Value &&
                                                        (
                                                            _AGEN_mainform.dt_centerline.Rows[i - 1][Col_3DSta] != DBNull.Value ||
                                                            _AGEN_mainform.dt_centerline.Rows[i - 1][Col_ahead] != DBNull.Value
                                                        ) &&
                                                        Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[i - 1][Col_x])) == true &&
                                                        Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[i - 1][Col_y])) == true &&
                                                        (
                                                            Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[i - 1][Col_3DSta])) == true ||
                                                            Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[i - 1][Col_ahead])) == true
                                                        ) &&
                                                        _AGEN_mainform.dt_centerline.Rows[i][Col_x] != DBNull.Value &&
                                                        _AGEN_mainform.dt_centerline.Rows[i][Col_y] != DBNull.Value &&
                                                        (
                                                            _AGEN_mainform.dt_centerline.Rows[i][Col_3DSta] != DBNull.Value ||
                                                            _AGEN_mainform.dt_centerline.Rows[i][Col_back] != DBNull.Value
                                                        ) &&
                                                        Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[i][Col_x])) == true &&
                                                        Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[i][Col_y])) == true &&
                                                        (
                                                            Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[i][Col_3DSta])) == true ||
                                                            Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[i][Col_back])) == true
                                                        )
                                                    )
                                                {
                                                    double x0 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i - 1][Col_x]);
                                                    double y0 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i - 1][Col_y]);
                                                    double sta0 = -1;
                                                    if (_AGEN_mainform.dt_centerline.Rows[i - 1][Col_3DSta] != DBNull.Value &&
                                                       Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[i - 1][Col_3DSta])) == true)
                                                    {
                                                        sta0 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i - 1][Col_3DSta]);
                                                    }

                                                    double x1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i][Col_x]);
                                                    double y1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i][Col_y]);
                                                    double sta1 = -1;

                                                    if (_AGEN_mainform.dt_centerline.Rows[i][Col_3DSta] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[i][Col_3DSta])) == true)
                                                    {
                                                        sta1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i][Col_3DSta]);
                                                    }

                                                    Point3d pt0 = new Point3d(x0, y0, 0);
                                                    Point3d pt1 = new Point3d(x1, y1, 0);




                                                    create_stationing_with_csf(od_name, sta0, sta1, pt0, pt1, gap1, texth, spacing_major, spacing_minor, tick_major, tick_minor, i);



                                                }
                                            }
                                        }
                                    }
                                }
                                #endregion
                                Trans1.Commit();

                                if (Functions.is_dan_popescu() == true)
                                {
                                    //Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(_AGEN_mainform.dt_station_equation);
                                }
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                _AGEN_mainform.tpage_processing.Hide();
                set_enable_true();
            }

        }


        private void create_stationing(Polyline3d Poly3D, Polyline Poly2D, string od_name, double start1, double gap1, double texth, double spacing_major, double spacing_minor, double tick_major, double tick_minor)
        {
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                    #region object data stationing

                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;



                    List<object> Lista_val = new List<object>();
                    List<Autodesk.Gis.Map.Constants.DataType> Lista_type = new List<Autodesk.Gis.Map.Constants.DataType>();

                    Lista_val.Add(comboBox_segment_name.Text);
                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                    Lista_val.Add(System.DateTime.Today.Year + "-" + System.DateTime.Today.Month + "-" + System.DateTime.Today.Day + " at " + System.DateTime.Now.Hour + ":" + System.DateTime.Now.Minute);
                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);


                    #endregion

                    double max_length = Poly2D.Length;

                    if (_AGEN_mainform.Project_type == "3D")
                    {
                        max_length = Poly3D.Length;
                    }

                    int lr = 1;
                    double extra_rot = 0;
                    if (_AGEN_mainform.Left_to_Right == false)
                    {
                        lr = -1;
                        extra_rot = Math.PI;
                    }
                    double first_label_major = Math.Floor((start1 + spacing_major) / spacing_major) * spacing_major;

                    if (checkBox_create_major_labels.Checked == true)
                    {
                        if (start1 + max_length >= first_label_major)
                        {
                            int no_major = Convert.ToInt32(Math.Ceiling((start1 + max_length - first_label_major) / spacing_major));

                            if (no_major > 0)
                            {
                                for (int i = 0; i < no_major; ++i)
                                {
                                    Point3d pt0 = new Point3d();

                                    double dist1 = (first_label_major - start1) + i * spacing_major;
                                    double dist1_2d = dist1;
                                    if (_AGEN_mainform.Project_type == "3D")
                                    {
                                        pt0 = Poly3D.GetPointAtDist(dist1);
                                        double param_for_2D = Poly3D.GetParameterAtDistance(dist1);
                                        if (Poly2D.EndParam < param_for_2D)
                                        {
                                            param_for_2D = Poly2D.EndParam;
                                        }
                                        dist1_2d = Poly2D.GetDistanceAtParameter(param_for_2D);

                                    }
                                    else
                                    {
                                        pt0 = Poly2D.GetPointAtDist(dist1);

                                    }



                                    Line Big1 = new Line(new Point3d(pt0.X - tick_major / 2, pt0.Y, 0), new Point3d(pt0.X + tick_major / 2, pt0.Y, 0));

                                    double param1 = Poly2D.GetParameterAtDistance(dist1_2d);
                                    double param2 = param1 + 1;
                                    if (Poly2D.EndParam < param2)
                                    {
                                        param1 = Poly2D.EndParam - 1;
                                        param2 = Poly2D.EndParam;
                                    }

                                    double bulge = 0;

                                    int par_bulge = Convert.ToInt32(Math.Floor(param1));


                                    Point3d point1 = new Point3d();

                                    Point3d point2 = new Point3d();

                                    double bear1 = -1;

                                    double rot1 = -1;

                                    if (_AGEN_mainform.Project_type == "3D")

                                    {
                                        point1 = Poly3D.GetPointAtParameter(Math.Floor(param1));
                                        point2 = Poly3D.GetPointAtParameter(Math.Floor(param2));

                                        bear1 = Functions.GET_Bearing_rad(point1.X, point1.Y, point2.X, point2.Y);

                                        rot1 = bear1 - lr * Math.PI / 2;
                                    }
                                    else
                                    {
                                        bulge = Poly2D.GetBulgeAt(par_bulge);
                                        point1 = Poly2D.GetPointAtParameter(par_bulge);
                                        point2 = Poly2D.GetPointAtParameter(Math.Floor(param2));

                                        if (bulge != 0)
                                        {
                                            CircularArc3d arc3d = Poly2D.GetArcSegmentAt(par_bulge);
                                            double bear_to_center = Functions.GET_Bearing_rad(pt0.X, pt0.Y, arc3d.Center.X, arc3d.Center.Y);
                                            if (bulge > 0)
                                            {
                                                if (lr == 1)
                                                {
                                                    rot1 = Math.PI + bear_to_center;
                                                }
                                                else
                                                {
                                                    rot1 = bear_to_center;
                                                }

                                            }
                                            else
                                            {

                                                if (lr == 1)
                                                {
                                                    rot1 = bear_to_center;
                                                }
                                                else
                                                {
                                                    rot1 = Math.PI + bear_to_center;
                                                }

                                            }

                                            bear1 = rot1 + lr * Math.PI / 2;

                                        }
                                        else
                                        {

                                            bear1 = Functions.GET_Bearing_rad(point1.X, point1.Y, point2.X, point2.Y);

                                            rot1 = bear1 - lr * Math.PI / 2;
                                        }

                                    }



                                    Big1.TransformBy(Matrix3d.Rotation(rot1, Vector3d.ZAxis, pt0));

                                    Big1.Layer = _AGEN_mainform.layer_stationing;
                                    Big1.ColorIndex = 256;


                                    BTrecord.AppendEntity(Big1);
                                    Trans1.AddNewlyCreatedDBObject(Big1, true);
                                    Functions.Populate_object_data_table_from_objectid(Tables1, Big1.ObjectId, od_name, Lista_val, Lista_type);


                                    if (checkBox_create_major_labels.Checked == true)
                                    {
                                        MText mt1 = new MText();
                                        if (checkBox_text_bottom.Checked == false)
                                        {
                                            Line temp_line1 = new Line(Big1.StartPoint, Big1.EndPoint);
                                            temp_line1.TransformBy(Matrix3d.Scaling((Big1.Length + gap1) / Big1.Length, Big1.EndPoint));
                                            mt1 = creaza_mtext_sta_bottom_center(temp_line1.StartPoint, Functions.Get_chainage_from_double(first_label_major + i * spacing_major, _AGEN_mainform.units_of_measurement, 0), texth, bear1 + extra_rot);
                                        }
                                        else
                                        {
                                            Line temp_line2 = new Line(Big1.StartPoint, Big1.EndPoint);
                                            temp_line2.TransformBy(Matrix3d.Scaling((Big1.Length + gap1) / Big1.Length, Big1.StartPoint));
                                            mt1 = creaza_mtext_sta_top_center(temp_line2.EndPoint, Functions.Get_chainage_from_double(first_label_major + i * spacing_major, _AGEN_mainform.units_of_measurement, 0), texth, bear1 + extra_rot);
                                        }

                                        mt1.Layer = _AGEN_mainform.layer_stationing;
                                        BTrecord.AppendEntity(mt1);
                                        Trans1.AddNewlyCreatedDBObject(mt1, true);
                                        Functions.Populate_object_data_table_from_objectid(Tables1, mt1.ObjectId, od_name, Lista_val, Lista_type);
                                    }

                                }
                            }
                        }
                    }




                    if (checkBox_draw_minor_lines.Checked == true)
                    {
                        double first_label_minor = Math.Floor((start1 + spacing_minor) / spacing_minor) * spacing_minor;

                        if (start1 + max_length >= first_label_minor)
                        {
                            int no_minor = Convert.ToInt32(Math.Ceiling((start1 + max_length - first_label_minor) / spacing_minor));

                            if (no_minor > 0)
                            {
                                for (int i = 0; i < no_minor; ++i)
                                {
                                    Point3d pt0 = new Point3d();

                                    if (_AGEN_mainform.Project_type == "3D")
                                    {
                                        pt0 = Poly3D.GetPointAtDist((first_label_minor - start1) + i * spacing_minor);
                                    }
                                    else
                                    {
                                        pt0 = Poly2D.GetPointAtDist((first_label_minor - start1) + i * spacing_minor);
                                    }


                                    Line small1 = new Line(new Point3d(pt0.X - tick_minor / 2, pt0.Y, 0), new Point3d(pt0.X + tick_minor / 2, pt0.Y, 0));

                                    double param1 = Poly2D.GetParameterAtPoint(Poly2D.GetClosestPointTo(pt0, Vector3d.ZAxis, false));
                                    double param2 = param1 + 1;
                                    if (Poly2D.EndParam < param2)
                                    {
                                        param1 = Poly2D.EndParam - 1;
                                        param2 = Poly2D.EndParam;
                                    }


                                    double bulge = 0;

                                    int par_bulge = Convert.ToInt32(Math.Floor(param1));


                                    Point3d point1 = new Point3d();

                                    Point3d point2 = new Point3d();


                                    double bear1 = -1;

                                    double rot1 = -1;

                                    if (_AGEN_mainform.Project_type == "3D")

                                    {
                                        point1 = Poly3D.GetPointAtParameter(Math.Floor(param1));
                                        point2 = Poly3D.GetPointAtParameter(Math.Floor(param2));

                                        bear1 = Functions.GET_Bearing_rad(point1.X, point1.Y, point2.X, point2.Y);

                                        rot1 = bear1 - lr * Math.PI / 2;
                                    }
                                    else
                                    {
                                        bulge = Poly2D.GetBulgeAt(par_bulge);
                                        point1 = Poly2D.GetPointAtParameter(par_bulge);
                                        point2 = Poly2D.GetPointAtParameter(Math.Floor(param2));

                                        if (bulge != 0)
                                        {
                                            CircularArc3d arc3d = Poly2D.GetArcSegmentAt(par_bulge);
                                            double bear_to_center = Functions.GET_Bearing_rad(pt0.X, pt0.Y, arc3d.Center.X, arc3d.Center.Y);

                                            if (bulge > 0)
                                            {
                                                if (lr == 1)
                                                {
                                                    rot1 = Math.PI + bear_to_center;
                                                }
                                                else
                                                {
                                                    rot1 = bear_to_center;
                                                }

                                            }
                                            else
                                            {

                                                if (lr == 1)
                                                {
                                                    rot1 = bear_to_center;
                                                }
                                                else
                                                {
                                                    rot1 = Math.PI + bear_to_center;
                                                }

                                            }

                                            bear1 = rot1 + lr * Math.PI / 2;

                                        }
                                        else
                                        {

                                            bear1 = Functions.GET_Bearing_rad(point1.X, point1.Y, point2.X, point2.Y);

                                            rot1 = bear1 - lr * Math.PI / 2;
                                        }

                                    }


                                    small1.TransformBy(Matrix3d.Rotation(rot1, Vector3d.ZAxis, pt0));

                                    small1.Layer = _AGEN_mainform.layer_stationing;
                                    small1.ColorIndex = 256;


                                    BTrecord.AppendEntity(small1);
                                    Trans1.AddNewlyCreatedDBObject(small1, true);

                                    if (checkBox_create_minor_labels.Checked == true)
                                    {


                                        MText mt1 = new MText();
                                        if (checkBox_text_bottom.Checked == false)
                                        {
                                            Line temp_line1 = new Line(small1.StartPoint, small1.EndPoint);
                                            temp_line1.TransformBy(Matrix3d.Scaling((small1.Length + gap1) / small1.Length, small1.EndPoint));
                                            mt1 = creaza_mtext_sta_bottom_center(temp_line1.StartPoint, Functions.Get_chainage_from_double(first_label_minor + i * spacing_minor, _AGEN_mainform.units_of_measurement, 0), texth, bear1 + extra_rot);
                                        }
                                        else
                                        {
                                            Line temp_line2 = new Line(small1.StartPoint, small1.EndPoint);
                                            temp_line2.TransformBy(Matrix3d.Scaling((small1.Length + gap1) / small1.Length, small1.StartPoint));
                                            mt1 = creaza_mtext_sta_top_center(temp_line2.EndPoint, Functions.Get_chainage_from_double(first_label_minor + i * spacing_minor, _AGEN_mainform.units_of_measurement, 0), texth, bear1 + extra_rot);
                                        }





                                        mt1.Layer = _AGEN_mainform.layer_stationing;
                                        BTrecord.AppendEntity(mt1);
                                        Trans1.AddNewlyCreatedDBObject(mt1, true);
                                        Functions.Populate_object_data_table_from_objectid(Tables1, mt1.ObjectId, od_name, Lista_val, Lista_type);
                                    }



                                }
                            }
                        }
                    }




                    if (Poly3D != null) Poly3D.Erase();


                    Trans1.Commit();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }



        private void add_zero_plus_zero_zero_stationing(Polyline Poly2D, string od_name, double start1, double gap1, double texth, double thick, int round1)
        {
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);


                    #region object data stationing
                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;


                    List<object> Lista_val = new List<object>();
                    List<Autodesk.Gis.Map.Constants.DataType> Lista_type = new List<Autodesk.Gis.Map.Constants.DataType>();

                    Lista_val.Add(comboBox_segment_name.Text);
                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                    Lista_val.Add(System.DateTime.Today.Year + "-" + System.DateTime.Today.Month + "-" + System.DateTime.Today.Day + " at " + System.DateTime.Now.Hour + ":" + System.DateTime.Now.Minute);
                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                    #endregion

                    int lr = 1;
                    double extra_rot = 0;
                    if (_AGEN_mainform.Left_to_Right == false)
                    {
                        lr = -1;
                        extra_rot = Math.PI;
                    }

                    Point3d pt1 = Poly2D.GetPointAtDist(0);
                    Point3d pt2 = Poly2D.GetPointAtParameter(1);
                    Line line_0 = new Line(new Point3d(pt1.X - thick / 2, pt1.Y, 0), new Point3d(pt1.X + thick / 2, pt1.Y, 0));

                    double bear1 = Functions.GET_Bearing_rad(pt1.X, pt1.Y, pt2.X, pt2.Y);
                    double rot1 = bear1 - lr * Math.PI / 2;

                    line_0.TransformBy(Matrix3d.Rotation(rot1, Vector3d.ZAxis, pt1));
                    line_0.Layer = _AGEN_mainform.layer_stationing;
                    line_0.ColorIndex = 256;

                    BTrecord.AppendEntity(line_0);
                    Trans1.AddNewlyCreatedDBObject(line_0, true);
                    Functions.Populate_object_data_table_from_objectid(Tables1, line_0.ObjectId, od_name, Lista_val, Lista_type);


                    MText mt1 = new MText();
                    if (checkBox_text_bottom.Checked == false)
                    {
                        Line temp_line1 = new Line(line_0.StartPoint, line_0.EndPoint);
                        temp_line1.TransformBy(Matrix3d.Scaling((line_0.Length + gap1) / line_0.Length, line_0.EndPoint));
                        mt1 = creaza_mtext_sta_bottom_center(temp_line1.StartPoint, Functions.Get_chainage_from_double(start1, _AGEN_mainform.units_of_measurement, round1), texth, bear1 + extra_rot);

                    }
                    else
                    {
                        Line temp_line2 = new Line(line_0.StartPoint, line_0.EndPoint);
                        temp_line2.TransformBy(Matrix3d.Scaling((line_0.Length + gap1) / line_0.Length, line_0.StartPoint));
                        mt1 = creaza_mtext_sta_top_center(temp_line2.EndPoint, Functions.Get_chainage_from_double(start1, _AGEN_mainform.units_of_measurement, round1), texth, bear1 + extra_rot);
                    }

                    mt1.Layer = _AGEN_mainform.layer_stationing;
                    BTrecord.AppendEntity(mt1);
                    Trans1.AddNewlyCreatedDBObject(mt1, true);
                    Functions.Populate_object_data_table_from_objectid(Tables1, mt1.ObjectId, od_name, Lista_val, Lista_type);



                    Trans1.Commit();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        private void add_last_stationing(Polyline Poly2D, string od_name, double end1, double gap1, double texth, double thick, int round1)
        {
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                    #region object data stationing
                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;


                    List<object> Lista_val = new List<object>();
                    List<Autodesk.Gis.Map.Constants.DataType> Lista_type = new List<Autodesk.Gis.Map.Constants.DataType>();

                    Lista_val.Add(comboBox_segment_name.Text);
                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                    Lista_val.Add(System.DateTime.Today.Year + "-" + System.DateTime.Today.Month + "-" + System.DateTime.Today.Day + " at " + System.DateTime.Now.Hour + ":" + System.DateTime.Now.Minute);
                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                    #endregion

                    int lr = 1;
                    double extra_rot = 0;
                    if (_AGEN_mainform.Left_to_Right == false)
                    {
                        lr = -1;
                        extra_rot = Math.PI;
                    }

                    Point3d pt1 = Poly2D.GetPointAtParameter(Poly2D.NumberOfVertices - 2);
                    Point3d pt2 = Poly2D.EndPoint;
                    Line line_0 = new Line(new Point3d(pt2.X - thick / 2, pt2.Y, 0), new Point3d(pt2.X + thick / 2, pt2.Y, 0));

                    double bear1 = Functions.GET_Bearing_rad(pt1.X, pt1.Y, pt2.X, pt2.Y);
                    double rot1 = bear1 - lr * Math.PI / 2;

                    line_0.TransformBy(Matrix3d.Rotation(rot1, Vector3d.ZAxis, pt2));
                    line_0.Layer = _AGEN_mainform.layer_stationing;
                    line_0.ColorIndex = 256;

                    BTrecord.AppendEntity(line_0);
                    Trans1.AddNewlyCreatedDBObject(line_0, true);
                    Functions.Populate_object_data_table_from_objectid(Tables1, line_0.ObjectId, od_name, Lista_val, Lista_type);

                    MText mt1 = new MText();
                    if (checkBox_text_bottom.Checked == false)
                    {
                        Line temp_line1 = new Line(line_0.StartPoint, line_0.EndPoint);
                        temp_line1.TransformBy(Matrix3d.Scaling((line_0.Length + gap1) / line_0.Length, line_0.EndPoint));
                        mt1 = creaza_mtext_sta_bottom_center(temp_line1.StartPoint, Functions.Get_chainage_from_double(end1, _AGEN_mainform.units_of_measurement, round1), texth, bear1 + extra_rot);
                    }
                    else
                    {
                        Line temp_line2 = new Line(line_0.StartPoint, line_0.EndPoint);
                        temp_line2.TransformBy(Matrix3d.Scaling((line_0.Length + gap1) / line_0.Length, line_0.StartPoint));
                        mt1 = creaza_mtext_sta_top_center(temp_line2.EndPoint, Functions.Get_chainage_from_double(end1, _AGEN_mainform.units_of_measurement, round1), texth, bear1 + extra_rot);
                    }



                    mt1.Layer = _AGEN_mainform.layer_stationing;
                    BTrecord.AppendEntity(mt1);
                    Trans1.AddNewlyCreatedDBObject(mt1, true);
                    Functions.Populate_object_data_table_from_objectid(Tables1, mt1.ObjectId, od_name, Lista_val, Lista_type);

                    Trans1.Commit();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void add_first_station_for_canada(Polyline Poly2D, string od_name, double start1, double gap1, double texth, double thick, int round1)
        {
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                    #region object data stationing
                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;


                    List<object> Lista_val = new List<object>();
                    List<Autodesk.Gis.Map.Constants.DataType> Lista_type = new List<Autodesk.Gis.Map.Constants.DataType>();

                    Lista_val.Add(comboBox_segment_name.Text);
                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                    Lista_val.Add(System.DateTime.Today.Year + "-" + System.DateTime.Today.Month + "-" + System.DateTime.Today.Day + " at " + System.DateTime.Now.Hour + ":" + System.DateTime.Now.Minute);
                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                    #endregion


                    if (_AGEN_mainform.dt_station_equation != null && _AGEN_mainform.dt_station_equation.Rows.Count > 0)
                    {

                        for (int i = 0; i < _AGEN_mainform.dt_station_equation.Rows.Count; ++i)
                        {
                            if (_AGEN_mainform.dt_station_equation.Rows[i][rr_end_x] != DBNull.Value && _AGEN_mainform.dt_station_equation.Rows[i][rr_end_y] != DBNull.Value && _AGEN_mainform.dt_station_equation.Rows[i][sta_ahead] != DBNull.Value)
                            {
                                double x = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i][rr_end_x]);
                                double y = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i][rr_end_y]);
                                Point3d pt_start = new Point3d(x, y, 0);
                                Point3d pt_poly_start = Poly2D.StartPoint;
                                double dist1 = Math.Pow(Math.Pow(pt_start.X - pt_poly_start.X, 2) + Math.Pow(pt_start.Y - pt_poly_start.Y, 2), 0.5);

                                if (dist1 < 0.001)
                                {
                                    start1 = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i][sta_ahead]);
                                    i = _AGEN_mainform.dt_station_equation.Rows.Count;
                                }
                            }
                        }
                    }

                    int lr = 1;
                    double extra_rot = 0;
                    if (_AGEN_mainform.Left_to_Right == false)
                    {
                        lr = -1;
                        extra_rot = Math.PI;
                    }

                    Point3d pt1 = Poly2D.GetPointAtDist(0);
                    Point3d pt2 = Poly2D.GetPointAtParameter(1);
                    Line line_0 = new Line(new Point3d(pt1.X - thick / 2, pt1.Y, 0), new Point3d(pt1.X + thick / 2, pt1.Y, 0));

                    double bear1 = Functions.GET_Bearing_rad(pt1.X, pt1.Y, pt2.X, pt2.Y);
                    double rot1 = bear1 - lr * Math.PI / 2;

                    line_0.TransformBy(Matrix3d.Rotation(rot1, Vector3d.ZAxis, pt1));
                    line_0.Layer = _AGEN_mainform.layer_stationing;
                    line_0.ColorIndex = 256;

                    BTrecord.AppendEntity(line_0);
                    Trans1.AddNewlyCreatedDBObject(line_0, true);
                    Functions.Populate_object_data_table_from_objectid(Tables1, line_0.ObjectId, od_name, Lista_val, Lista_type);


                    Line l_t = new Line(line_0.StartPoint, line_0.EndPoint);
                    l_t.TransformBy(Matrix3d.Scaling((line_0.Length + gap1) / line_0.Length, line_0.EndPoint));

                    MText mt1 = creaza_mtext_sta_bottom_center(l_t.StartPoint, Functions.Get_chainage_from_double(start1, _AGEN_mainform.units_of_measurement, round1), texth, bear1 + extra_rot);

                    mt1.Layer = _AGEN_mainform.layer_stationing;
                    BTrecord.AppendEntity(mt1);
                    Trans1.AddNewlyCreatedDBObject(mt1, true);
                    Functions.Populate_object_data_table_from_objectid(Tables1, mt1.ObjectId, od_name, Lista_val, Lista_type);



                    Trans1.Commit();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }



        private void create_stationing_with_csf(string od_name, double sta0, double Sta1, Point3d pt0, Point3d pt1, double gap1, double texth, double spacing_major, double spacing_minor, double tick_major, double tick_minor, int index)
        {
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                    Polyline Poly2D = new Polyline();
                    Poly2D.AddVertexAt(0, new Point2d(pt0.X, pt0.Y), 0, 0, 0);
                    Poly2D.AddVertexAt(1, new Point2d(pt1.X, pt1.Y), 0, 0, 0);
                    #region object data stationing

                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;



                    List<object> Lista_val = new List<object>();
                    List<Autodesk.Gis.Map.Constants.DataType> Lista_type = new List<Autodesk.Gis.Map.Constants.DataType>();

                    Lista_val.Add(comboBox_segment_name.Text);
                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                    Lista_val.Add(System.DateTime.Today.Year + "-" + System.DateTime.Today.Month + "-" + System.DateTime.Today.Day + " at " + System.DateTime.Now.Hour + ":" + System.DateTime.Now.Minute);
                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);


                    #endregion

                    int lr = 1;
                    double extra_rot = 0;
                    if (_AGEN_mainform.Left_to_Right == false)
                    {
                        lr = -1;
                        extra_rot = Math.PI;
                    }


                    if (_AGEN_mainform.dt_station_equation != null && _AGEN_mainform.dt_station_equation.Rows.Count > 0)
                    {

                        for (int i = 0; i < _AGEN_mainform.dt_station_equation.Rows.Count; ++i)
                        {

                            if (_AGEN_mainform.dt_station_equation.Rows[i][rr_end_x] != DBNull.Value && _AGEN_mainform.dt_station_equation.Rows[i][rr_end_y] != DBNull.Value &&
                                _AGEN_mainform.dt_station_equation.Rows[i][sta_back] != DBNull.Value && _AGEN_mainform.dt_station_equation.Rows[i][sta_ahead] != DBNull.Value)
                            {
                                double x = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i][rr_end_x]);
                                double y = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i][rr_end_y]);
                                Point3d pt_eq = new Point3d(x, y, 0);

                                Point3d pt_on_poly = Poly2D.GetClosestPointTo(pt_eq, Vector3d.ZAxis, false);
                                double d1 = Math.Pow(Math.Pow(pt_eq.X - pt_on_poly.X, 2) + Math.Pow(pt_eq.Y - pt_on_poly.Y, 2), 0.5);
                                double meas1 = Poly2D.GetDistAtPoint(pt_on_poly);
                                if (d1 < 0.001 && meas1 > 0 && meas1 <= Poly2D.Length)
                                {
                                    double back1 = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i][sta_back]);
                                    double ahead1 = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i][sta_ahead]);

                                    double fl_major = Math.Ceiling(sta0 / spacing_major) * spacing_major;
                                    if (fl_major <= back1)
                                    {

                                        int no_major = Convert.ToInt32(Math.Ceiling((back1 - fl_major) / spacing_major));

                                        if (no_major > 0)
                                        {
                                            for (int k = 0; k < no_major; ++k)
                                            {
                                                double dist1 = (fl_major - sta0 + k * spacing_major) * Poly2D.Length / (back1 - sta0);

                                                Point3d pos1 = Poly2D.GetPointAtDist((dist1));
                                                double label_major = fl_major + k * spacing_major;
                                                Line Big1 = new Line(new Point3d(pos1.X - tick_major / 2, pos1.Y, 0), new Point3d(pos1.X + tick_major / 2, pos1.Y, 0));

                                                double bear1 = Functions.GET_Bearing_rad(pt0.X, pt0.Y, pt1.X, pt1.Y);

                                                double rot1 = bear1 - lr * Math.PI / 2;

                                                Big1.TransformBy(Matrix3d.Rotation(rot1, Vector3d.ZAxis, pos1));

                                                Big1.Layer = _AGEN_mainform.layer_stationing;
                                                Big1.ColorIndex = 256;

                                                if (checkBox_draw_major_lines.Checked == true)
                                                {
                                                    BTrecord.AppendEntity(Big1);
                                                    Trans1.AddNewlyCreatedDBObject(Big1, true);
                                                    Functions.Populate_object_data_table_from_objectid(Tables1, Big1.ObjectId, od_name, Lista_val, Lista_type);
                                                }

                                                if (checkBox_create_major_labels.Checked == true)
                                                {
                                                    Line l_t = new Line(Big1.StartPoint, Big1.EndPoint);
                                                    l_t.TransformBy(Matrix3d.Scaling((Big1.Length + gap1) / Big1.Length, Big1.EndPoint));

                                                    MText mt1 = creaza_mtext_sta_bottom_center(l_t.StartPoint, Functions.Get_chainage_from_double(fl_major + k * spacing_major, _AGEN_mainform.units_of_measurement, 2), texth, bear1 + extra_rot);

                                                    mt1.Layer = _AGEN_mainform.layer_stationing;
                                                    BTrecord.AppendEntity(mt1);
                                                    Trans1.AddNewlyCreatedDBObject(mt1, true);
                                                    Functions.Populate_object_data_table_from_objectid(Tables1, mt1.ObjectId, od_name, Lista_val, Lista_type);
                                                }

                                            }
                                        }
                                    }

                                    double fl_minor = Math.Ceiling(sta0 / spacing_minor) * spacing_minor;

                                    if (fl_minor <= back1)
                                    {

                                        int no_minor = Convert.ToInt32(Math.Ceiling((back1 - fl_minor) / spacing_minor));

                                        if (no_minor > 0)
                                        {
                                            for (int k = 0; k < no_minor; ++k)
                                            {
                                                double dist1 = (fl_minor - sta0 + k * spacing_minor) * Poly2D.Length / (back1 - sta0);

                                                Point3d pos1 = Poly2D.GetPointAtDist((dist1));
                                                double label_minor = fl_minor + k * spacing_minor;
                                                Line Big1 = new Line(new Point3d(pos1.X - tick_minor / 2, pos1.Y, 0), new Point3d(pos1.X + tick_minor / 2, pos1.Y, 0));


                                                double bear1 = Functions.GET_Bearing_rad(pt0.X, pt0.Y, pt1.X, pt1.Y);

                                                double rot1 = bear1 - lr * Math.PI / 2;

                                                Big1.TransformBy(Matrix3d.Rotation(rot1, Vector3d.ZAxis, pos1));

                                                Big1.Layer = _AGEN_mainform.layer_stationing;
                                                Big1.ColorIndex = 256;

                                                if (checkBox_draw_minor_lines.Checked == true)
                                                {
                                                    BTrecord.AppendEntity(Big1);
                                                    Trans1.AddNewlyCreatedDBObject(Big1, true);
                                                    Functions.Populate_object_data_table_from_objectid(Tables1, Big1.ObjectId, od_name, Lista_val, Lista_type);
                                                }

                                                if (checkBox_create_minor_labels.Checked == true)
                                                {
                                                    Line l_t = new Line(Big1.StartPoint, Big1.EndPoint);
                                                    l_t.TransformBy(Matrix3d.Scaling((Big1.Length + gap1) / Big1.Length, Big1.EndPoint));

                                                    MText mt1 = creaza_mtext_sta_bottom_center(l_t.StartPoint, Functions.Get_chainage_from_double(fl_minor + k * spacing_minor, _AGEN_mainform.units_of_measurement, 2), texth, bear1 + extra_rot);

                                                    mt1.Layer = _AGEN_mainform.layer_stationing;
                                                    BTrecord.AppendEntity(mt1);
                                                    Trans1.AddNewlyCreatedDBObject(mt1, true);
                                                    Functions.Populate_object_data_table_from_objectid(Tables1, mt1.ObjectId, od_name, Lista_val, Lista_type);
                                                }

                                            }
                                        }
                                    }
                                    sta0 = ahead1;
                                }
                            }
                        }

                    }

                    double first_label_major = Math.Ceiling(sta0 / spacing_major) * spacing_major;

                    if (first_label_major <= Sta1)
                    {

                        int no_major = Convert.ToInt32(Math.Ceiling((Sta1 - first_label_major) / spacing_major));

                        if (no_major > 0)
                        {
                            for (int i = 0; i < no_major; ++i)
                            {
                                double dist1 = (first_label_major - sta0 + i * spacing_major) * Poly2D.Length / (Sta1 - sta0);

                                Point3d pos1 = Poly2D.GetPointAtDist((dist1));
                                double label_major = first_label_major + i * spacing_major;
                                Line Big1 = new Line(new Point3d(pos1.X - tick_major / 2, pos1.Y, 0), new Point3d(pos1.X + tick_major / 2, pos1.Y, 0));

                                double bear1 = Functions.GET_Bearing_rad(pt0.X, pt0.Y, pt1.X, pt1.Y);

                                double rot1 = bear1 - lr * Math.PI / 2;

                                Big1.TransformBy(Matrix3d.Rotation(rot1, Vector3d.ZAxis, pos1));

                                Big1.Layer = _AGEN_mainform.layer_stationing;
                                Big1.ColorIndex = 256;

                                if (checkBox_draw_major_lines.Checked == true)
                                {
                                    BTrecord.AppendEntity(Big1);
                                    Trans1.AddNewlyCreatedDBObject(Big1, true);
                                    Functions.Populate_object_data_table_from_objectid(Tables1, Big1.ObjectId, od_name, Lista_val, Lista_type);
                                }

                                if (checkBox_create_major_labels.Checked == true)
                                {
                                    Line l_t = new Line(Big1.StartPoint, Big1.EndPoint);
                                    l_t.TransformBy(Matrix3d.Scaling((Big1.Length + gap1) / Big1.Length, Big1.EndPoint));

                                    MText mt1 = creaza_mtext_sta_bottom_center(l_t.StartPoint, Functions.Get_chainage_from_double(first_label_major + i * spacing_major, _AGEN_mainform.units_of_measurement, 2), texth, bear1 + extra_rot);

                                    mt1.Layer = _AGEN_mainform.layer_stationing;
                                    BTrecord.AppendEntity(mt1);
                                    Trans1.AddNewlyCreatedDBObject(mt1, true);
                                    Functions.Populate_object_data_table_from_objectid(Tables1, mt1.ObjectId, od_name, Lista_val, Lista_type);
                                }

                            }
                        }
                    }


                    double first_label_minor = Math.Ceiling(sta0 / spacing_minor) * spacing_minor;

                    if (first_label_minor <= Sta1)
                    {

                        int no_minor = Convert.ToInt32(Math.Ceiling((Sta1 - first_label_minor) / spacing_minor));

                        if (no_minor > 0)
                        {
                            for (int i = 0; i < no_minor; ++i)
                            {
                                double dist1 = (first_label_minor - sta0 + i * spacing_minor) * Poly2D.Length / (Sta1 - sta0);

                                Point3d pos1 = Poly2D.GetPointAtDist((dist1));
                                double label_minor = first_label_minor + i * spacing_minor;
                                Line Big1 = new Line(new Point3d(pos1.X - tick_minor / 2, pos1.Y, 0), new Point3d(pos1.X + tick_minor / 2, pos1.Y, 0));


                                double bear1 = Functions.GET_Bearing_rad(pt0.X, pt0.Y, pt1.X, pt1.Y);

                                double rot1 = bear1 - lr * Math.PI / 2;

                                Big1.TransformBy(Matrix3d.Rotation(rot1, Vector3d.ZAxis, pos1));

                                Big1.Layer = _AGEN_mainform.layer_stationing;
                                Big1.ColorIndex = 256;

                                if (checkBox_draw_minor_lines.Checked == true)
                                {
                                    BTrecord.AppendEntity(Big1);
                                    Trans1.AddNewlyCreatedDBObject(Big1, true);
                                    Functions.Populate_object_data_table_from_objectid(Tables1, Big1.ObjectId, od_name, Lista_val, Lista_type);
                                }

                                if (checkBox_create_minor_labels.Checked == true)
                                {
                                    Line l_t = new Line(Big1.StartPoint, Big1.EndPoint);
                                    l_t.TransformBy(Matrix3d.Scaling((Big1.Length + gap1) / Big1.Length, Big1.EndPoint));

                                    MText mt1 = creaza_mtext_sta_bottom_center(l_t.StartPoint, Functions.Get_chainage_from_double(first_label_minor + i * spacing_minor, _AGEN_mainform.units_of_measurement, 2), texth, bear1 + extra_rot);

                                    mt1.Layer = _AGEN_mainform.layer_stationing;
                                    BTrecord.AppendEntity(mt1);
                                    Trans1.AddNewlyCreatedDBObject(mt1, true);
                                    Functions.Populate_object_data_table_from_objectid(Tables1, mt1.ObjectId, od_name, Lista_val, Lista_type);
                                }

                            }
                        }
                    }




                    Trans1.Commit();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\ni=" + index.ToString() + "\r\npt0=" + pt0.ToString() + "\r\npt1=" + pt1.ToString());
            }


        }

        public void delete_entities_with_OD(string layer_name, string od_table_name)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock lock1 = ThisDrawing.LockDocument())
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.Database.TransactionManager.StartTransaction())
                {
                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                    foreach (ObjectId id1 in BTrecord)
                    {
                        Entity ent1 = Trans1.GetObject(id1, OpenMode.ForRead) as Entity;
                        if (ent1 != null)
                        {
                            if (ent1.Layer == layer_name)
                            {
                                Autodesk.Gis.Map.ObjectData.Records Records1;
                                bool delete1 = false;
                                if (Tables1.IsTableDefined(od_table_name) == true)
                                {
                                    Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[od_table_name];
                                    using (Records1 = Tabla1.GetObjectTableRecords(Convert.ToUInt32(0), ent1.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, true))
                                    {
                                        if (Records1.Count > 0)
                                        {
                                            Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;
                                            foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                            {
                                                if (delete1 == false)
                                                {
                                                    for (int i = 0; i < Record1.Count; ++i)
                                                    {
                                                        Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[i];
                                                        string Nume_field = Field_def1.Name;
                                                        string Valoare1 = Record1[i].StrValue;
                                                        if (Nume_field == "SegmentName")
                                                        {
                                                            string segment1 = _AGEN_mainform.tpage_setup.Get_segment_name1();
                                                            if (segment1 == "not defined") segment1 = "";
                                                            if (Valoare1 == segment1)
                                                            {
                                                                delete1 = true;
                                                                i = Record1.Count;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    if (delete1 == true)
                                    {
                                        ent1.UpgradeOpen();
                                        ent1.Erase();
                                    }
                                }
                            }
                        }
                    }
                    Trans1.Commit();
                }
            }
        }

        public static MText creaza_mtext_sta_bottom_center(Point3d pt_ins, string continut, double texth, double rot1)
        {
            MText mtext1 = new MText();
            mtext1.Attachment = AttachmentPoint.BottomCenter;
            mtext1.Contents = continut;
            mtext1.TextHeight = texth;
            mtext1.BackgroundFill = true;
            mtext1.UseBackgroundColor = true;
            mtext1.BackgroundScaleFactor = 1.2;
            mtext1.Location = pt_ins;
            mtext1.Rotation = rot1;
            mtext1.ColorIndex = 256;

            return mtext1;
        }

        public static MText creaza_mtext_sta_top_center(Point3d pt_ins, string continut, double texth, double rot1)
        {
            MText mtext1 = new MText();
            mtext1.Attachment = AttachmentPoint.TopCenter;
            mtext1.Contents = continut;
            mtext1.TextHeight = texth;
            mtext1.BackgroundFill = true;
            mtext1.UseBackgroundColor = true;
            mtext1.BackgroundScaleFactor = 1.2;
            mtext1.Location = pt_ins;
            mtext1.Rotation = rot1;
            mtext1.ColorIndex = 256;

            return mtext1;
        }

        private void button_eq_refresh_blocks_Click(object sender, EventArgs e)
        {
            Functions.Incarca_existing_Blocks_with_attributes_to_combobox(comboBox_eq_block);
        }

        private void comboBox_eq_block_SelectedIndexChanged(object sender, EventArgs e)
        {
            string nume1 = comboBox_eq_block.Text;
            Functions.Incarca_existing_Atributes_to_combobox(nume1, comboBox_eq_bs);
            Functions.Incarca_existing_Atributes_to_combobox(nume1, comboBox_eq_as);
            Functions.Incarca_existing_Atributes_to_combobox(nume1, comboBox_eq_diff);
        }

        private void button_eq_insert_block_Click(object sender, EventArgs e)
        {
            Ag = this.MdiParent as _AGEN_mainform;
            Functions.Kill_excel();

            if (Functions.Get_if_workbook_is_open_in_Excel("centerline.xlsx") == true)
            {
                MessageBox.Show("Please close the centerline file");
                return;
            }


            string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();

            if (System.IO.Directory.Exists(ProjFolder) == false)
            {
                MessageBox.Show("No project Loaded");
                return;
            }

            if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
            {
                ProjFolder = ProjFolder + "\\";
            }

            string fisier_cl = ProjFolder + _AGEN_mainform.cl_excel_name;


            if (System.IO.File.Exists(fisier_cl) == false)
            {
                MessageBox.Show("No centerline file found");
                return;
            }

            string fisier_si = ProjFolder + _AGEN_mainform.sheet_index_excel_name;


            if (System.IO.File.Exists(fisier_si) == false)
            {
                MessageBox.Show("No sheet index file found");
                return;
            }


            _AGEN_mainform.tpage_processing.Show();

            set_enable_false();

            if (_AGEN_mainform.dt_centerline == null || _AGEN_mainform.dt_centerline.Rows.Count == 0 || _AGEN_mainform.tpage_setup.get_no_segments() > 1)
            {
                Load_centerline_and_station_equation(fisier_cl);
            }

            if (_AGEN_mainform.dt_centerline.Rows.Count == 0)
            {
                MessageBox.Show("No centerline data found");
                set_enable_true();
                _AGEN_mainform.tpage_processing.Hide();
                return;
            }

            if (_AGEN_mainform.dt_station_equation.Rows.Count == 0)
            {
                MessageBox.Show("No station equations data found");
                set_enable_true();
                _AGEN_mainform.tpage_processing.Hide();
                return;
            }

            if (_AGEN_mainform.dt_sheet_index == null)
            {
                _AGEN_mainform.dt_sheet_index = Load_existing_sheet_index(fisier_si);
            }



            if (_AGEN_mainform.dt_centerline != null)
            {
                if (_AGEN_mainform.dt_centerline.Rows.Count > 0)
                {
                    try
                    {
                        delete_entities_with_OD(_AGEN_mainform.layer_eq_blocks, "Agen_eq");
                        Functions.Creaza_layer(_AGEN_mainform.layer_eq_blocks, 2, true);

                        Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                        Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                        using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                        {

                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                            {
                                Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                                Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                                #region object data table
                                Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                                Functions.Create_eq_od_table();

                                List<object> Lista_val = new List<object>();
                                List<Autodesk.Gis.Map.Constants.DataType> Lista_type = new List<Autodesk.Gis.Map.Constants.DataType>();

                                Lista_val.Add(comboBox_segment_name.Text);
                                Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                Lista_val.Add(System.DateTime.Today.Year + "-" + System.DateTime.Today.Month + "-" + System.DateTime.Today.Day + " at " + System.DateTime.Now.Hour + ":" + System.DateTime.Now.Minute);
                                Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                                #endregion

                                if (_AGEN_mainform.dt_station_equation != null)
                                {
                                    if (_AGEN_mainform.dt_station_equation.Rows.Count > 0)
                                    {
                                        Polyline3d Poly3D = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);
                                        Polyline Poly2d = Functions.Build_2d_poly_for_scanning(_AGEN_mainform.dt_centerline);
                                        for (int i = 0; i < _AGEN_mainform.dt_station_equation.Rows.Count; ++i)
                                        {
                                            if (_AGEN_mainform.dt_station_equation.Rows[i]["Station Back"] != DBNull.Value &&
                                                _AGEN_mainform.dt_station_equation.Rows[i]["Station Ahead"] != DBNull.Value &&
                                                _AGEN_mainform.dt_station_equation.Rows[i]["Reroute End X"] != DBNull.Value &&
                                                _AGEN_mainform.dt_station_equation.Rows[i]["Reroute End Y"] != DBNull.Value)
                                            {

                                                bool display = false;
                                                if (_AGEN_mainform.dt_station_equation.Rows[i]["Show in plan"] == DBNull.Value)
                                                {
                                                    MessageBox.Show("please update the station equation column show in plan with yes or no");
                                                    set_enable_true();
                                                    _AGEN_mainform.tpage_processing.Hide();
                                                    return;
                                                }
                                                if (_AGEN_mainform.dt_station_equation.Rows[i]["Show in plan"] != DBNull.Value)
                                                {
                                                    if (Convert.ToString(_AGEN_mainform.dt_station_equation.Rows[i]["Show in plan"]).ToUpper() == "YES")
                                                    {
                                                        display = true;
                                                    }
                                                    else if (Convert.ToString(_AGEN_mainform.dt_station_equation.Rows[i]["Show in plan"]).ToUpper() == "NO")
                                                    {

                                                    }
                                                    else
                                                    {
                                                        MessageBox.Show("please update the station equation column [show in plan] with yes or no");
                                                        set_enable_true();
                                                        _AGEN_mainform.tpage_processing.Hide();
                                                        return;
                                                    }

                                                }

                                                if (display == true)
                                                {
                                                    double SB = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i]["Station Back"]);
                                                    double SA = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i]["Station Ahead"]);
                                                    double x1 = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i]["Reroute End X"]);
                                                    double y1 = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i]["Reroute End Y"]);

                                                    if (SB != 0)
                                                    {

                                                        double param1 = Poly2d.GetParameterAtPoint(Poly2d.GetClosestPointTo(new Point3d(x1, y1, Poly2d.Elevation), Vector3d.ZAxis, false));
                                                        double sta1 = Poly3D.GetDistanceAtParameter(param1);


                                                        System.Collections.Specialized.StringCollection col_atr = new System.Collections.Specialized.StringCollection();
                                                        System.Collections.Specialized.StringCollection col_val = new System.Collections.Specialized.StringCollection();
                                                        col_atr.Add(comboBox_eq_bs.Text);
                                                        col_val.Add(Functions.Get_chainage_from_double(SB, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1));
                                                        col_atr.Add(comboBox_eq_as.Text);
                                                        col_val.Add(Functions.Get_chainage_from_double(SA, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1));
                                                        col_atr.Add(comboBox_eq_diff.Text);
                                                        string suff1 = "'";
                                                        if (_AGEN_mainform.units_of_measurement == "m") suff1 = "m";
                                                        col_val.Add(Functions.Get_String_Rounded(Math.Round(SB, _AGEN_mainform.round1) - Math.Round(SA, _AGEN_mainform.round1), _AGEN_mainform.round1) + suff1);

                                                        string Col_M1 = "StaBeg";
                                                        string Col_M2 = "StaEnd";
                                                        string Col_rot = "Rotation";
                                                        double Rot1 = 0;

                                                        if (_AGEN_mainform.dt_sheet_index.Rows.Count > 0)
                                                        {
                                                            for (int k = 0; k < _AGEN_mainform.dt_sheet_index.Rows.Count; ++k)
                                                            {
                                                                if (_AGEN_mainform.dt_sheet_index.Rows[k][Col_M1] != DBNull.Value &&
                                                                    _AGEN_mainform.dt_sheet_index.Rows[k][Col_M2] != DBNull.Value &&
                                                                    _AGEN_mainform.dt_sheet_index.Rows[k][Col_rot] != DBNull.Value)
                                                                {
                                                                    double M1 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[k][Col_M1]);
                                                                    double M2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[k][Col_M2]);
                                                                    double rr = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[k][Col_rot]);
                                                                    if (sta1 >= M1 && sta1 <= M2)
                                                                    {
                                                                        Rot1 = rr * Math.PI / 180;
                                                                        k = _AGEN_mainform.dt_sheet_index.Rows.Count;
                                                                    }


                                                                }
                                                            }
                                                        }

                                                        double scale1 = 1 / _AGEN_mainform.Vw_scale;
                                                        if (Functions.IsNumeric(textBox_scale_eq.Text) == true)
                                                        {
                                                            scale1 = Convert.ToDouble(textBox_scale_eq.Text);
                                                        }

                                                        Point3d pt_ins = new Point3d(x1, y1, 0);
                                                        BlockReference Block1 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "",
                                                                comboBox_eq_block.Text, pt_ins, scale1, Rot1, _AGEN_mainform.layer_eq_blocks, col_atr, col_val);

                                                        Functions.Populate_object_data_table_from_objectid(Tables1, Block1.ObjectId, "Agen_eq", Lista_val, Lista_type);

                                                    }
                                                }



                                            }
                                        }

                                        Poly3D.Erase();

                                    }
                                }


                                Trans1.Commit();
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                _AGEN_mainform.tpage_processing.Hide();
                set_enable_true();
            }




        }

        private void button_refresh_pi_blocks_Click(object sender, EventArgs e)
        {
            Functions.Incarca_existing_Blocks_to_combobox(comboBox_pi_block);
        }

        private void comboBox_pi_block_SelectedIndexChanged(object sender, EventArgs e)
        {
            Functions.Incarca_existing_Atributes_to_combobox(comboBox_pi_block.Text, comboBox_pi_atr_sta);
            Functions.Incarca_existing_Atributes_to_combobox(comboBox_pi_block.Text, comboBox_pi_atr_defl);
        }

        private void button_insert_pi_blocks_Click(object sender, EventArgs e)
        {
            Ag = this.MdiParent as _AGEN_mainform;
            Functions.Kill_excel();


            if (Functions.Get_if_workbook_is_open_in_Excel("centerline.xlsx") == true)
            {
                MessageBox.Show("Please close the centerline file");
                return;
            }


            string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();

            if (System.IO.Directory.Exists(ProjFolder) == false)
            {
                MessageBox.Show("No project Loaded");
                return;
            }

            if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
            {
                ProjFolder = ProjFolder + "\\";
            }

            string fisier_cl = ProjFolder + _AGEN_mainform.cl_excel_name;


            if (System.IO.File.Exists(fisier_cl) == false)
            {
                MessageBox.Show("No centerline file found");
                return;
            }

            string fisier_si = ProjFolder + _AGEN_mainform.sheet_index_excel_name;

            _AGEN_mainform.tpage_processing.Show();

            set_enable_false();

            if (_AGEN_mainform.dt_centerline == null || _AGEN_mainform.dt_centerline.Rows.Count == 0 || _AGEN_mainform.tpage_setup.get_no_segments() > 1)
            {
                Load_centerline_and_station_equation(fisier_cl);
            }

            if (_AGEN_mainform.dt_centerline.Rows.Count == 0)
            {
                MessageBox.Show("No centerline data found");
                set_enable_true();
                _AGEN_mainform.tpage_processing.Hide();
                return;
            }


            if (_AGEN_mainform.dt_sheet_index == null & System.IO.File.Exists(fisier_si) == true)
            {
                _AGEN_mainform.dt_sheet_index = Load_existing_sheet_index(fisier_si);
            }

            if (_AGEN_mainform.dt_centerline != null)
            {
                if (_AGEN_mainform.dt_centerline.Rows.Count > 0)
                {
                    try
                    {
                        delete_entities_with_OD(_AGEN_mainform.layer_pi_blocks, "Agen_pi");
                        Functions.Creaza_layer(_AGEN_mainform.layer_pi_blocks, 2, true);

                        Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                        Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                        using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                        {

                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                            {
                                Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                                Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                                #region object data table
                                Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                                Functions.Create_pi_od_table();

                                List<object> Lista_val = new List<object>();
                                List<Autodesk.Gis.Map.Constants.DataType> Lista_type = new List<Autodesk.Gis.Map.Constants.DataType>();

                                Lista_val.Add(comboBox_segment_name.Text);
                                Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                Lista_val.Add(System.DateTime.Today.Year + "-" + System.DateTime.Today.Month + "-" + System.DateTime.Today.Day + " at " + System.DateTime.Now.Hour + ":" + System.DateTime.Now.Minute);
                                Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                #endregion

                                Polyline3d Poly3D = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);
                                Polyline Poly2D = Functions.Build_2dpoly_from_3d(Poly3D);

                                string Col_x = "X";
                                string Col_y = "Y";
                                string Col_DeflAng = "DeflAng";
                                string Col_DeflAngDMS = "DeflAngDMS";

                                for (int i = 0; i < _AGEN_mainform.dt_centerline.Rows.Count; ++i)
                                {
                                    if (_AGEN_mainform.dt_centerline.Rows[i][Col_x] != DBNull.Value && _AGEN_mainform.dt_centerline.Rows[i][Col_y] != DBNull.Value)
                                    {
                                        double x = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i][Col_x]);
                                        double y = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i][Col_y]);

                                        Point3d pt1 = new Point3d(x, y, 0);

                                        System.Collections.Specialized.StringCollection col_atr = new System.Collections.Specialized.StringCollection();
                                        System.Collections.Specialized.StringCollection col_val = new System.Collections.Specialized.StringCollection();

                                        string statie1 = "";
                                        string deflectie1 = "";

                                        double param1 = Poly2D.GetParameterAtPoint(Poly2D.GetClosestPointTo(new Point3d(x, y, 0), Vector3d.ZAxis, false));
                                        double sta_meas = Poly3D.GetDistanceAtParameter(param1);

                                        double min_pi = -1;
                                        double Defl_pi = 0;

                                        if (_AGEN_mainform.dt_centerline.Rows[i][Col_DeflAng] != DBNull.Value)
                                        {
                                            string defl_str = Convert.ToString(_AGEN_mainform.dt_centerline.Rows[i][Col_DeflAng]);
                                            if (Functions.IsNumeric(defl_str) == true)
                                            {
                                                if (Functions.IsNumeric(textBox_pi_min_angle.Text) == true)
                                                {
                                                    Defl_pi = Convert.ToDouble(defl_str);
                                                    min_pi = Convert.ToDouble(textBox_pi_min_angle.Text);
                                                }
                                            }
                                        }

                                        if (Defl_pi >= min_pi)
                                        {
                                            if (comboBox_pi_atr_sta.Text != "" &&
                                            (checkBox_pi_show_stations.Checked == true | checkBox_pi_show_deflection.Checked == true))
                                            {
                                                if (checkBox_pi_show_stations.Checked == true)
                                                {
                                                    double sta_displayed = Functions.Station_equation_of(sta_meas, _AGEN_mainform.dt_station_equation);
                                                    statie1 = Functions.Get_chainage_from_double(sta_displayed, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);
                                                }

                                                if (checkBox_pi_show_deflection.Checked == true)
                                                {
                                                    if (_AGEN_mainform.dt_centerline.Rows[i][Col_DeflAngDMS] != DBNull.Value && _AGEN_mainform.dt_centerline.Rows[i][Col_DeflAng] != DBNull.Value)
                                                    {
                                                        string defldms = Convert.ToString(_AGEN_mainform.dt_centerline.Rows[i][Col_DeflAngDMS]);
                                                        string defl_str = Convert.ToString(_AGEN_mainform.dt_centerline.Rows[i][Col_DeflAng]);
                                                        if (Functions.IsNumeric(defl_str) == true)
                                                        {
                                                            if (Functions.IsNumeric(textBox_pi_min_angle.Text) == true)
                                                            {
                                                                double defl = Convert.ToDouble(defl_str);
                                                                double defl1 = Convert.ToDouble(textBox_pi_min_angle.Text);
                                                                if (defl >= defl1)
                                                                {
                                                                    deflectie1 = defldms;
                                                                }
                                                            }
                                                        }
                                                    }
                                                }

                                                if (statie1 != "")
                                                {
                                                    col_atr.Add(comboBox_pi_atr_sta.Text);
                                                    col_val.Add(statie1);
                                                }
                                                if (deflectie1 != "")
                                                {
                                                    col_atr.Add(comboBox_pi_atr_defl.Text);
                                                    col_val.Add(deflectie1);
                                                }
                                            }

                                            double Rot1 = 0;
                                            if (_AGEN_mainform.dt_sheet_index != null)
                                            {
                                                if (_AGEN_mainform.dt_sheet_index.Rows.Count > 0 && (statie1 != "" || deflectie1 != ""))
                                                {
                                                    string Col_M1 = "StaBeg";
                                                    string Col_M2 = "StaEnd";
                                                    string Col_rot = "Rotation";

                                                    for (int k = 0; k < _AGEN_mainform.dt_sheet_index.Rows.Count; ++k)
                                                    {
                                                        if (_AGEN_mainform.dt_sheet_index.Rows[k][Col_M1] != DBNull.Value &&
                                                            _AGEN_mainform.dt_sheet_index.Rows[k][Col_M2] != DBNull.Value &&
                                                            _AGEN_mainform.dt_sheet_index.Rows[k][Col_rot] != DBNull.Value)
                                                        {
                                                            double M1 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[k][Col_M1]);
                                                            double M2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[k][Col_M2]);
                                                            double rr = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[k][Col_rot]);
                                                            if (sta_meas >= M1 && sta_meas <= M2)
                                                            {
                                                                Rot1 = rr * Math.PI / 180;
                                                                k = _AGEN_mainform.dt_sheet_index.Rows.Count;
                                                            }
                                                        }
                                                    }
                                                }
                                            }

                                            double scale1 = 1 / _AGEN_mainform.Vw_scale;
                                            if (Functions.IsNumeric(textBox_scale_pi.Text) == true)
                                            {
                                                scale1 = Convert.ToDouble(textBox_scale_pi.Text);
                                            }

                                            Point3d pt_ins = new Point3d(pt1.X, pt1.Y, 0);
                                            BlockReference block1 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "",
                                                    comboBox_pi_block.Text, pt_ins, scale1, Rot1, _AGEN_mainform.layer_pi_blocks, col_atr, col_val);

                                            Functions.Populate_object_data_table_from_objectid(Tables1, block1.ObjectId, "Agen_pi", Lista_val, Lista_type);
                                        }
                                    }
                                }
                                Poly3D.Erase();
                                Trans1.Commit();
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                _AGEN_mainform.tpage_processing.Hide();
                set_enable_true();
            }
        }


        private void button_mpkp_refresh_blocks_Click(object sender, EventArgs e)
        {
            Functions.Incarca_existing_Blocks_with_attributes_to_combobox(comboBox_mpkp_blocks);
        }

        private void comboBox_mpkp_blocks_SelectedIndexChanged(object sender, EventArgs e)
        {
            Functions.Incarca_existing_Atributes_to_combobox(comboBox_mpkp_blocks.Text, comboBox_mpkp_attribute);
        }

        private void button_mpkp_insert_Click(object sender, EventArgs e)
        {
            Ag = this.MdiParent as _AGEN_mainform;

            if (Functions.IsNumeric(textBox_kpmp_spacing.Text) == false)
            {
                MessageBox.Show("no spacing specified");
                return;
            }


            double spacing = Math.Abs(Convert.ToDouble(textBox_kpmp_spacing.Text));

            if (_AGEN_mainform.units_of_measurement == "m")
            {
                spacing = spacing * 1000;
            }
            else
            {
                spacing = spacing * 5280;
            }


            if (spacing <= 0)
            {
                MessageBox.Show("spacing is negative");
                return;
            }


            Functions.Kill_excel();

            if (Functions.Get_if_workbook_is_open_in_Excel("centerline.xlsx") == true)
            {
                MessageBox.Show("Please close the centerline file");
                return;
            }

            string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();

            if (System.IO.Directory.Exists(ProjFolder) == false)
            {
                MessageBox.Show("No project Loaded");
                return;
            }

            if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
            {
                ProjFolder = ProjFolder + "\\";
            }

            string fisier_cl = ProjFolder + _AGEN_mainform.cl_excel_name;


            if (System.IO.File.Exists(fisier_cl) == false)
            {
                MessageBox.Show("No centerline file found");
                return;
            }

            _AGEN_mainform.tpage_processing.Show();

            double start1 = 0;

            set_enable_false();

            Load_centerline_and_station_equation(fisier_cl);

            if (_AGEN_mainform.dt_centerline != null)
            {
                if (_AGEN_mainform.dt_centerline.Rows.Count > 0)
                {
                    try
                    {


                        delete_entities_with_OD(_AGEN_mainform.layer_mp_blocks, "Agen_mp_block");

                        Functions.Creaza_layer(_AGEN_mainform.layer_mp_blocks, 2, true);


                        Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                        Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                        using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                        {

                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                            {
                                Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                                Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                                #region object data stationing

                                Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                                Functions.Create_kpmp_od_table();
                                List<object> Lista_val = new List<object>();
                                List<Autodesk.Gis.Map.Constants.DataType> Lista_type = new List<Autodesk.Gis.Map.Constants.DataType>();
                                Lista_val.Add(comboBox_segment_name.Text);
                                Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                                Lista_val.Add(System.DateTime.Today.Year + "-" + System.DateTime.Today.Month + "-" + System.DateTime.Today.Day + " at " + System.DateTime.Now.Hour + ":" + System.DateTime.Now.Minute);
                                Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                #endregion



                                string suf_at_start = "";
                                string suf_at_end = "";

                                Polyline3d Poly3D = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);
                                Polyline Poly2D = Functions.Build_2d_poly_for_scanning(_AGEN_mainform.dt_centerline);

                                #region USA     
                                if (_AGEN_mainform.COUNTRY == "USA")
                                {


                                    bool is_equated = false;
                                    bool is_last_inserted = false;
                                    double sta_end = Poly3D.Length;






                                    if (_AGEN_mainform.dt_station_equation != null)
                                    {
                                        if (_AGEN_mainform.dt_station_equation.Rows.Count > 0)
                                        {
                                            double param1p = 0;
                                            double SAp = 0;

                                            for (int i = 0; i < _AGEN_mainform.dt_station_equation.Rows.Count; ++i)
                                            {
                                                if (_AGEN_mainform.dt_station_equation.Rows[i]["Station Back"] != DBNull.Value &&
                                                    _AGEN_mainform.dt_station_equation.Rows[i]["Station Ahead"] != DBNull.Value &&
                                                    _AGEN_mainform.dt_station_equation.Rows[i]["Reroute End X"] != DBNull.Value &&
                                                    _AGEN_mainform.dt_station_equation.Rows[i]["Reroute End Y"] != DBNull.Value &&
                                                    _AGEN_mainform.dt_station_equation.Rows[i]["Reroute End Z"] != DBNull.Value &&
                                                     _AGEN_mainform.dt_station_equation.Rows[i]["Reroute Start X"] != DBNull.Value &&
                                                    _AGEN_mainform.dt_station_equation.Rows[i]["Reroute Start Y"] != DBNull.Value &&
                                                    _AGEN_mainform.dt_station_equation.Rows[i]["Reroute Start Z"] != DBNull.Value)
                                                {

                                                    string Suffix1 = "";
                                                    if (checkBox_mpkp_eq_suffix.Checked == true)
                                                    {
                                                        if (_AGEN_mainform.dt_station_equation.Rows[i]["Version"] != DBNull.Value)
                                                        {
                                                            Suffix1 = Convert.ToString(_AGEN_mainform.dt_station_equation.Rows[i]["Version"]);
                                                        }
                                                    }

                                                    double SB = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i]["Station Back"]);
                                                    double SA = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i]["Station Ahead"]);

                                                    Point3d pt_start = new Point3d(Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i]["Reroute Start X"]),
                                                                                    Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i]["Reroute Start Y"]),
                                                                                        Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i]["Reroute Start Z"]));

                                                    Point3d pt_end = new Point3d(Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i]["Reroute End X"]),
                                                        Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i]["Reroute End Y"]),
                                                        Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i]["Reroute End Z"]));


                                                    double param1 = Poly2D.GetParameterAtPoint(Poly2D.GetClosestPointTo(new Point3d(pt_end.X, pt_end.Y, Poly2D.Elevation), Vector3d.ZAxis, false));
                                                    if (SB != 0)
                                                    {
                                                        is_equated = true;



                                                        Polyline3d Poly3D_eq = new Polyline3d();
                                                        BTrecord.AppendEntity(Poly3D_eq);
                                                        Trans1.AddNewlyCreatedDBObject(Poly3D_eq, true);

                                                        if (Math.Ceiling(param1p) != param1p)
                                                        {
                                                            PolylineVertex3d Vertex_last = new PolylineVertex3d(new Point3d(Poly3D.GetPointAtParameter(param1p).X, Poly3D.GetPointAtParameter(param1p).Y, Poly3D.GetPointAtParameter(param1p).Z));
                                                            Poly3D_eq.AppendVertex(Vertex_last);
                                                            Trans1.AddNewlyCreatedDBObject(Vertex_last, true);
                                                        }

                                                        for (int k = Convert.ToInt32(Math.Ceiling(param1p)); k <= Math.Floor(param1); ++k)
                                                        {
                                                            PolylineVertex3d Vertex_eq = new PolylineVertex3d(new Point3d(Poly3D.GetPointAtParameter(k).X, Poly3D.GetPointAtParameter(k).Y, Poly3D.GetPointAtParameter(k).Z));
                                                            Poly3D_eq.AppendVertex(Vertex_eq);
                                                            Trans1.AddNewlyCreatedDBObject(Vertex_eq, true);

                                                        }
                                                        if (Math.Floor(param1) != param1)
                                                        {
                                                            PolylineVertex3d Vertex_last = new PolylineVertex3d(new Point3d(Poly3D.GetPointAtParameter(param1).X, Poly3D.GetPointAtParameter(param1).Y, Poly3D.GetPointAtParameter(param1).Z));
                                                            Poly3D_eq.AppendVertex(Vertex_last);
                                                            Trans1.AddNewlyCreatedDBObject(Vertex_last, true);
                                                        }

                                                        Polyline p2d = Functions.Build_2dpoly_from_3d(Poly3D_eq);
                                                        double paramA = p2d.GetParameterAtPoint(p2d.GetClosestPointTo(new Point3d(pt_start.X, pt_start.Y, p2d.Elevation), Vector3d.ZAxis, false));
                                                        double paramB = p2d.GetParameterAtPoint(p2d.GetClosestPointTo(new Point3d(pt_end.X, pt_end.Y, p2d.Elevation), Vector3d.ZAxis, false));

                                                        Insert_kpmp_blocks(Poly3D_eq, SAp, spacing, Suffix1, paramA, paramB, Tables1, Lista_val, Lista_type);

                                                        if (i == _AGEN_mainform.dt_station_equation.Rows.Count - 1)
                                                        {
                                                            double dist1 = Poly3D.GetDistanceAtParameter(param1);
                                                            double len1 = Poly3D.Length;
                                                            sta_end = SA + len1 - dist1;

                                                            if (Math.Abs(Math.Round(dist1 - len1, 0)) <= 1)
                                                            {
                                                                suf_at_end = Suffix1;

                                                                BlockReference Block2 = Insert_kpmp_block_at_end(Poly3D, sta_end, suf_at_end);
                                                                Functions.Populate_object_data_table_from_objectid(Tables1, Block2.ObjectId, "Agen_mp_block", Lista_val, Lista_type);
                                                                is_last_inserted = true;
                                                            }
                                                        }


                                                        param1p = param1;
                                                    }
                                                    else
                                                    {
                                                        suf_at_start = Suffix1;
                                                    }
                                                    SAp = SA;



                                                }
                                            }

                                            if (param1p < Poly3D.EndParam)
                                            {

                                                double param1 = Poly3D.EndParam;


                                                Polyline3d Poly3D_eq = new Polyline3d();
                                                BTrecord.AppendEntity(Poly3D_eq);
                                                Trans1.AddNewlyCreatedDBObject(Poly3D_eq, true);

                                                if (Math.Ceiling(param1p) != param1p)
                                                {
                                                    PolylineVertex3d Vertex_last = new PolylineVertex3d(new Point3d(Poly3D.GetPointAtParameter(param1p).X, Poly3D.GetPointAtParameter(param1p).Y, Poly3D.GetPointAtParameter(param1p).Z));
                                                    Poly3D_eq.AppendVertex(Vertex_last);
                                                    Trans1.AddNewlyCreatedDBObject(Vertex_last, true);
                                                }

                                                for (int k = Convert.ToInt32(Math.Ceiling(param1p)); k <= param1; ++k)
                                                {
                                                    PolylineVertex3d Vertex_eq = new PolylineVertex3d(new Point3d(Poly3D.GetPointAtParameter(k).X, Poly3D.GetPointAtParameter(k).Y, Poly3D.GetPointAtParameter(k).Z));
                                                    Poly3D_eq.AppendVertex(Vertex_eq);
                                                    Trans1.AddNewlyCreatedDBObject(Vertex_eq, true);

                                                }


                                                Insert_kpmp_blocks(Poly3D_eq, SAp, spacing, "", 0, param1, Tables1, Lista_val, Lista_type);

                                                if (is_last_inserted == false)
                                                {
                                                    BlockReference Block2 = Insert_kpmp_block_at_end(Poly3D, sta_end, suf_at_end);
                                                    Functions.Populate_object_data_table_from_objectid(Tables1, Block2.ObjectId, "Agen_mp_block", Lista_val, Lista_type);
                                                }

                                            }


                                        }
                                    }


                                    if (is_equated == false)
                                    {
                                        bool is_2d = true;
                                        if (_AGEN_mainform.Project_type == "3D") is_2d = false;
                                        Insert_kpmp_blocks_with_curves(Poly3D, Poly2D, is_2d, start1, spacing, "", 0, Poly3D.EndParam, Tables1, Lista_val, Lista_type);
                                        BlockReference Block2 = Insert_kpmp_block_at_end(Poly3D, Poly3D.Length, suf_at_end);
                                        Functions.Populate_object_data_table_from_objectid(Tables1, Block2.ObjectId, "Agen_mp_block", Lista_val, Lista_type);
                                    }

                                    BlockReference Block1 = Insert_kpmp_block_at_start(Poly3D, start1, suf_at_start);
                                    Functions.Populate_object_data_table_from_objectid(Tables1, Block1.ObjectId, "Agen_mp_block", Lista_val, Lista_type);



                                    Poly3D.Erase();
                                }
                                #endregion

                                #region CANADA
                                if (_AGEN_mainform.COUNTRY == "CANADA")
                                {
                                    if (_AGEN_mainform.dt_centerline != null)
                                    {
                                        if (_AGEN_mainform.dt_centerline.Rows.Count > 1)
                                        {

                                            string Col_x = "X";
                                            string Col_y = "Y";
                                            string Col_3DSta = "3DSta";

                                            bool ins_at_0 = true;

                                            sort_canadian_station_eq(Poly2D);

                                            for (int i = 1; i < _AGEN_mainform.dt_centerline.Rows.Count; ++i)
                                            {
                                                if (
                                                        _AGEN_mainform.dt_centerline.Rows[i - 1][Col_x] != DBNull.Value &&
                                                        _AGEN_mainform.dt_centerline.Rows[i - 1][Col_y] != DBNull.Value &&
                                                        _AGEN_mainform.dt_centerline.Rows[i - 1][Col_3DSta] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[i - 1][Col_x])) == true &&
                                                        Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[i - 1][Col_y])) == true &&
                                                        Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[i - 1][Col_3DSta])) == true &&
                                                        _AGEN_mainform.dt_centerline.Rows[i][Col_x] != DBNull.Value &&
                                                        _AGEN_mainform.dt_centerline.Rows[i][Col_y] != DBNull.Value &&
                                                        _AGEN_mainform.dt_centerline.Rows[i][Col_3DSta] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[i][Col_x])) == true &&
                                                        Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[i][Col_y])) == true &&
                                                        (
                                                            Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[i][Col_3DSta])) == true
                                                        )
                                                    )
                                                {


                                                    double x0 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i - 1][Col_x]);
                                                    double y0 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i - 1][Col_y]);
                                                    double sta0 = -1;
                                                    if (_AGEN_mainform.dt_centerline.Rows[i - 1][Col_3DSta] != DBNull.Value &&
                                                       Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[i - 1][Col_3DSta])) == true)
                                                    {
                                                        sta0 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i - 1][Col_3DSta]);
                                                        if (sta0 == 0 && i == 1) ins_at_0 = false;
                                                    }


                                                    double x1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i][Col_x]);
                                                    double y1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i][Col_y]);
                                                    double sta1 = -1;

                                                    if (_AGEN_mainform.dt_centerline.Rows[i][Col_3DSta] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[i][Col_3DSta])) == true)
                                                    {
                                                        sta1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i][Col_3DSta]);
                                                    }


                                                    Point3d pt0 = new Point3d(x0, y0, 0);
                                                    Point3d pt1 = new Point3d(x1, y1, 0);


                                                    Insert_kp_blocks_with_csf(sta0, sta1, pt0, pt1, spacing, Tables1, Lista_val, Lista_type);


                                                }
                                            }

                                            if (ins_at_0 == true &&
                                                    _AGEN_mainform.dt_centerline.Rows[0][Col_x] != DBNull.Value && _AGEN_mainform.dt_centerline.Rows[0][Col_y] != DBNull.Value &&
                                                    _AGEN_mainform.dt_centerline.Rows[0][Col_3DSta] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[0][Col_x])) == true && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[0][Col_y])) == true &&
                                                    Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[0][Col_3DSta])) == true &&
                                                    _AGEN_mainform.dt_centerline.Rows[1][Col_x] != DBNull.Value && _AGEN_mainform.dt_centerline.Rows[1][Col_y] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[1][Col_x])) == true && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[1][Col_y])) == true
                                                )
                                            {


                                                double x0 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[0][Col_x]);
                                                double y0 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[0][Col_y]);
                                                double sta0 = -1;
                                                if (_AGEN_mainform.dt_centerline.Rows[0][Col_3DSta] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[0][Col_3DSta])) == true)
                                                {
                                                    sta0 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[0][Col_3DSta]);
                                                }

                                                double x1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[1][Col_x]);
                                                double y1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[1][Col_y]);

                                                double bear1 = Functions.GET_Bearing_rad(x0, y0, x1, y1);

                                                if (_AGEN_mainform.Left_to_Right == false) bear1 = bear1 + Math.PI;
                                                System.Collections.Specialized.StringCollection col_atr = new System.Collections.Specialized.StringCollection();
                                                System.Collections.Specialized.StringCollection col_val = new System.Collections.Specialized.StringCollection();

                                                col_atr.Add(comboBox_mpkp_attribute.Text);

                                                int round1 = 0;

                                                if (comboBox_kpmp_units_precision.Text == "0.0")
                                                {
                                                    round1 = 1;
                                                }
                                                else if (comboBox_kpmp_units_precision.Text == "0.00")
                                                {
                                                    round1 = 2;
                                                }
                                                else if (comboBox_kpmp_units_precision.Text == "0.000")
                                                {
                                                    round1 = 3;
                                                }


                                                string val1 = Functions.Get_String_Rounded(sta0 / 1000, round1);

                                                col_val.Add(val1);

                                                double scale1 = 1 / _AGEN_mainform.Vw_scale;
                                                if (Functions.IsNumeric(textBox_scale_mp.Text) == true)
                                                {
                                                    scale1 = Convert.ToDouble(textBox_scale_mp.Text);
                                                }

                                                BlockReference Block1 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "",
                                                      comboBox_mpkp_blocks.Text, new Point3d(x0, y0, 0), scale1, bear1, _AGEN_mainform.layer_mp_blocks, col_atr, col_val);

                                                Functions.Populate_object_data_table_from_objectid(Tables1, Block1.ObjectId, "Agen_mp_block", Lista_val, Lista_type);

                                            }

                                            int last_idx = _AGEN_mainform.dt_centerline.Rows.Count - 1;

                                            if (_AGEN_mainform.dt_centerline.Rows[last_idx][Col_x] != DBNull.Value && _AGEN_mainform.dt_centerline.Rows[last_idx][Col_y] != DBNull.Value &&
                                                  _AGEN_mainform.dt_centerline.Rows[last_idx][Col_3DSta] != DBNull.Value &&
                                                  Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[last_idx][Col_x])) == true && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[last_idx][Col_y])) == true &&
                                                  Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[last_idx][Col_3DSta])) == true &&
                                                    _AGEN_mainform.dt_centerline.Rows[last_idx - 1][Col_x] != DBNull.Value && _AGEN_mainform.dt_centerline.Rows[last_idx - 1][Col_y] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[last_idx - 1][Col_x])) == true && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[last_idx - 1][Col_y])) == true
                                                )
                                            {


                                                double x0 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[last_idx][Col_x]);
                                                double y0 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[last_idx][Col_y]);

                                                double x1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[last_idx - 1][Col_x]);
                                                double y1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[last_idx - 1][Col_y]);

                                                double sta0 = -1;
                                                if (_AGEN_mainform.dt_centerline.Rows[last_idx][Col_3DSta] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[last_idx][Col_3DSta])) == true)
                                                {
                                                    sta0 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[last_idx][Col_3DSta]);
                                                }



                                                double bear1 = Functions.GET_Bearing_rad(x1, y1, x0, y0);

                                                if (_AGEN_mainform.Left_to_Right == false) bear1 = bear1 + Math.PI;
                                                System.Collections.Specialized.StringCollection col_atr = new System.Collections.Specialized.StringCollection();
                                                System.Collections.Specialized.StringCollection col_val = new System.Collections.Specialized.StringCollection();

                                                col_atr.Add(comboBox_mpkp_attribute.Text);

                                                int round1 = 1;

                                                if (comboBox_kpmp_units_precision.Text == "0.0")
                                                {
                                                    round1 = 2;
                                                }
                                                else if (comboBox_kpmp_units_precision.Text == "0.00")
                                                {
                                                    round1 = 3;
                                                }
                                                else if (comboBox_kpmp_units_precision.Text == "0.000")
                                                {
                                                    round1 = 4;
                                                }


                                                string val1 = Functions.Get_String_Rounded(sta0 / 1000, round1);

                                                col_val.Add(val1);

                                                double scale1 = 1 / _AGEN_mainform.Vw_scale;
                                                if (Functions.IsNumeric(textBox_scale_mp.Text) == true)
                                                {
                                                    scale1 = Convert.ToDouble(textBox_scale_mp.Text);
                                                }

                                                BlockReference Block1 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "",
                                                      comboBox_mpkp_blocks.Text, new Point3d(x0, y0, 0), scale1, bear1, _AGEN_mainform.layer_mp_blocks, col_atr, col_val);

                                                Functions.Populate_object_data_table_from_objectid(Tables1, Block1.ObjectId, "Agen_mp_block", Lista_val, Lista_type);

                                            }


                                        }
                                    }
                                }
                                #endregion



                                Trans1.Commit();
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                _AGEN_mainform.tpage_processing.Hide();
                set_enable_true();
            }

        }

        private void Insert_kpmp_blocks(Polyline3d Poly3D, double start1, double spacing1, string Suffix1, double rrstart_param, double rrend_param,
                                                                                    Autodesk.Gis.Map.ObjectData.Tables Tables1, List<object> Lista_val, List<Autodesk.Gis.Map.Constants.DataType> Lista_type)
        {
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);


                    double first_lbl = Math.Floor((start1 + spacing1) / spacing1) * spacing1;

                    if (start1 + Poly3D.Length >= first_lbl)
                    {
                        int no_blocks = Convert.ToInt32(Math.Ceiling((start1 + Poly3D.Length - first_lbl) / spacing1));

                        if (no_blocks > 0)
                        {

                            for (int i = 0; i < no_blocks; ++i)
                            {
                                double sta_meas = first_lbl - start1 + i * spacing1;
                                Point3d pt_ins = Poly3D.GetPointAtDist(sta_meas);
                                double param1 = Poly3D.GetParameterAtDistance(sta_meas);
                                string new_suffix = "";

                                if (param1 <= rrend_param && param1 >= rrstart_param)
                                {
                                    new_suffix = Suffix1;
                                }


                                double param2 = param1 + 1;
                                if (Poly3D.EndParam < param2)
                                {
                                    param1 = Poly3D.EndParam - 1;
                                    param2 = Poly3D.EndParam;
                                }



                                Point3d point1 = Poly3D.GetPointAtParameter(Math.Floor(param1));

                                Point3d point2 = Poly3D.GetPointAtParameter(Math.Floor(param2));

                                double bear1 = Functions.GET_Bearing_rad(point1.X, point1.Y, point2.X, point2.Y);
                                double position_measured = first_lbl + i * spacing1;
                                System.Collections.Specialized.StringCollection col_atr = new System.Collections.Specialized.StringCollection();
                                System.Collections.Specialized.StringCollection col_val = new System.Collections.Specialized.StringCollection();

                                col_atr.Add(comboBox_mpkp_attribute.Text);

                                int round1 = 0;

                                if (comboBox_kpmp_units_precision.Text == "0.0")
                                {
                                    round1 = 1;
                                }
                                else if (comboBox_kpmp_units_precision.Text == "0.00")
                                {
                                    round1 = 2;
                                }
                                else if (comboBox_kpmp_units_precision.Text == "0.000")
                                {
                                    round1 = 3;
                                }

                                double fact1 = 5280;
                                if (_AGEN_mainform.units_of_measurement == "m") fact1 = 1000;

                                string val1 = Functions.Get_String_Rounded((position_measured / fact1), round1) + new_suffix;

                                col_val.Add(val1);

                                double scale1 = 1 / _AGEN_mainform.Vw_scale;
                                if (Functions.IsNumeric(textBox_scale_mp.Text) == true)
                                {
                                    scale1 = Convert.ToDouble(textBox_scale_mp.Text);
                                }

                                BlockReference Block1 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "",
                                      comboBox_mpkp_blocks.Text, pt_ins, scale1, bear1, _AGEN_mainform.layer_mp_blocks, col_atr, col_val);
                                Functions.Populate_object_data_table_from_objectid(Tables1, Block1.ObjectId, "Agen_mp_block", Lista_val, Lista_type);


                            }
                        }
                    }
                    Poly3D.Erase();
                    Trans1.Commit();

                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        private void Insert_kpmp_blocks_with_curves(Polyline3d Poly3D, Polyline poly1, bool is_2d, double start1, double spacing1, string Suffix1, double rrstart_param, double rrend_param,
                                                                            Autodesk.Gis.Map.ObjectData.Tables Tables1, List<object> Lista_val, List<Autodesk.Gis.Map.Constants.DataType> Lista_type)
        {
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);


                    double first_lbl = Math.Floor((start1 + spacing1) / spacing1) * spacing1;

                    if (start1 + Poly3D.Length >= first_lbl)
                    {
                        int no_blocks = Convert.ToInt32(Math.Ceiling((start1 + Poly3D.Length - first_lbl) / spacing1));

                        if (no_blocks > 0)
                        {

                            for (int i = 0; i < no_blocks; ++i)
                            {
                                double sta_meas = first_lbl - start1 + i * spacing1;


                                Point3d pt_ins = new Point3d();
                                if (is_2d == false)
                                {
                                    pt_ins = Poly3D.GetPointAtDist(sta_meas);
                                }
                                else
                                {
                                    pt_ins = poly1.GetPointAtDist(sta_meas);
                                }

                                double paramx = 0;
                                if (is_2d == true)
                                {
                                    paramx = poly1.GetParameterAtDistance(sta_meas);
                                }
                                else
                                {
                                    paramx = Poly3D.GetParameterAtDistance(sta_meas);
                                }

                                int param1 = Convert.ToInt32(Math.Floor(paramx));

                                string new_suffix = "";

                                if (paramx <= rrend_param && paramx >= rrstart_param)
                                {
                                    new_suffix = Suffix1;
                                }


                                int param2 = param1 + 1;
                                if (poly1.EndParam < param2)
                                {
                                    param1 = poly1.NumberOfVertices - 2;
                                    param2 = param1 + 1;
                                }



                                double bulge1 = poly1.GetBulgeAt(param1);


                                Point3d point1 = poly1.GetPointAtParameter(param1);

                                Point3d point2 = poly1.GetPointAtParameter(param2);
                                double extra1 = 0;

                                if (bulge1 != 0)
                                {
                                    CircularArc3d arc1 = poly1.GetArcSegmentAt(param1);
                                    point1 = arc1.Center;
                                    point2 = pt_ins;
                                    extra1 = -Math.PI / 2;

                                    //Polyline poly_del = new Polyline();
                                    //poly_del.AddVertexAt(0, new Point2d(point1.X, point1.Y), 0, 0, 0);
                                    //poly_del.AddVertexAt(1, new Point2d(point2.X, point2.Y), 0, 0, 0);
                                    //poly_del.ColorIndex = 4;
                                    //BTrecord.AppendEntity(poly_del);
                                    //Trans1.AddNewlyCreatedDBObject(poly_del, true);

                                }

                                double bear1 = Functions.GET_Bearing_rad(point1.X, point1.Y, point2.X, point2.Y) + extra1;


                                double position_measured = first_lbl + i * spacing1;





                                System.Collections.Specialized.StringCollection col_atr = new System.Collections.Specialized.StringCollection();
                                System.Collections.Specialized.StringCollection col_val = new System.Collections.Specialized.StringCollection();

                                col_atr.Add(comboBox_mpkp_attribute.Text);

                                int round1 = 0;

                                if (comboBox_kpmp_units_precision.Text == "0.0")
                                {
                                    round1 = 1;
                                }
                                else if (comboBox_kpmp_units_precision.Text == "0.00")
                                {
                                    round1 = 2;
                                }
                                else if (comboBox_kpmp_units_precision.Text == "0.000")
                                {
                                    round1 = 3;
                                }

                                double fact1 = 5280;
                                if (_AGEN_mainform.units_of_measurement == "m") fact1 = 1000;

                                string val1 = Functions.Get_String_Rounded((position_measured / fact1), round1) + new_suffix;

                                col_val.Add(val1);

                                double scale1 = 1 / _AGEN_mainform.Vw_scale;
                                if (Functions.IsNumeric(textBox_scale_mp.Text) == true)
                                {
                                    scale1 = Convert.ToDouble(textBox_scale_mp.Text);
                                }

                                BlockReference Block1 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "",
                                      comboBox_mpkp_blocks.Text, pt_ins, scale1, bear1, _AGEN_mainform.layer_mp_blocks, col_atr, col_val);
                                Functions.Populate_object_data_table_from_objectid(Tables1, Block1.ObjectId, "Agen_mp_block", Lista_val, Lista_type);


                            }
                        }
                    }
                    Poly3D.Erase();
                    Trans1.Commit();

                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }


        private BlockReference Insert_kpmp_block_at_start(Polyline3d Poly3D, double start1, string Suffix1)
        {
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                    Point3d pt_ins = Poly3D.StartPoint;
                    Point3d point1 = Poly3D.GetPointAtParameter(0);
                    Point3d point2 = Poly3D.GetPointAtParameter(1);
                    double bear1 = Functions.GET_Bearing_rad(point1.X, point1.Y, point2.X, point2.Y);

                    System.Collections.Specialized.StringCollection col_atr = new System.Collections.Specialized.StringCollection();
                    System.Collections.Specialized.StringCollection col_val = new System.Collections.Specialized.StringCollection();

                    BlockReference Block1 = null;

                    col_atr.Add(comboBox_mpkp_attribute.Text);

                    int round1 = 0;

                    if (comboBox_kpmp_units_precision.Text == "0.0")
                    {
                        round1 = 1;
                    }
                    else if (comboBox_kpmp_units_precision.Text == "0.00")
                    {
                        round1 = 2;
                    }
                    else if (comboBox_kpmp_units_precision.Text == "0.000")
                    {
                        round1 = 3;
                    }

                    string val1 = Functions.Get_String_Rounded(start1, round1) + Suffix1;

                    col_val.Add(val1);

                    double scale1 = 1 / _AGEN_mainform.Vw_scale;
                    if (Functions.IsNumeric(textBox_scale_mp.Text) == true)
                    {
                        scale1 = Convert.ToDouble(textBox_scale_mp.Text);
                    }
                    Block1 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "",
                        comboBox_mpkp_blocks.Text, pt_ins, scale1, bear1, _AGEN_mainform.layer_mp_blocks, col_atr, col_val);

                    Poly3D.Erase();
                    Trans1.Commit();
                    return Block1;
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }


        }

        private BlockReference Insert_kpmp_block_at_end(Polyline3d Poly3D, double sta_end, string Suffix1)
        {
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                    Point3d pt_ins = Poly3D.EndPoint;
                    double param1 = Poly3D.EndParam;

                    BlockReference Block1 = null;

                    double param2 = param1 - 1;

                    Point3d point1 = Poly3D.GetPointAtParameter(Math.Floor(param2));

                    Point3d point2 = Poly3D.GetPointAtParameter(Math.Floor(param1));

                    double bear1 = Functions.GET_Bearing_rad(point1.X, point1.Y, point2.X, point2.Y);

                    System.Collections.Specialized.StringCollection col_atr = new System.Collections.Specialized.StringCollection();
                    System.Collections.Specialized.StringCollection col_val = new System.Collections.Specialized.StringCollection();

                    col_atr.Add(comboBox_mpkp_attribute.Text);

                    int round1 = 0;

                    if (comboBox_kpmp_units_precision.Text == "0.0")
                    {
                        round1 = 1;
                    }
                    else if (comboBox_kpmp_units_precision.Text == "0.00")
                    {
                        round1 = 2;
                    }
                    else if (comboBox_kpmp_units_precision.Text == "0.000")
                    {
                        round1 = 3;
                    }


                    double fact1 = 5280;
                    if (_AGEN_mainform.units_of_measurement == "m") fact1 = 1000;

                    string val1 = Functions.Get_String_Rounded((sta_end / fact1), round1) + Suffix1;



                    col_val.Add(val1);

                    double scale1 = 1 / _AGEN_mainform.Vw_scale;
                    if (Functions.IsNumeric(textBox_scale_mp.Text) == true)
                    {
                        scale1 = Convert.ToDouble(textBox_scale_mp.Text);
                    }

                    Block1 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "",
                        comboBox_mpkp_blocks.Text, pt_ins, scale1, bear1, _AGEN_mainform.layer_mp_blocks, col_atr, col_val);

                    Poly3D.Erase();
                    Trans1.Commit();
                    return Block1;
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }

        private void Insert_kp_blocks_with_csf(double sta0, double sta1, Point3d pt0, Point3d pt1, double spacing1,
                                               Autodesk.Gis.Map.ObjectData.Tables Tables1,
                                               List<object> Lista_val, List<Autodesk.Gis.Map.Constants.DataType> Lista_type)
        {
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                    Polyline Poly2D = new Polyline();
                    Poly2D.AddVertexAt(0, new Point2d(pt0.X, pt0.Y), 0, 0, 0);
                    Poly2D.AddVertexAt(1, new Point2d(pt1.X, pt1.Y), 0, 0, 0);

                    if (_AGEN_mainform.dt_station_equation != null && _AGEN_mainform.dt_station_equation.Rows.Count > 0)
                    {

                        string rr_end_x = "Reroute End X";
                        string rr_end_y = "Reroute End Y";

                        for (int i = 0; i < _AGEN_mainform.dt_station_equation.Rows.Count; ++i)
                        {
                            if (_AGEN_mainform.dt_station_equation.Rows[i][rr_end_x] != DBNull.Value &&
                                _AGEN_mainform.dt_station_equation.Rows[i][rr_end_y] != DBNull.Value)
                            {
                                double x1 = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i][rr_end_x]);
                                double y1 = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i][rr_end_y]);

                                Point3d pt_eq = new Point3d(x1, y1, Poly2D.Elevation);

                                Point3d pt_on_poly = Poly2D.GetClosestPointTo(pt_eq, Vector3d.ZAxis, false);
                                double d1 = Math.Pow(Math.Pow(pt_eq.X - pt_on_poly.X, 2) + Math.Pow(pt_eq.Y - pt_on_poly.Y, 2), 0.5);
                                double meas1 = Poly2D.GetDistAtPoint(pt_on_poly);
                                if (d1 < 0.001 && meas1 > 0 && meas1 <= Poly2D.Length)
                                {
                                    double back1 = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i][1]);
                                    double ahead1 = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i][2]);

                                    double fl_b = Math.Ceiling(sta0 / spacing1) * spacing1;
                                    if (fl_b <= back1)
                                    {

                                        int no_blocks = Convert.ToInt32(Math.Ceiling((back1 - fl_b) / spacing1));

                                        if (no_blocks > 0)
                                        {
                                            for (int k = 0; k < no_blocks; ++k)
                                            {
                                                double dist1 = (fl_b - sta0 + k * spacing1) * Poly2D.Length / (back1 - sta0);

                                                Point3d pos1 = Poly2D.GetPointAtDist((dist1));
                                                double position_measured = fl_b + k * spacing1;


                                                double bear1 = Functions.GET_Bearing_rad(pt0.X, pt0.Y, pt1.X, pt1.Y);
                                                if (_AGEN_mainform.Left_to_Right == false) bear1 = bear1 + Math.PI;



                                                System.Collections.Specialized.StringCollection col_atr = new System.Collections.Specialized.StringCollection();
                                                System.Collections.Specialized.StringCollection col_val = new System.Collections.Specialized.StringCollection();

                                                col_atr.Add(comboBox_mpkp_attribute.Text);

                                                int round1 = 0;

                                                if (comboBox_kpmp_units_precision.Text == "0.0")
                                                {
                                                    round1 = 1;
                                                }
                                                else if (comboBox_kpmp_units_precision.Text == "0.00")
                                                {
                                                    round1 = 2;
                                                }
                                                else if (comboBox_kpmp_units_precision.Text == "0.000")
                                                {
                                                    round1 = 3;
                                                }


                                                string val1 = Functions.Get_String_Rounded(position_measured / 1000, round1);

                                                col_val.Add(val1);

                                                double scale1 = 1 / _AGEN_mainform.Vw_scale;
                                                if (Functions.IsNumeric(textBox_scale_mp.Text) == true)
                                                {
                                                    scale1 = Convert.ToDouble(textBox_scale_mp.Text);
                                                }

                                                BlockReference Block1 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "",
                                                      comboBox_mpkp_blocks.Text, pos1, scale1, bear1, _AGEN_mainform.layer_mp_blocks, col_atr, col_val);

                                                Functions.Populate_object_data_table_from_objectid(Tables1, Block1.ObjectId, "Agen_mp_block", Lista_val, Lista_type);

                                            }
                                        }
                                    }

                                    sta0 = ahead1;
                                }
                            }
                        }

                    }
                    double first_lbl = Math.Ceiling(sta0 / spacing1) * spacing1;

                    if (first_lbl <= sta1)
                    {

                        int no_blocks = Convert.ToInt32(Math.Ceiling((sta1 - first_lbl) / spacing1));

                        if (no_blocks > 0)
                        {
                            for (int i = 0; i < no_blocks; ++i)
                            {
                                double dist1 = (first_lbl - sta0 + i * spacing1) * (sta1 - sta0) / Poly2D.Length;

                                Point3d pt_ins = Poly2D.GetPointAtDist((dist1));
                                double position_measured = first_lbl + i * spacing1;


                                double bear1 = Functions.GET_Bearing_rad(pt0.X, pt0.Y, pt1.X, pt1.Y);
                                if (_AGEN_mainform.Left_to_Right == false) bear1 = bear1 + Math.PI;



                                System.Collections.Specialized.StringCollection col_atr = new System.Collections.Specialized.StringCollection();
                                System.Collections.Specialized.StringCollection col_val = new System.Collections.Specialized.StringCollection();

                                col_atr.Add(comboBox_mpkp_attribute.Text);

                                int round1 = 0;

                                if (comboBox_kpmp_units_precision.Text == "0.0")
                                {
                                    round1 = 1;
                                }
                                else if (comboBox_kpmp_units_precision.Text == "0.00")
                                {
                                    round1 = 2;
                                }
                                else if (comboBox_kpmp_units_precision.Text == "0.000")
                                {
                                    round1 = 3;
                                }


                                string val1 = Functions.Get_String_Rounded(position_measured / 1000, round1);

                                col_val.Add(val1);

                                double scale1 = 1 / _AGEN_mainform.Vw_scale;
                                if (Functions.IsNumeric(textBox_scale_mp.Text) == true)
                                {
                                    scale1 = Convert.ToDouble(textBox_scale_mp.Text);
                                }

                                BlockReference Block1 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "",
                                      comboBox_mpkp_blocks.Text, pt_ins, scale1, bear1, _AGEN_mainform.layer_mp_blocks, col_atr, col_val);

                                Functions.Populate_object_data_table_from_objectid(Tables1, Block1.ObjectId, "Agen_mp_block", Lista_val, Lista_type);
                            }
                        }
                    }




                    Trans1.Commit();

                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }


        private void checkBox_pi_show_deflection_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_pi_show_deflection.Checked == false)
            {
                label_defl.Visible = false;
                comboBox_pi_atr_defl.Visible = false;
            }
            else
            {
                label_defl.Visible = true;
                comboBox_pi_atr_defl.Visible = true;
            }
        }

        private void radioButton_new_config_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton_new_config.Checked == true)
            {

                Ag.Set_textBox_config_file_location("");
            }
        }












        private void checkBox_pi_show_stations_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_pi_show_stations.Checked == false)
            {
                label_sta_atr.Visible = false;
                comboBox_pi_atr_sta.Visible = false;
            }
            else
            {
                label_sta_atr.Visible = true;
                comboBox_pi_atr_sta.Visible = true;
            }
        }



        private void button_show_segment_list_Click(object sender, EventArgs e)
        {
            foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
            {
                if (Forma1 is Alignment_mdi.AGEN_segments_form)
                {
                    Forma1.Focus();
                    Forma1.WindowState = System.Windows.Forms.FormWindowState.Normal;
                    Forma1.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - Forma1.Width) / 2,
                      (Screen.PrimaryScreen.WorkingArea.Height - Forma1.Height) / 2);
                    return;
                }
            }
            try
            {
                Alignment_mdi.AGEN_segments_form forma2 = new Alignment_mdi.AGEN_segments_form();
                Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma2);
                forma2.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - forma2.Width) / 2,
                     (Screen.PrimaryScreen.WorkingArea.Height - forma2.Height) / 2);
            }
            catch (System.Exception EX)
            {
                MessageBox.Show(EX.Message);
            }
        }

        private void comboBox_segment_name_SelectedIndexChanged(object sender, EventArgs e)
        {
            _AGEN_mainform.tpage_sheetindex.add_segment_name(comboBox_segment_name.Text);
            
            combo_SelectedIndexChanged(comboBox_segment_name);

            _AGEN_mainform.layer_stationing = _AGEN_mainform.layer_stationing_original;
        }

        public void set_combobox_segment_name()
        {
            comboBox_segment_name.SelectedIndex = comboBox_segment_name.Items.IndexOf(_AGEN_mainform.current_segment);
            _AGEN_mainform.layer_stationing = _AGEN_mainform.layer_stationing_original;
        }

        public void combo_SelectedIndexChanged(ComboBox combo1)
        {
            if (is_loading == false)
            {
                set_enable_false();


                Microsoft.Office.Interop.Excel.Application Excel1 = null;

                try
                {
                    Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                }
                catch (System.Exception ex)
                {
                    Excel1 = new Microsoft.Office.Interop.Excel.Application();

                }

                if (Excel1.Workbooks.Count == 0) Excel1.Visible = _AGEN_mainform.ExcelVisible;



                string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                {
                    ProjF = ProjF + "\\";
                }

                string fisier_cl = ProjF + _AGEN_mainform.cl_excel_name;

                if (System.IO.File.Exists(fisier_cl) == false)
                {
                    Set_centerline_label_to_red();
                    if (Excel1.Workbooks.Count == 0)
                    {
                        Excel1.Quit();
                    }
                    else
                    {
                        Excel1.Visible = true;
                    }
                    if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }
                else
                {
                    _AGEN_mainform.tpage_setup.Load_centerline_and_station_equation(fisier_cl);
                    _AGEN_mainform.tpage_st_eq.Populate_datagridview_with_equation_data();

                    if (_AGEN_mainform.dt_centerline != null && _AGEN_mainform.dt_centerline.Rows.Count > 0)
                    {
                        Set_centerline_label_to_green();

                        if (combo1.Text.Replace(" ", "") != "")
                        {
                            _AGEN_mainform.current_segment = combo1.Text;

                            Functions.set_regular_band_insertion_points();

                            push_selected_index_changed();
                        }
                        #region sheet index

                        Microsoft.Office.Interop.Excel.Workbook Workbook2 = null;
                        Microsoft.Office.Interop.Excel.Worksheet W2 = null;
                        Microsoft.Office.Interop.Excel.Workbook Workbook3 = null;
                        try
                        {

                            if (Excel1 == null)
                            {
                                MessageBox.Show("PROBLEM WITH EXCEL!");
                                return;
                            }

                            string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                            string fisier_si = ProjFolder + _AGEN_mainform.sheet_index_excel_name;
                            if (System.IO.File.Exists(fisier_si) == true)
                            {
                                Workbook2 = Excel1.Workbooks.Open(fisier_si);
                                W2 = Workbook2.Worksheets[1];
                                _AGEN_mainform.dt_sheet_index = Functions.Build_Data_table_sheet_index_from_excel(W2, _AGEN_mainform.Start_row_Sheet_index + 1);

                                Workbook2.Close();
                            }
                            else
                            {
                                _AGEN_mainform.dt_sheet_index = null;
                            }

                            _AGEN_mainform.tpage_sheetindex.set_dataGridView_sheet_index();

                            if (combo1.Text.Replace(" ", "") != "")
                            {
                                Workbook3 = Excel1.Workbooks.Open(_AGEN_mainform.config_path);
                                foreach (Microsoft.Office.Interop.Excel.Worksheet W3 in Workbook3.Worksheets)
                                {
                                    if (W3.Name == _AGEN_mainform.first_custom_band + "_cfg_" + _AGEN_mainform.current_segment)
                                    {
                                        transfer_custom_band_settings_to_controls(W3);
                                    }
                                    if (W3.Name == "pdc2_" + _AGEN_mainform.current_segment)
                                    {
                                        transfer_profile_settings_to_controls(W3);
                                    }

                                }
                                Workbook3.Close();
                            }

                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                        finally
                        {
                            if (Excel1.Workbooks.Count == 0)
                            {
                                Excel1.Quit();
                            }
                            else
                            {
                                Excel1.Visible = true;
                            }
                            if (W2 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W2);
                            if (Workbook2 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook2);
                            if (Workbook3 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook3);
                            if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                        }
                        #endregion
                    }
                    else
                    {
                        Set_centerline_label_to_red();
                    }
                }
                set_enable_true();
            }
        }




        private void radioButton_usa_CheckedChanged(object sender, EventArgs e)
        {

            if (radioButton_usa.Checked == true)
            {
                panel_canada.Visible = false;
                _AGEN_mainform.COUNTRY = "USA";
                panel_USA.Visible = true;
                panel_canada.Location = new Point(-1, 64);

            }
            else
            {
                _AGEN_mainform.COUNTRY = "CANADA";
                panel_canada.Visible = true;
                panel_canada.Location = new Point(-1, 3);
                panel_USA.Visible = false;

            }

        }

        public void set_display_to_feet_or_meters()
        {
            if (_AGEN_mainform.units_of_measurement == "m")
            {
                label_kpmp_precision.Text = "KP Precision";
                label_kpmp_um.Text = "Kilometer";
                label_kpmp_block.Text = "KP Block";
                label_kpmp_Block_attrib.Text = "KP Block Attribute";
                checkBox_mpkp_eq_suffix.Text = "Create KP Suffix";
                _AGEN_mainform.tpage_sheetindex.set_radioButton_use3D_stations(true);
                _AGEN_mainform.Project_type = "3D";
                _AGEN_mainform.tpage_crossing_draw.set_checkBox_include_property_lines(false);
                _AGEN_mainform.tpage_crossing_draw.set_checkBox_split_station_value(true);
                _AGEN_mainform.tpage_crossing_draw.set_checkBox_draw_angle_symbol_value(true);
            }
            else
            {
                label_kpmp_precision.Text = "MP Precision";
                label_kpmp_um.Text = "Mile";
                label_kpmp_block.Text = "MP Block";
                label_kpmp_Block_attrib.Text = "MP Block Attribute";
                checkBox_mpkp_eq_suffix.Text = "Create MP Suffix";
                _AGEN_mainform.tpage_crossing_draw.set_checkBox_include_property_lines(true);
                _AGEN_mainform.tpage_crossing_draw.set_checkBox_split_station_value(false);
                _AGEN_mainform.tpage_crossing_draw.set_checkBox_draw_angle_symbol_value(false);
            }

        }

        private void button_display_tpage_load_cl_xl_Click(object sender, EventArgs e)
        {
            _AGEN_mainform.tpage_processing.Hide();
            _AGEN_mainform.tpage_blank.Hide();
            _AGEN_mainform.tpage_setup.Hide();
            _AGEN_mainform.tpage_viewport_settings.Hide();
            _AGEN_mainform.tpage_tblk_attrib.Hide();
            _AGEN_mainform.tpage_band_analize.Hide();
            _AGEN_mainform.tpage_sheetindex.Hide();
            _AGEN_mainform.tpage_layer_alias.Hide();
            _AGEN_mainform.tpage_crossing_scan.Hide();
            _AGEN_mainform.tpage_crossing_draw.Hide();
            _AGEN_mainform.tpage_profilescan.Hide();
            _AGEN_mainform.tpage_profdraw.Hide();
            _AGEN_mainform.tpage_owner_scan.Hide();
            _AGEN_mainform.tpage_owner_draw.Hide();
            _AGEN_mainform.tpage_mat.Hide();
            _AGEN_mainform.tpage_cust_scan.Hide();
            _AGEN_mainform.tpage_cust_draw.Hide();
            _AGEN_mainform.tpage_sheet_gen.Hide();

            _AGEN_mainform.tpage_tools.Hide();
            _AGEN_mainform.tpage_st_eq.Hide();
            _AGEN_mainform.tpage_cl_xl.Show();
        }

        public int get_no_segments()
        {
            return comboBox_segment_name.Items.Count;
        }

        public void transfer_segment_data()
        {
            Fill_combobox_segments();
            _AGEN_mainform.tpage_st_eq.Fill_combobox_segments();
            _AGEN_mainform.tpage_sheetindex.Fill_combobox_segments();
            _AGEN_mainform.tpage_crossing_scan.Fill_combobox_segments();
            _AGEN_mainform.tpage_crossing_draw.Fill_combobox_segments();
            _AGEN_mainform.tpage_profilescan.Fill_combobox_segments();
            _AGEN_mainform.tpage_profdraw.Fill_combobox_segments();
            _AGEN_mainform.tpage_owner_scan.Fill_combobox_segments();
            _AGEN_mainform.tpage_owner_draw.Fill_combobox_segments();
            _AGEN_mainform.tpage_mat.Fill_combobox_segments();
            _AGEN_mainform.tpage_cust_scan.Fill_combobox_segments();
            _AGEN_mainform.tpage_cust_draw.Fill_combobox_segments();
            _AGEN_mainform.tpage_sheet_gen.Fill_combobox_segments();
            _AGEN_mainform.tpage_cl_xl.Fill_combobox_segments();
        }

        private void push_selected_index_changed()
        {
            _AGEN_mainform.tpage_st_eq.set_combobox_segment_name();
            _AGEN_mainform.tpage_sheetindex.set_combobox_segment_name();
            _AGEN_mainform.tpage_crossing_scan.set_combobox_segment_name();
            _AGEN_mainform.tpage_crossing_draw.set_combobox_segment_name();
            _AGEN_mainform.tpage_profilescan.set_combobox_segment_name();
            _AGEN_mainform.tpage_profdraw.set_combobox_segment_name();
            _AGEN_mainform.tpage_owner_scan.set_combobox_segment_name();
            _AGEN_mainform.tpage_owner_draw.set_combobox_segment_name();
            _AGEN_mainform.tpage_mat.set_combobox_segment_name();
            _AGEN_mainform.tpage_cust_scan.set_combobox_segment_name();
            _AGEN_mainform.tpage_cust_draw.set_combobox_segment_name();
            _AGEN_mainform.tpage_sheet_gen.set_combobox_segment_name();
            _AGEN_mainform.tpage_cl_xl.set_combobox_segment_name();
        }
    }
}
