using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace Alignment_mdi
{
    public partial class Mat_Design_form : Form
    {
        //Global Variables
        public static System.Data.DataTable dt_mat_library = null;
        public static System.Data.DataTable dt_filter = null;
        public static List<string> category_list = null;
        public static Mat_Design_form tpage_matl_design = null;

        string col_MMID = "MMID";
        string col_Item_No = "ItemNo";

        string Col_3DSta = "3DSta";


        string pipes_layer = "Pipes";
        string pipes_od = "Pipes";

        string elbows_layer = "Elbows";
        string elbows_od = "Elbows";


        string fab_layer = "Assembly";
        string fab_od = "Assembly";

        string class_layer = "Class";
        string class_od = "Class";



        string buoyancy_layer = "Buoyancy";
        string buoyancy_pt_layer = "BuoyancyPT";
        string buoyancy_od = "Buoyancy";
        string buoyancy_pt_od = "BuoyancyPT";

        string cpac_layer = "CPAC";
        string cpac_od = "CPAC";

        string es_layer = "EandS";
        string es_od = "EandS";

        string hydrotest_layerPT = "HydrostaticTestPT";
        string hydrotest_odPT = "HydrostaticTestPT";

        string hydrotest_layer = "HydrostaticTest";
        string hydrotest_od = "HydrostaticTest";

        string xing_layer = "Crossing";
        string xing_od = "Crossing";

        string geotech_layer = "Geohazard";
        string geotech_layer_pt = "GeohazardPT";
        string geotech_od = "Geohazard";
        string geotech_od_pt = "GeohazardPT";

        string muskeg_layer = "Musk_drainage";
        string muskeg_od = "Musk_drainage";

        string preexisting_layer = "Pre_existing_pipe";
        string preexisting_od = "Pre_existing_pipe";

        string transition_layer = "Transition";
        string transition_od = "Transition";

        string doc_layer = "Depth_of_Cover";
        string doc_od = "Depth_of_Cover";


        string col_geotech_sta1 = "Begin Station";
        string col_geotech_od_sta1 = "Start";
        string xl_geotech_sta1 = "E";

        string col_geotech_sta2 = "End Station";
        string col_geotech_od_sta2 = "End";
        string xl_geotech_sta2 = "I";

        string col_geotech_descr1 = "Begin Description";
        string col_geotech_od_descr1 = "Beg_Descr";
        string xl_geotech_descr1 = "F";

        string col_geotech_descr2 = "End Description";
        string col_geotech_od_descr2 = "End_Descr";
        string xl_geotech_descr2 = "J";

        string col_geotech_class = "Geohazard Class";
        string col_geotech_od_class = "Class";
        string xl_geotech_class = "N";

        string col_geotech_type = "Geohazard Type";
        string col_geotech_od_type = "Type";
        string xl_geotech_type = "O";

        string col_geotech_label = "Golder_Hazard_Label";
        string col_geotech_od_label = "Golder_Lbl";
        string xl_geotech_label = "P";



        string col_doc_sta1 = "Begin Station";
        string col_doc_od_sta1 = "Start";
        string xl_doc_sta1 = "F";

        string col_doc_sta2 = "End Station";
        string col_doc_od_sta2 = "End";
        string xl_doc_sta2 = "K";

        string col_doc_min_cvr = "Minimum Depth of Cover";
        string col_doc_od_min_cvr = "MinCVR";
        string xl_doc_min_cvr = "Q";

        string tab_doc = "DOC&Ditch";


        string col_muskeg_sta1 = "Begin Station";
        string col_muskeg_od_sta = "Station";
        string xl_muskeg_sta1 = "E";

        string col_muskeg_sta2 = "End Station";
        string xl_muskeg_sta2 = "I";

        string col_muskeg_descr1 = "Begin Description";
        string col_muskeg_od_descr = "Descr";
        string xl_muskeg_descr1 = "F";

        string col_muskeg_descr2 = "End Description";
        string xl_muskeg_descr2 = "J";


        string col_muskeg_label = "Golder Label";
        string col_muskeg_od_label = "Golder_Label";
        string xl_muskeg_label = "O";

        string tab_muskeg = "Muskeg&Drainage";


        string col_od_pipe_type = "PipeType";
        string col_od_wt = "WT";
        string col_od_coat = "Coating";
        string col_od_class = "Class";
        string col_od_mat_descr = "MatDescr";
        string col_od_descr = "Descr";
        string col_od_start = "Start";
        string col_od_end = "End";




        string col_elbow_ref_id = "Reference_id";

        string col_ref_dwg_id = "Reference dwg";
        string col_od_ref_dwg_id = "Ref_dwg";
        string col_od_min_depth = "Min_cvr";
        string col_min_depth = "Minimum cover";
        string col_agen_cvr = "Agen_cvr";
        string col_xing_method = "Crossing_method";
        string col_od_xing_method = "XingMethod";

        string col_feature = "Feature";
        string col_spacing = "Spacing";
        string col_count = "FeatCount";
        string col_start_elbow_adjacent = "Start";
        string col_end_elbow_adjacent = "End";
        string col_start = "Start";
        string col_end = "End";

        string col_just = "Justif";
        string col_length = "Length";

        string col_sta = "STA";
        string col_station = "Station";
        string col_eq1 = "Equipment1";
        string col_eq2 = "Equipment2";
        string col_eq3 = "Equipment3";
        string col_descr = "Descr";
        string col_descr1 = "Descr1";
        string col_descr2 = "Descr2";

        string col_tst_sec = "Test Section";


        string col_od_descr1 = "Descr_start";
        string col_od_descr2 = "Descr_end";

        string col_ditchplug = "DitchPlug";
        string col_xingid = "Xing ID";
        string col_od_xingid = "XingID";
        string col_od_xingtype = "XingType";
        string col_xingtype = "Xing Type";
        string col_notes = "Notes";
        string col_sta1 = "Sta1";
        string col_sta2 = "Sta2";
        string col_len = "Length";

        string col_fac_name = "Name";

        string col_es_spacing = "Feature_spacing";

        string col_elbow_id = "Elbow ID";
        string col_od_elbow_id = "ElbowID";
        string xl_elbow_id = "E";

        string col_bend_type = "Bend Type";
        string col_od_bend_type = "BendType";
        string xl_bend_type = "H";
        string col_elbow_ref_dwg = "Reference Drawing ID";
        string col_od_elbow_ref_dwg = "RefDwgID";
        string xl_elbow_ref_dwg = "K";


        string col_od_pipe_class = "PipeClass";
        string xl_pipe_class = "Q";

        string col_elbow_angle = "Angle (deg)";
        string col_od_elbow_angle = "Angle";
        string col_elbow_notes = "Notes";
        string xl_elbow_notes = "S";

        string col_TL = "Total Length";

        string col_radius = "Radius";
        string col_pup = "Pup Length\r\n(multiply by 2)\r\nfor checks";
        string xl_elbow_pipe_type = "P";
        string xl_elbow_sta1 = "M";
        string xl_elbow_sta2 = "N";
        string xl_elbow_descr = "I";
        string xl_elbow_wt = "R";
        string xl_elbow_just = "S";
        string xl_elbow_additional_start = "T";
        string xl_elbow_additional_end = "V";
        string xl_elbow_pi = "L";
        string xl_elbow_angle = "G";
        string xl_elbow_defl = "H";

        string col_elbow_pi = "PI";

        string xl_hydrotest_sta1 = "F";
        string xl_hydrotest_descr1 = "G";
        string xl_hydrotest_sta2 = "K";
        string xl_hydrotest_descr2 = "L";

        string col_pipe_type = "Pipe Type";
        string col_wt = "Wall Thickness";
        string col_dwg = "Drawing";
        string col_m1 = "M1";
        string col_m2 = "M2";

        string excel_cell = "Excel";
        string excel_cell1 = "Excel1";

        System.Drawing.Font font10 = null;
        System.Drawing.Font font8 = null;



        string col_2dbeg = "2DStaBeg";
        string col_2dsta = "2DSta";
        string col_2dend = "2DStaEnd";
        string col_3dbeg = "3DStaBeg";
        string col_3dsta = "3DSta";
        string col_3dend = "3DStaEnd";
        string col_eqStabeg = "EqStaBeg";
        string col_eqSta = "EqSta";
        string col_eqStaend = "EqStaEnd";
        string col_2dlen = "2D Length";
        string col_3dlen = "3D Length";
        string col_altdesc = "AltDesc";
        string col_symbol = "Symbol";
        string col_xbeg = "X_Beg";
        string col_ybeg = "Y_Beg";
        string col_xend = "X_End";
        string col_yend = "Y_End";
        string col_x = "X";
        string col_y = "Y";
        string col_z = "Z";
        string col_block = "BLOCK";
        string col_blockdescr = "DESCR";
        string col_note1 = "NOTE1";
        string col_qty = "QTY";

        string col_mat = "MAT";

        string col_id = "ID";
        string col_id2 = "ID2";
        string col_cvr = "CVR";

        string col_mstartcanada = "MeasuredStartCanada";
        string col_mendcanada = "MeasuredEndCanada";
        string col_mcanada = "MeasuredCanada";
        string col_visibility = "Visibility";

        string col_defl = "Deflection";
        string col_is_elbow = "IS Elbow";


        public Mat_Design_form()
        {
            InitializeComponent();
            tpage_matl_design = this;
            font10 = new System.Drawing.Font("Arial", 10f, FontStyle.Bold);
            font8 = new System.Drawing.Font("Arial", 8f, FontStyle.Bold);

            ComboBox_nps.Items.Add("NPS 1/2");
            ComboBox_nps.Items.Add("NPS 1");
            ComboBox_nps.Items.Add("NPS 1.5");
            ComboBox_nps.Items.Add("NPS 2");
            ComboBox_nps.Items.Add("NPS 3");
            ComboBox_nps.Items.Add("NPS 4");
            ComboBox_nps.Items.Add("NPS 6");
            ComboBox_nps.Items.Add("NPS 8");
            ComboBox_nps.Items.Add("NPS 10");
            ComboBox_nps.Items.Add("NPS 12");
            ComboBox_nps.Items.Add("NPS 16");
            ComboBox_nps.Items.Add("NPS 20");
            ComboBox_nps.Items.Add("NPS 24");
            ComboBox_nps.Items.Add("NPS 30");
            ComboBox_nps.Items.Add("NPS 36");
            ComboBox_nps.Items.Add("NPS 42");
            ComboBox_nps.Items.Add("NPS 48");
            ComboBox_nps.SelectedIndex = 16;

            label_result_nps_radius.Text = Convert.ToString(get_from_NPS_radius_for_pipes_from_inches_to_milimeters(48) / 1000);
            if (Functions.IsNumeric(textBox_field_bend_multiplier.Text) == true)
            {
                double fbm = Convert.ToDouble(textBox_field_bend_multiplier.Text);
                label_result_field_bend_radius.Text = Convert.ToString(2 * fbm * get_from_NPS_radius_for_pipes_from_inches_to_milimeters(48) / 1000);

            }

            if (Functions.IsNumeric(textBox_elbow_multiplier.Text) == true)
            {
                double em = Convert.ToDouble(textBox_elbow_multiplier.Text);
                label_result_elbow_radius.Text = Convert.ToString(2 * em * get_from_NPS_radius_for_pipes_from_inches_to_milimeters(48) / 1000);

            }

        }






        #region set enable true or false    
        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();


            lista_butoane.Add(comboBox_xl_canada);
            lista_butoane.Add(button_load_eng_db_canada);



            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = false;
            }
        }
        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();



            lista_butoane.Add(comboBox_xl_canada);
            lista_butoane.Add(button_load_eng_db_canada);





            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }
        #endregion

        #region hector work
        public string ToTitleCase(string string1)
        {
            TextInfo textinfo1 = new CultureInfo("en-us", false).TextInfo;

            string1 = textinfo1.ToTitleCase(string1.ToLower());
            Console.WriteLine(string1);

            return string1;
        }




        private void comboBox_xl_DropDown(object sender, EventArgs e)
        {
            ComboBox combo1 = sender as ComboBox;
            Functions.Load_opened_workbooks_to_combobox(combo1);
            combo1.DropDownWidth = Functions.get_dropdown_width(combo1);

        }
        #endregion

        private void button_load_eng_db_Click(object sender, EventArgs e)
        {
            int start1 = 3;
            int end1 = 1000;
            string col_pipe_class = "Pipe Class";
            string col_od_pipe_class = "PipeClass";
            string col_coating = "Coating";


            string col_src = "Source";

            string col_defl = "Deflection";
            string col_back = "Back";
            string col_ahead = "Ahead";


            #region data table defs
            string tab_materials = "Materials";
            System.Data.DataTable dt_materials = new System.Data.DataTable();
            dt_materials.Columns.Add(col_pipe_type, typeof(string));
            dt_materials.Columns.Add(col_descr, typeof(string));
            dt_materials.Columns.Add(col_pipe_class, typeof(string));
            dt_materials.Columns.Add(col_coating, typeof(string));
            dt_materials.Columns.Add(col_wt, typeof(string));
            dt_materials.Columns.Add(col_notes, typeof(string));
            dt_materials.Columns.Add(excel_cell, typeof(int));
            dt_materials.Columns.Add("pipe", typeof(double));
            dt_materials.Columns.Add("elbow", typeof(double));

            System.Data.DataTable dt_counts = new System.Data.DataTable();
            dt_counts.Columns.Add(col_dwg, typeof(string));
            dt_counts.Columns.Add(col_m1, typeof(double));
            dt_counts.Columns.Add(col_m2, typeof(double));

            System.Data.DataTable dt_eq = new System.Data.DataTable();
            dt_eq.Columns.Add(col_back, typeof(double));
            dt_eq.Columns.Add(col_ahead, typeof(double));

            string tab_class = "ClassLocation";
            System.Data.DataTable dt_class = new System.Data.DataTable();
            dt_class.Columns.Add(col_sta1, typeof(double));
            dt_class.Columns.Add(col_sta2, typeof(double));
            dt_class.Columns.Add(col_pipe_type, typeof(string));
            dt_class.Columns.Add(col_wt, typeof(double));
            dt_class.Columns.Add(col_descr, typeof(string));
            dt_class.Columns.Add(excel_cell, typeof(int));
            dt_class.Columns.Add(col_descr1, typeof(string));
            dt_class.Columns.Add(col_descr2, typeof(string));

            string tab_stress = "HW-Stress";
            System.Data.DataTable dt_stress = new System.Data.DataTable();
            dt_stress.Columns.Add(col_sta1, typeof(double));
            dt_stress.Columns.Add(col_sta2, typeof(double));
            dt_stress.Columns.Add(col_pipe_type, typeof(string));
            dt_stress.Columns.Add(col_wt, typeof(double));
            dt_stress.Columns.Add(col_just, typeof(string));
            dt_stress.Columns.Add(col_notes, typeof(string));
            dt_stress.Columns.Add(excel_cell, typeof(int));
            dt_stress.Columns.Add(excel_cell1, typeof(int));

            string tab_geotech = "GeoHazards";
            System.Data.DataTable dt_geotech = new System.Data.DataTable();
            dt_geotech.Columns.Add(col_geotech_sta1, typeof(double));
            dt_geotech.Columns.Add(col_geotech_sta2, typeof(double));
            dt_geotech.Columns.Add(col_geotech_descr1, typeof(string));
            dt_geotech.Columns.Add(col_geotech_descr2, typeof(string));
            dt_geotech.Columns.Add(col_geotech_class, typeof(string));
            dt_geotech.Columns.Add(col_geotech_type, typeof(string));
            dt_geotech.Columns.Add(col_geotech_label, typeof(string));
            dt_geotech.Columns.Add(col_len, typeof(double));
            dt_geotech.Columns.Add(col_notes, typeof(string));

            dt_geotech.Columns.Add(col_sta1, typeof(double));
            dt_geotech.Columns.Add(col_sta2, typeof(double));
            dt_geotech.Columns.Add(col_pipe_type, typeof(string));
            dt_geotech.Columns.Add(col_wt, typeof(double));
            dt_geotech.Columns.Add(col_just, typeof(string));
            dt_geotech.Columns.Add(excel_cell, typeof(int));

            dt_geotech.Columns.Add("x1", typeof(double));
            dt_geotech.Columns.Add("y1", typeof(double));
            dt_geotech.Columns.Add("x2", typeof(double));
            dt_geotech.Columns.Add("y2", typeof(double));
            dt_geotech.Columns.Add("layer", typeof(string));

            string tab_fab = "HW-Facility";
            System.Data.DataTable dt_fab = new System.Data.DataTable();
            dt_fab.Columns.Add(col_sta1, typeof(double));
            dt_fab.Columns.Add(col_sta2, typeof(double));
            dt_fab.Columns.Add(col_fac_name, typeof(string));
            dt_fab.Columns.Add(col_descr1, typeof(string));
            dt_fab.Columns.Add(col_descr2, typeof(string));
            dt_fab.Columns.Add(col_pipe_type, typeof(string));
            dt_fab.Columns.Add(col_wt, typeof(string));
            dt_fab.Columns.Add(col_just, typeof(string));
            dt_fab.Columns.Add(excel_cell, typeof(int));


            string tab_elbow = "Elbow Data";
            System.Data.DataTable dt_elbow = new System.Data.DataTable();
            dt_elbow.Columns.Add(col_sta1, typeof(double));
            dt_elbow.Columns.Add(col_sta2, typeof(double));
            dt_elbow.Columns.Add(col_pipe_type, typeof(string));
            dt_elbow.Columns.Add(col_wt, typeof(double));
            dt_elbow.Columns.Add(col_descr, typeof(string));
            dt_elbow.Columns.Add(col_elbow_pi, typeof(double));
            dt_elbow.Columns.Add(col_start_elbow_adjacent, typeof(double));
            dt_elbow.Columns.Add(col_end_elbow_adjacent, typeof(double));
            dt_elbow.Columns.Add(col_just, typeof(string));
            dt_elbow.Columns.Add(excel_cell, typeof(int));
            dt_elbow.Columns.Add(col_elbow_angle, typeof(double));
            dt_elbow.Columns.Add(col_defl, typeof(string));
            dt_elbow.Columns.Add(col_elbow_id, typeof(string));
            dt_elbow.Columns.Add(col_bend_type, typeof(string));
            dt_elbow.Columns.Add(col_elbow_ref_dwg, typeof(string));
            dt_elbow.Columns.Add(col_pipe_class, typeof(string));
            dt_elbow.Columns.Add(col_elbow_notes, typeof(string));
            dt_elbow.Columns.Add(col_elbow_ref_id, typeof(string));

            string tab_crossings = "Crossings";
            System.Data.DataTable dt_crossing = new System.Data.DataTable();
            dt_crossing.Columns.Add(col_sta1, typeof(double));
            dt_crossing.Columns.Add(col_sta2, typeof(double));
            dt_crossing.Columns.Add(col_pipe_type, typeof(string));
            dt_crossing.Columns.Add(col_wt, typeof(double));
            dt_crossing.Columns.Add(col_just, typeof(string));
            dt_crossing.Columns.Add(col_sta, typeof(double));
            dt_crossing.Columns.Add(excel_cell, typeof(int));

            string tab_hydro = "HW-Hydrostatic";
            System.Data.DataTable dt_hydro = new System.Data.DataTable();
            dt_hydro.Columns.Add(col_sta1, typeof(double));
            dt_hydro.Columns.Add(col_sta2, typeof(double));
            dt_hydro.Columns.Add(col_pipe_type, typeof(string));
            dt_hydro.Columns.Add(col_wt, typeof(double));
            dt_hydro.Columns.Add(col_just, typeof(string));
            dt_hydro.Columns.Add(col_notes, typeof(string));
            dt_hydro.Columns.Add(excel_cell, typeof(int));

            string tab_buoyancy = "Buoyancy";
            System.Data.DataTable dt_buoy = new System.Data.DataTable();
            dt_buoy.Columns.Add(col_feature, typeof(string));
            dt_buoy.Columns.Add(col_start, typeof(double));
            dt_buoy.Columns.Add(col_end, typeof(double));
            dt_buoy.Columns.Add(col_spacing, typeof(double));
            dt_buoy.Columns.Add(col_count, typeof(double));
            dt_buoy.Columns.Add(col_descr1, typeof(string));
            dt_buoy.Columns.Add(col_descr2, typeof(string));
            dt_buoy.Columns.Add(col_just, typeof(string));
            dt_buoy.Columns.Add(col_notes, typeof(string));
            dt_buoy.Columns.Add(excel_cell, typeof(int));
            dt_buoy.Columns.Add("x1", typeof(double));
            dt_buoy.Columns.Add("y1", typeof(double));
            dt_buoy.Columns.Add("x2", typeof(double));
            dt_buoy.Columns.Add("y2", typeof(double));
            dt_buoy.Columns.Add("layer", typeof(string));
            dt_buoy.Columns.Add("ci", typeof(short));
            System.Data.DataTable dt_long_strap = new System.Data.DataTable();
            dt_long_strap = dt_buoy.Clone();

            string tab_cpac = "CPAC";
            System.Data.DataTable dt_cpac = new System.Data.DataTable();
            dt_cpac.Columns.Add(col_sta, typeof(double));
            dt_cpac.Columns.Add(col_descr, typeof(string));
            dt_cpac.Columns.Add(col_descr2, typeof(string));
            dt_cpac.Columns.Add(col_eq1, typeof(string));
            dt_cpac.Columns.Add(col_eq2, typeof(string));
            dt_cpac.Columns.Add(col_eq3, typeof(string));
            dt_cpac.Columns.Add(col_just, typeof(string));
            dt_cpac.Columns.Add(col_notes, typeof(string));
            dt_cpac.Columns.Add(excel_cell, typeof(int));
            dt_cpac.Columns.Add("x1", typeof(double));
            dt_cpac.Columns.Add("y1", typeof(double));
            dt_cpac.Columns.Add("layer", typeof(string));
            dt_cpac.Columns.Add("ci", typeof(short));

            string tab_es = "E&S";
            System.Data.DataTable dt_es = new System.Data.DataTable();
            dt_es.Columns.Add(col_sta, typeof(double));
            dt_es.Columns.Add(col_ditchplug, typeof(string));
            dt_es.Columns.Add(col_es_spacing, typeof(string));
            dt_es.Columns.Add(col_notes, typeof(string));
            dt_es.Columns.Add(excel_cell, typeof(int));
            dt_es.Columns.Add("x1", typeof(double));
            dt_es.Columns.Add("y1", typeof(double));
            dt_es.Columns.Add("layer", typeof(string));
            dt_es.Columns.Add("ci", typeof(short));


            string tab_pre_existing_pipe = "Pre-Existing Pipe";
            System.Data.DataTable dt_pre_existing = new System.Data.DataTable();
            dt_pre_existing.Columns.Add(col_start, typeof(double));
            dt_pre_existing.Columns.Add(col_end, typeof(double));
            dt_pre_existing.Columns.Add(col_descr, typeof(string));
            dt_pre_existing.Columns.Add(col_just, typeof(string));
            dt_pre_existing.Columns.Add(col_notes, typeof(string));

            string tab_hydrotest = "HydroTestSections";
            System.Data.DataTable dt_hydrotest = new System.Data.DataTable();
            dt_hydrotest.Columns.Add(col_sta1, typeof(double));
            dt_hydrotest.Columns.Add(col_descr1, typeof(string));
            dt_hydrotest.Columns.Add(col_sta2, typeof(double));
            dt_hydrotest.Columns.Add(col_descr2, typeof(string));
            dt_hydrotest.Columns.Add(col_tst_sec, typeof(string));
            dt_hydrotest.Columns.Add(excel_cell, typeof(int));
            dt_hydrotest.Columns.Add("x1", typeof(double));
            dt_hydrotest.Columns.Add("y1", typeof(double));
            dt_hydrotest.Columns.Add("x2", typeof(double));
            dt_hydrotest.Columns.Add("y2", typeof(double));
            dt_hydrotest.Columns.Add("layer", typeof(string));
            dt_hydrotest.Columns.Add("ci", typeof(short));

            System.Data.DataTable dt_xing = new System.Data.DataTable();
            dt_xing.Columns.Add(col_xingid, typeof(string));
            dt_xing.Columns.Add(col_xingtype, typeof(string));
            dt_xing.Columns.Add(col_sta, typeof(double));
            dt_xing.Columns.Add(col_descr1, typeof(string));
            dt_xing.Columns.Add(col_descr2, typeof(string));
            dt_xing.Columns.Add(col_ref_dwg_id, typeof(string));
            dt_xing.Columns.Add(col_min_depth, typeof(string));
            dt_xing.Columns.Add(col_agen_cvr, typeof(string));
            dt_xing.Columns.Add(col_xing_method, typeof(string));
            dt_xing.Columns.Add(col_pipe_type, typeof(string));
            dt_xing.Columns.Add(col_pipe_class, typeof(string));
            dt_xing.Columns.Add(col_wt, typeof(double));
            dt_xing.Columns.Add(col_just, typeof(string));
            dt_xing.Columns.Add(excel_cell, typeof(int));
            dt_xing.Columns.Add("x1", typeof(double));
            dt_xing.Columns.Add("y1", typeof(double));
            dt_xing.Columns.Add("layer", typeof(string));
            dt_xing.Columns.Add("ci", typeof(short));

            System.Data.DataTable dt_doc = new System.Data.DataTable();
            dt_doc.Columns.Add(col_doc_sta1, typeof(double));
            dt_doc.Columns.Add(col_doc_sta2, typeof(double));
            dt_doc.Columns.Add(col_doc_min_cvr, typeof(string));
            dt_doc.Columns.Add(col_len, typeof(double));
            dt_doc.Columns.Add(col_just, typeof(string));
            dt_doc.Columns.Add(col_notes, typeof(string));
            dt_doc.Columns.Add(col_descr1, typeof(string));
            dt_doc.Columns.Add(col_descr2, typeof(string));

            dt_doc.Columns.Add("x1", typeof(double));
            dt_doc.Columns.Add("y1", typeof(double));
            dt_doc.Columns.Add("x2", typeof(double));
            dt_doc.Columns.Add("y2", typeof(double));
            dt_doc.Columns.Add("layer", typeof(string));

            System.Data.DataTable dt_muskeg = new System.Data.DataTable();
            dt_muskeg.Columns.Add(col_muskeg_sta1, typeof(double));
            dt_muskeg.Columns.Add(col_muskeg_sta2, typeof(double));
            dt_muskeg.Columns.Add(col_muskeg_descr1, typeof(string));
            dt_muskeg.Columns.Add(col_muskeg_descr2, typeof(string));
            dt_muskeg.Columns.Add(col_muskeg_label, typeof(string));
            dt_muskeg.Columns.Add("x1", typeof(double));
            dt_muskeg.Columns.Add("y1", typeof(double));
            dt_muskeg.Columns.Add("x2", typeof(double));
            dt_muskeg.Columns.Add("y2", typeof(double));
            dt_muskeg.Columns.Add("layer", typeof(string));

            System.Data.DataTable dt_compiled = new System.Data.DataTable();
            dt_compiled.Columns.Add(col_sta1, typeof(double));
            dt_compiled.Columns.Add(col_sta2, typeof(double));
            dt_compiled.Columns.Add(col_len, typeof(double));
            dt_compiled.Columns.Add(col_pipe_type, typeof(string));
            dt_compiled.Columns.Add(col_descr, typeof(string));
            dt_compiled.Columns.Add(col_pipe_class, typeof(string));
            dt_compiled.Columns.Add(col_wt, typeof(double));
            dt_compiled.Columns.Add(col_just, typeof(string));
            dt_compiled.Columns.Add(col_src, typeof(string));
            dt_compiled.Columns.Add(col_elbow_pi, typeof(double));
            dt_compiled.Columns.Add(col_coating, typeof(string));
            dt_compiled.Columns.Add(col_defl, typeof(string));
            dt_compiled.Columns.Add(col_elbow_ref_id, typeof(string));
            dt_compiled.Columns.Add(col_notes, typeof(string));

            System.Data.DataTable dt1 = new System.Data.DataTable();
            dt1 = dt_compiled.Clone();

            System.Data.DataTable dt2 = new System.Data.DataTable();
            dt2 = dt_compiled.Clone();

            System.Data.DataTable dt_mat_lin = new System.Data.DataTable();
            dt_mat_lin.Columns.Add(col_MMID, typeof(string));
            dt_mat_lin.Columns.Add(col_Item_No, typeof(string));
            dt_mat_lin.Columns.Add(col_2dbeg, typeof(double));
            dt_mat_lin.Columns.Add(col_2dend, typeof(double));
            dt_mat_lin.Columns.Add(col_3dbeg, typeof(double));
            dt_mat_lin.Columns.Add(col_3dend, typeof(double));
            dt_mat_lin.Columns.Add(col_eqStabeg, typeof(double));
            dt_mat_lin.Columns.Add(col_eqStaend, typeof(double));
            dt_mat_lin.Columns.Add(col_2dlen, typeof(double));
            dt_mat_lin.Columns.Add(col_3dlen, typeof(double));
            dt_mat_lin.Columns.Add(col_altdesc, typeof(string));
            dt_mat_lin.Columns.Add(col_xbeg, typeof(double));
            dt_mat_lin.Columns.Add(col_ybeg, typeof(double));
            dt_mat_lin.Columns.Add(col_xend, typeof(double));
            dt_mat_lin.Columns.Add(col_yend, typeof(double));
            dt_mat_lin.Columns.Add(col_block, typeof(string));
            dt_mat_lin.Columns.Add(col_blockdescr, typeof(string));
            dt_mat_lin.Columns.Add(col_mat, typeof(string));
            dt_mat_lin.Columns.Add(col_sta, typeof(string));
            dt_mat_lin.Columns.Add(col_id, typeof(string));
            dt_mat_lin.Columns.Add(col_mstartcanada, typeof(double));
            dt_mat_lin.Columns.Add(col_mendcanada, typeof(double));
            dt_mat_lin.Columns.Add(col_visibility, typeof(string));

            System.Data.DataTable dt_mat_pt = new System.Data.DataTable();
            dt_mat_pt.Columns.Add(col_MMID, typeof(string));
            dt_mat_pt.Columns.Add(col_Item_No, typeof(string));
            dt_mat_pt.Columns.Add(col_2dsta, typeof(double));
            dt_mat_pt.Columns.Add(col_3dsta, typeof(double));
            dt_mat_pt.Columns.Add(col_eqSta, typeof(double));
            dt_mat_pt.Columns.Add(col_symbol, typeof(string));
            dt_mat_pt.Columns.Add(col_altdesc, typeof(string));
            dt_mat_pt.Columns.Add(col_x, typeof(double));
            dt_mat_pt.Columns.Add(col_y, typeof(double));
            dt_mat_pt.Columns.Add(col_block, typeof(string));
            dt_mat_pt.Columns.Add(col_blockdescr, typeof(string));
            dt_mat_pt.Columns.Add(col_id, typeof(string));
            dt_mat_pt.Columns.Add(col_id2, typeof(string));
            dt_mat_pt.Columns.Add(col_cvr, typeof(string));
            dt_mat_pt.Columns.Add(col_mcanada, typeof(double));
            dt_mat_pt.Columns.Add(col_visibility, typeof(string));

            System.Data.DataTable dt_mat_extra = new System.Data.DataTable();
            dt_mat_extra.Columns.Add(col_MMID, typeof(string));
            dt_mat_extra.Columns.Add(col_Item_No, typeof(string));
            dt_mat_extra.Columns.Add(col_2dbeg, typeof(double));
            dt_mat_extra.Columns.Add(col_2dend, typeof(double));
            dt_mat_extra.Columns.Add(col_3dbeg, typeof(double));
            dt_mat_extra.Columns.Add(col_3dend, typeof(double));
            dt_mat_extra.Columns.Add(col_eqStabeg, typeof(double));
            dt_mat_extra.Columns.Add(col_eqStaend, typeof(double));
            dt_mat_extra.Columns.Add(col_2dlen, typeof(double));
            dt_mat_extra.Columns.Add(col_3dlen, typeof(double));
            dt_mat_extra.Columns.Add(col_altdesc, typeof(string));
            dt_mat_extra.Columns.Add(col_xbeg, typeof(double));
            dt_mat_extra.Columns.Add(col_ybeg, typeof(double));
            dt_mat_extra.Columns.Add(col_xend, typeof(double));
            dt_mat_extra.Columns.Add(col_yend, typeof(double));
            dt_mat_extra.Columns.Add(col_block, typeof(string));
            dt_mat_extra.Columns.Add(col_blockdescr, typeof(string));
            dt_mat_extra.Columns.Add(col_note1, typeof(string));
            dt_mat_extra.Columns.Add(col_mat, typeof(string));
            dt_mat_extra.Columns.Add(col_qty, typeof(double));
            dt_mat_extra.Columns.Add(col_mstartcanada, typeof(double));
            dt_mat_extra.Columns.Add(col_mendcanada, typeof(double));
            dt_mat_extra.Columns.Add(col_visibility, typeof(string));
            #endregion

            System.Data.DataTable dt_cl = ds_main.dt_centerline;
            System.Data.DataTable dt_top = ds_main.dt_top;

            string filename = comboBox_xl_canada.Text;

            System.Data.DataTable dt_t = new System.Data.DataTable();
            string tab_transition = "Transition Table";
            List<string> lista_mat = new List<string>();

            Microsoft.Office.Interop.Excel.Worksheet W_mat = null;

            try
            {
                if (filename.Length > 0)
                {
                    set_enable_false();
                    W_mat = Functions.Get_opened_worksheet_from_Excel_by_name(filename, tab_materials);
                    if (W_mat != null)
                    {

                        string xl_pipe_type = "E";
                        string xl_description = "F";
                        string xl_pipe_class = "G";
                        string xl_coating = "H";
                        string xl_wt = "I";
                        string xl_notes = "J";

                        string h1a = "pipe";
                        string h1b = "type";
                        string h2 = "coating";
                        string h3a = "wall";
                        string h3b = "thickness";

                        string value1 = W_mat.Range["E2"].Value2;
                        string value2 = W_mat.Range["H2"].Value2;
                        string value3 = W_mat.Range["I2"].Value2;

                        if (value1 == null) value1 = "";
                        if (value2 == null) value2 = "";
                        if (value3 == null) value3 = "";

                        value1 = value1.ToLower();
                        value2 = value2.ToLower();
                        value3 = value3.ToLower();

                        if (value1.Contains(h1a) == false || value1.Contains(h1b) == false)
                        {
                            MessageBox.Show(h1a + " " + h1b + " is not found. Check E2 on " + tab_materials + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }


                        if (value2.Contains(h2) == false)
                        {
                            MessageBox.Show(h2 + " is not found. Check H2 on " + tab_materials + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }


                        if (value3.Contains(h3a) == false || value3.Contains(h3b) == false)
                        {
                            MessageBox.Show(h3a + " " + h3b + " is not found. Check I2 on " + tab_materials + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }

                        List<string> lista_col = new List<string>();
                        List<string> lista_colxl = new List<string>();

                        lista_col.Add(col_pipe_type);
                        lista_col.Add(col_descr);
                        lista_col.Add(col_pipe_class);
                        lista_col.Add(col_coating);
                        lista_col.Add(col_notes);
                        lista_col.Add(col_wt);


                        lista_colxl.Add(xl_pipe_type);
                        lista_colxl.Add(xl_description);
                        lista_colxl.Add(xl_pipe_class);
                        lista_colxl.Add(xl_coating);
                        lista_colxl.Add(xl_notes);
                        lista_colxl.Add(xl_wt);


                        dt_materials = Functions.build_data_table_from_excel(dt_materials, W_mat, start1, end1, lista_col, lista_colxl);


                        if (dt_materials.Rows.Count == 0)
                        {
                            MessageBox.Show("No data found in" + tab_materials + "\r\noperation aborted");
                            set_enable_true();
                            return;
                        }

                        dt_t.Columns.Add("N_A", typeof(string));
                        for (int i = 0; i < dt_materials.Rows.Count; ++i)
                        {
                            string mat1 = Convert.ToString(dt_materials.Rows[i][col_pipe_type]);
                            if (lista_mat.Contains(mat1) == false)
                            {
                                lista_mat.Add(mat1);
                                dt_t.Columns.Add(mat1, typeof(string));
                                dt_t.Rows.Add();
                                dt_t.Rows[dt_t.Rows.Count - 1][0] = mat1;
                            }
                        }


                        string xl_dwg = "R";
                        string xl_m1 = "S";
                        string xl_m2 = "T";

                        lista_col = new List<string>();
                        lista_colxl = new List<string>();

                        lista_col.Add(col_dwg);
                        lista_col.Add(col_m1);
                        lista_col.Add(col_m2);

                        lista_colxl.Add(xl_dwg);
                        lista_colxl.Add(xl_m1);
                        lista_colxl.Add(xl_m2);

                        dt_counts = Functions.build_data_table_from_excel(dt_counts, W_mat, start1, end1, lista_col, lista_colxl);


                        string xl_back = "AN";
                        string xl_ahead = "AO";


                        lista_col = new List<string>();
                        lista_colxl = new List<string>();

                        lista_col.Add(col_back);
                        lista_col.Add(col_ahead);

                        lista_colxl.Add(xl_back);
                        lista_colxl.Add(xl_ahead);


                        dt_eq = Functions.build_data_table_from_excel(dt_eq, W_mat, start1, end1, lista_col, lista_colxl);

                    }
                    else
                    {
                        MessageBox.Show(tab_materials + " not found\r\noperation aborted");
                        set_enable_true();
                        return;
                    }

                    Microsoft.Office.Interop.Excel.Worksheet W11 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, tab_transition);
                    if (W11 != null)
                    {
                        string h1a = "pipe";
                        string h1b = "type";


                        string value1 = W11.Range["E2"].Value2;

                        if (value1 == null) value1 = "";

                        value1 = value1.ToLower();


                        if (value1.Contains(h1a) == false || value1.Contains(h1b) == false)
                        {
                            MessageBox.Show(h1a + " " + h1b + " is not found. Check E2 on " + tab_transition + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }

                        object[,] values1 = new object[1, 21];
                        values1 = W11.Range["F2:Z2"].Value2;
                        for (int j = 1; j <= values1.Length; ++j)
                        {
                            object Valoare1 = values1[1, j];
                            if (Valoare1 != null)
                            {
                                string mat1 = Convert.ToString(Valoare1);
                                if (lista_mat.Contains(mat1) == true)
                                {
                                    int index_row = lista_mat.IndexOf(mat1);
                                    string cur_column = Functions.get_excel_column_letter(index_row + 6);
                                    object[,] values2 = new object[21, 1];
                                    values2 = W11.Range[cur_column + "3:" + cur_column + "23"].Value2;
                                    for (int i = 1; i <= values2.Length; ++i)
                                    {
                                        object Valoare2 = values2[i, 1];
                                        if (Valoare2 != null && Convert.ToString(Valoare2) == "T")
                                        {
                                            dt_t.Rows[i - 1][mat1] = "T";
                                        }
                                    }
                                }
                            }
                            else
                            {
                                j = values1.Length;
                            }
                        }
                    }



                    Microsoft.Office.Interop.Excel.Worksheet W2 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, tab_class);
                    if (W2 != null)
                    {
                        start1 = 4;
                        string xl_pipe_type = "F";
                        string xl_wt = "G";

                        string xl_sta1 = "H";
                        string xl_sta2 = "L";
                        string xl_description = "P";
                        string xl_description1 = "I";
                        string xl_description2 = "M";

                        string h1a = "pipe";
                        string h1b = "type";
                        string h2 = "station";
                        string h3 = "w.t.";
                        string h4 = "description";

                        string value1 = W2.Range["F3"].Value2;
                        string value2 = W2.Range["H3"].Value2;
                        string value22 = W2.Range["L3"].Value2;
                        string value3 = W2.Range["G3"].Value2;
                        string value4 = W2.Range["I3"].Value2;
                        string value44 = W2.Range["M3"].Value2;

                        if (value1 == null) value1 = "";
                        if (value2 == null) value2 = "";
                        if (value22 == null) value22 = "";
                        if (value3 == null) value3 = "";
                        if (value4 == null) value4 = "";
                        if (value44 == null) value44 = "";

                        value1 = value1.ToLower();
                        value2 = value2.ToLower();
                        value22 = value22.ToLower();
                        value3 = value3.ToLower();
                        value4 = value4.ToLower();
                        value44 = value44.ToLower();


                        if (value1.Contains(h1a) == false || value1.Contains(h1b) == false)
                        {
                            MessageBox.Show(h1a + " " + h1b + " is not found. Check F3 on " + tab_class + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }

                        if (value2.Contains(h2) == false)
                        {
                            MessageBox.Show(h2 + " is not found. Check H3 on " + tab_class + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }

                        if (value22.Contains(h2) == false)
                        {
                            MessageBox.Show(h2 + " is not found. Check L3 on " + tab_class + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }

                        if (value3.Contains(h3) == false)
                        {
                            MessageBox.Show(h3 + " is not found. Check G3 on " + tab_class + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }

                        if (value4.Contains(h4) == false)
                        {
                            MessageBox.Show(h4 + " is not found. Check I3 on " + tab_class + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }


                        if (value44.Contains(h4) == false)
                        {
                            MessageBox.Show(h4 + " is not found. Check M3 on " + tab_class + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }


                        List<string> lista_col = new List<string>();
                        List<string> lista_colxl = new List<string>();

                        lista_col.Add(col_sta1);
                        lista_col.Add(col_sta2);
                        lista_col.Add(col_pipe_type);
                        lista_col.Add(col_descr);
                        lista_col.Add(col_wt);
                        lista_col.Add(col_descr1);
                        lista_col.Add(col_descr2);



                        lista_colxl.Add(xl_sta1);
                        lista_colxl.Add(xl_sta2);
                        lista_colxl.Add(xl_pipe_type);
                        lista_colxl.Add(xl_description);
                        lista_colxl.Add(xl_wt);
                        lista_colxl.Add(xl_description1);
                        lista_colxl.Add(xl_description2);


                        dt_class = Functions.build_data_table_from_excel(dt_class, W2, start1, end1, lista_col, lista_colxl);

                        dt_class = Functions.Sort_data_table(dt_class, col_sta1);
                        if (dt_class.Rows.Count == 0)
                        {
                            MessageBox.Show("No data found in" + tab_class + "\r\noperation aborted");
                            set_enable_true();
                            return;
                        }


                        Functions.Round_data_table(dt_class, 1);

                        int wrong_line = -1;
                        if (check_wall_thickness(dt_class, dt_materials, col_pipe_type, col_wt, start1, ref wrong_line) == false)
                        {
                            MessageBox.Show("error on " + tab_class + "\r\nwall thickness missmatch on row " + wrong_line.ToString() + "\r\noperation aborted");
                            set_enable_true();
                            return;
                        }
                        wrong_line = -1;

                        if (check_gaps_and_overlaps(dt_class, start1, ref wrong_line) == false)
                        {
                            MessageBox.Show("error on " + tab_class + "\r\ngaps or overlaps on row " + wrong_line.ToString() + "\r\noperation aborted");
                            set_enable_true();
                            return;
                        }

                        wrong_line = -1;

                        if (check_sta2_bigger_than_sta1(dt_class, col_sta1, col_sta2, start1, ref wrong_line) == false)
                        {
                            MessageBox.Show("error on " + tab_class + "\r\nend sta smaller than start sta on row " + wrong_line.ToString() + "\r\noperation aborted");
                            set_enable_true();
                            return;
                        }

                    }
                    else
                    {
                        MessageBox.Show(tab_class + " not found\r\noperation aborted");
                        set_enable_true();
                        return;
                    }



                    Microsoft.Office.Interop.Excel.Worksheet W_fac = Functions.Get_opened_worksheet_from_Excel_by_name(filename, tab_fab);
                    if (W_fac != null)
                    {
                        start1 = 4;
                        string xl_sta1 = "I";
                        string xl_sta2 = "L";
                        string xl_just = "R";
                        string xl_name = "E";
                        string xl_descr1 = "F";
                        string xl_descr2 = "G";
                        string xl_pipetype = "O";
                        string xl_wt = "Q";

                        string h1a = "pipe";
                        string h1b = "type";
                        string h2 = "station";
                        string h3 = "w.t.";
                        string h4 = "justification";

                        string value1 = W_fac.Range["O3"].Value2;
                        string value2 = W_fac.Range["I3"].Value2;
                        string value22 = W_fac.Range["L3"].Value2;
                        string value3 = W_fac.Range["Q3"].Value2;
                        string value4 = W_fac.Range["R3"].Value2;


                        if (value1 != null && (value1.ToLower().Contains(h1a) == false || value1.Contains(h1b) == false))
                        {
                            MessageBox.Show(h1a + " " + h1b + " is not found. Check O3 on " + tab_fab + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }
                        if (value2 != null && value2.ToLower().Contains(h2) == false)
                        {
                            MessageBox.Show(h2 + " is not found. Check I3 on " + tab_fab + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }

                        if (value22 != null && value22.ToLower().Contains(h2) == false)
                        {
                            MessageBox.Show(h2 + " is not found. Check L3 on " + tab_fab + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }


                        if (value3 != null && value3.ToLower().Contains(h3) == false)
                        {
                            MessageBox.Show(h3 + " is not found. Check Q3 on " + tab_fab + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }


                        if (value4 != null && value4.ToLower().Contains(h4) == false)
                        {
                            MessageBox.Show(h4 + " is not found. Check R3 on " + tab_fab + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }


                        List<string> lista_col = new List<string>();
                        List<string> lista_colxl = new List<string>();

                        lista_col.Add(col_sta1);
                        lista_col.Add(col_sta2);
                        lista_col.Add(col_fac_name);
                        lista_col.Add(col_descr1);
                        lista_col.Add(col_descr2);
                        lista_col.Add(col_just);
                        lista_col.Add(col_pipe_type);
                        lista_col.Add(col_wt);

                        lista_colxl.Add(xl_sta1);
                        lista_colxl.Add(xl_sta2);
                        lista_colxl.Add(xl_name);
                        lista_colxl.Add(xl_descr1);
                        lista_colxl.Add(xl_descr2);
                        lista_colxl.Add(xl_just);
                        lista_colxl.Add(xl_pipetype);
                        lista_colxl.Add(xl_wt);


                        dt_fab = Functions.build_data_table_from_excel(dt_fab, W_fac, start1, end1, lista_col, lista_colxl);

                        dt_fab = Functions.Sort_data_table(dt_fab, col_sta1);



                        Functions.Round_data_table(dt_fab, 1);

                        int wrong_line = -1;

                        if (check_overlaps(dt_fab, col_sta1, col_sta2, start1, ref wrong_line) == false)
                        {
                            MessageBox.Show("error on " + tab_fab + "\r\noverlaps on row " + wrong_line.ToString() + "\r\noperation aborted");
                            set_enable_true();
                            return;
                        }

                        wrong_line = -1;

                        if (check_sta2_bigger_than_sta1(dt_fab, col_sta1, col_sta2, start1, ref wrong_line) == false)
                        {
                            MessageBox.Show("error on " + tab_fab + "\r\nend sta smaller than start sta on row " + wrong_line.ToString() + "\r\noperation aborted");
                            set_enable_true();
                            return;
                        }
                    }


                    Microsoft.Office.Interop.Excel.Worksheet W_preexisting = Functions.Get_opened_worksheet_from_Excel_by_name(filename, tab_pre_existing_pipe);
                    if (W_preexisting != null)
                    {
                        start1 = 4;
                        string xl_sta1 = "G";
                        string xl_sta2 = "J";
                        string xl_descr = "E";
                        string xl_just = "M";
                        string xl_notes = "N";


                        List<string> lista_col = new List<string>();
                        List<string> lista_colxl = new List<string>();

                        lista_col.Add(col_start);
                        lista_col.Add(col_end);
                        lista_col.Add(col_descr);
                        lista_col.Add(col_just);
                        lista_col.Add(col_notes);

                        lista_colxl.Add(xl_sta1);
                        lista_colxl.Add(xl_sta2);
                        lista_colxl.Add(xl_descr);
                        lista_colxl.Add(xl_just);
                        lista_colxl.Add(xl_notes);

                        dt_pre_existing = Functions.build_data_table_from_excel(dt_pre_existing, W_preexisting, start1, end1, lista_col, lista_colxl);

                        dt_pre_existing = Functions.Sort_data_table(dt_pre_existing, col_start);


                        Functions.Round_data_table(dt_pre_existing, 1);

                        int wrong_line = -1;

                        if (check_overlaps(dt_pre_existing, col_start, col_end, start1, ref wrong_line) == false)
                        {
                            MessageBox.Show("error on " + tab_pre_existing_pipe + "\r\noverlaps on row " + wrong_line.ToString() + "\r\noperation aborted");
                            set_enable_true();
                            return;
                        }

                        wrong_line = -1;

                        if (check_sta2_bigger_than_sta1(dt_pre_existing, col_start, col_end, start1, ref wrong_line) == false)
                        {
                            MessageBox.Show("error on " + tab_pre_existing_pipe + "\r\nend sta smaller than start sta on row " + wrong_line.ToString() + "\r\noperation aborted");
                            set_enable_true();
                            return;
                        }



                    }


                    Microsoft.Office.Interop.Excel.Worksheet W3 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, tab_elbow);
                    if (W3 != null)
                    {
                        start1 = 4;

                        List<string> lista_col = new List<string>();
                        List<string> lista_colxl = new List<string>();

                        lista_col.Add(col_sta1);
                        lista_col.Add(col_sta2);
                        lista_col.Add(col_pipe_type);
                        lista_col.Add(col_descr);
                        lista_col.Add(col_wt);
                        lista_col.Add(col_just);
                        lista_col.Add(col_start_elbow_adjacent);
                        lista_col.Add(col_end_elbow_adjacent);
                        lista_col.Add(col_elbow_pi);
                        lista_col.Add(col_elbow_angle);
                        lista_col.Add(col_defl);
                        lista_col.Add(col_elbow_ref_dwg);
                        lista_col.Add(col_elbow_id);
                        lista_col.Add(col_bend_type);
                        lista_col.Add(col_pipe_class);
                        lista_col.Add(col_elbow_notes);

                        lista_colxl.Add(xl_elbow_sta1);
                        lista_colxl.Add(xl_elbow_sta2);
                        lista_colxl.Add(xl_elbow_pipe_type);
                        lista_colxl.Add(xl_elbow_descr);
                        lista_colxl.Add(xl_elbow_wt);
                        lista_colxl.Add(xl_elbow_just);
                        lista_colxl.Add(xl_elbow_additional_start);
                        lista_colxl.Add(xl_elbow_additional_end);
                        lista_colxl.Add(xl_elbow_pi);
                        lista_colxl.Add(xl_elbow_angle);
                        lista_colxl.Add(xl_elbow_defl);
                        lista_colxl.Add(xl_elbow_ref_dwg);
                        lista_colxl.Add(xl_elbow_id);
                        lista_colxl.Add(xl_bend_type);
                        lista_colxl.Add(xl_pipe_class);
                        lista_colxl.Add(xl_elbow_notes);

                        string h1a = "pipe";
                        string h1b = "type";
                        string h2 = "pi";
                        string h3 = "w.t.";
                        string h4 = "start";
                        string h5 = "end";

                        string value1 = W3.Range["P2"].Value2;
                        string value2 = W3.Range["L3"].Value2;
                        string value3 = W3.Range["R2"].Value2;
                        string value4 = W3.Range["M3"].Value2;
                        string value44 = W3.Range["T3"].Value2;
                        string value5 = W3.Range["N3"].Value2;
                        string value55 = W3.Range["V3"].Value2;

                        if (value1 == null) value1 = "";
                        if (value2 == null) value2 = "";
                        if (value3 == null) value3 = "";
                        if (value4 == null) value4 = "";
                        if (value44 == null) value44 = "";
                        if (value5 == null) value5 = "";
                        if (value55 == null) value55 = "";

                        value1 = value1.ToLower();
                        value2 = value2.ToLower();
                        value3 = value3.ToLower();
                        value4 = value4.ToLower();
                        value44 = value44.ToLower();
                        value5 = value5.ToLower();
                        value55 = value55.ToLower();

                        if (value1.Contains(h1a) == false || value1.Contains(h1b) == false)
                        {
                            MessageBox.Show(h1a + " " + h1b + " is not found. Check P2 on " + tab_elbow + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }
                        if (value2.Contains(h2) == false)
                        {
                            MessageBox.Show(h2 + " is not found. Check L3 on " + tab_elbow + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }

                        if (value3.Contains(h3) == false)
                        {
                            MessageBox.Show(h3 + " is not found. Check R2 on " + tab_elbow + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }

                        if (value4.Contains(h4) == false)
                        {
                            MessageBox.Show(h4 + " is not found. Check N3 on " + tab_elbow + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }

                        if (value44.Contains(h4) == false)
                        {
                            MessageBox.Show(h4 + " is not found. Check V3 on " + tab_elbow + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }

                        if (value5.Contains(h5) == false)
                        {
                            MessageBox.Show(h5 + " is not found. Check M3 on " + tab_elbow + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }


                        if (value55.Contains(h5) == false)
                        {
                            MessageBox.Show(h5 + " is not found. Check T3 on " + tab_elbow + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }

                        dt_elbow = Functions.build_data_table_from_excel(dt_elbow, W3, start1, end1, lista_col, lista_colxl);
                        dt_elbow = Functions.Sort_data_table(dt_elbow, col_sta1);
                        Functions.Round_data_table(dt_elbow, 1);

                        //Functions.Transfer_datatable_to_new_excel_spreadsheet(dt_elbow, "0.0");

                        int wrong_line = -1;
                        if (check_wall_thickness(dt_elbow, dt_materials, col_pipe_type, col_wt, start1, ref wrong_line) == false)
                        {
                            MessageBox.Show("error on " + tab_elbow + "\r\nwall thickness missmatch on row " + wrong_line.ToString() + "\r\noperation aborted");
                            set_enable_true();
                            return;
                        }

                        wrong_line = -1;
                        string comment = "no error";
                        if (check_overlaps_for_elbows(dt_elbow, start1, ref wrong_line, ref comment) == false)
                        {
                            MessageBox.Show("error on " + tab_elbow + "\r\noverlaps on row " + wrong_line.ToString() + "\r\n" + comment + "\r\noperation aborted");
                            set_enable_true();
                            return;
                        }


                        wrong_line = -1;

                        if (check_sta2_bigger_than_sta1(dt_elbow, col_sta1, col_sta2, start1, ref wrong_line) == false)
                        {
                            MessageBox.Show("error on " + tab_elbow + "\r\nend sta smaller than start sta on row " + wrong_line.ToString() + "\r\noperation aborted");
                            set_enable_true();
                            return;
                        }

                        if (dt_elbow.Rows.Count > 0)
                        {

                            if (dt_fab.Rows.Count > 0)
                            {
                                for (int i = dt_elbow.Rows.Count - 1; i >= 0; --i)
                                {
                                    double sta1 = Convert.ToDouble(dt_elbow.Rows[i][col_sta1]);
                                    double sta2 = Convert.ToDouble(dt_elbow.Rows[i][col_sta2]);

                                    for (int j = 0; j < dt_fab.Rows.Count; ++j)
                                    {
                                        if (dt_fab.Rows[j][col_sta1] != DBNull.Value && dt_fab.Rows[j][col_sta2] != DBNull.Value)
                                        {
                                            double sta_fab1 = Convert.ToDouble(dt_fab.Rows[j][col_sta1]);
                                            double sta_fab2 = Convert.ToDouble(dt_fab.Rows[j][col_sta2]);

                                            if (sta_fab2 >= sta2 && sta_fab1 <= sta1)
                                            {
                                                dt_elbow.Rows[i].Delete();
                                                j = dt_fab.Rows.Count;
                                            }
                                        }
                                    }
                                }
                            }


                            for (int i = 0; i < dt_elbow.Rows.Count; ++i)
                            {
                                if (dt_elbow.Rows[i][col_sta1] != DBNull.Value && dt_elbow.Rows[i][col_sta2] != DBNull.Value && dt_elbow.Rows[i][col_pipe_type] != DBNull.Value)
                                {
                                    double sta1 = Convert.ToDouble(dt_elbow.Rows[i][col_sta1]);
                                    double sta2 = Convert.ToDouble(dt_elbow.Rows[i][col_sta2]);
                                    string mat1 = Convert.ToString(dt_elbow.Rows[i][col_pipe_type]);

                                    for (int j = 0; j < dt_materials.Rows.Count; ++j)
                                    {
                                        if (dt_materials.Rows[j][col_pipe_type] != DBNull.Value)
                                        {
                                            string mat2 = Convert.ToString(dt_materials.Rows[j][col_pipe_type]);

                                            if (mat1 == mat2)
                                            {
                                                double existing_len = 0;



                                                double extra1 = 0;

                                                if (dt_eq != null && dt_eq.Rows.Count > 0)
                                                {
                                                    for (int k = 0; k < dt_eq.Rows.Count; ++k)
                                                    {
                                                        if (dt_eq.Rows[k][col_back] != DBNull.Value && dt_eq.Rows[k][col_ahead] != DBNull.Value)
                                                        {
                                                            double back1 = Convert.ToDouble(dt_eq.Rows[k][col_back]);
                                                            double ahead1 = Convert.ToDouble(dt_eq.Rows[k][col_ahead]);

                                                            if (sta1 < back1 && ahead1 < sta2 && sta2 > back1)
                                                            {
                                                                extra1 = extra1 + ahead1 - back1;
                                                            }

                                                        }
                                                    }
                                                }


                                                if (dt_materials.Rows[j]["elbow"] != DBNull.Value)
                                                {
                                                    existing_len = Convert.ToDouble(dt_materials.Rows[j]["elbow"]);

                                                }

                                                dt_materials.Rows[j]["elbow"] = Math.Round(Convert.ToDecimal(existing_len) + Convert.ToDecimal(sta2) - Convert.ToDecimal(sta1) - Convert.ToDecimal(extra1), 1);

                                            }

                                        }
                                    }

                                }

                            }
                        }
                    }

                    Microsoft.Office.Interop.Excel.Worksheet W4 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, tab_crossings);
                    if (W4 != null)
                    {
                        start1 = 4;
                        string xl_sta1 = "O";
                        string xl_sta2 = "R";
                        string xl_sta = "G";
                        string xl_wt = "AC";
                        string xl_just = "AD";
                        string xl_pipe_type = "AA";



                        string xl_xingid = "E";
                        string xl_xingtype = "F";
                        string xl_descr1 = "K";
                        string xl_descr2 = "L";
                        string xl_ref_dwg_id = "N";
                        string xl_min_depth = "W";
                        string xl_xing_method = "Z";
                        string xl_pipe_class = "AB";
                        string xl_agen_cvr = "Y";


                        List<string> lista_col = new List<string>();
                        List<string> lista_colxl = new List<string>();

                        lista_col.Add(col_sta);
                        lista_col.Add(col_sta1);
                        lista_col.Add(col_sta2);
                        lista_col.Add(col_pipe_type);
                        lista_col.Add(col_wt);
                        lista_col.Add(col_just);

                        lista_colxl.Add(xl_sta);
                        lista_colxl.Add(xl_sta1);
                        lista_colxl.Add(xl_sta2);
                        lista_colxl.Add(xl_pipe_type);
                        lista_colxl.Add(xl_wt);
                        lista_colxl.Add(xl_just);



                        string h1a = "pipe";
                        string h1b = "type";
                        string h2 = "station";
                        string h3 = "w.t.";


                        string value1 = W4.Range["AA2"].Value2;
                        string value2 = W4.Range["O3"].Value2;
                        string value22 = W4.Range["R3"].Value2;
                        string value3 = W4.Range["AC2"].Value2;

                        if (value1 == null) value1 = "";
                        if (value2 == null) value2 = "";
                        if (value22 == null) value22 = "";
                        if (value3 == null) value3 = "";

                        value1 = value1.ToLower();
                        value2 = value2.ToLower();
                        value22 = value22.ToLower();
                        value3 = value3.ToLower();

                        if (value1.Contains(h1a) == false || value1.Contains(h1b) == false)
                        {
                            MessageBox.Show(h1a + " " + h1b + " is not found. Check AA2 on " + tab_crossings + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }

                        if (value2.Contains(h2) == false)
                        {
                            MessageBox.Show(h2 + " is not found. Check O3 on " + tab_crossings + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }
                        if (value22.Contains(h2) == false)
                        {
                            MessageBox.Show(h2 + " is not found. Check R3 on " + tab_crossings + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }

                        if (value3.Contains(h3) == false)
                        {
                            MessageBox.Show(h3 + " is not found. Check AC3 on " + tab_crossings + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }

                        dt_crossing = Functions.build_data_table_from_excel(dt_crossing, W4,
                                                start1, end1, lista_col, lista_colxl);

                        for (int i = dt_crossing.Rows.Count - 1; i >= 0; --i)
                        {
                            if (dt_crossing.Rows[i][col_sta1] == DBNull.Value || dt_crossing.Rows[i][col_sta2] == DBNull.Value)
                            {
                                dt_crossing.Rows[i].Delete();
                            }
                        }

                        dt_crossing = Functions.Sort_data_table(dt_crossing, col_sta1);

                        Functions.Round_data_table(dt_crossing, 1);

                        //Functions.Transfer_datatable_to_new_excel_spreadsheet(dt_crossing, "0.0");

                        int wrong_line = -1;

                        if (check_wall_thickness(dt_crossing, dt_materials, col_pipe_type, col_wt, start1, ref wrong_line) == false)
                        {
                            MessageBox.Show("error on " + tab_crossings + "\r\nwall thickness missmatch on row " + wrong_line.ToString() + "\r\noperation aborted");
                            set_enable_true();
                            return;
                        }

                        wrong_line = -1;

                        if (check_overlaps(dt_crossing, col_sta1, col_sta2, start1, ref wrong_line) == false)
                        {
                            MessageBox.Show("error on " + tab_crossings + "\r\noverlaps on row " + wrong_line.ToString() + "\r\noperation aborted");
                            set_enable_true();
                            return;
                        }

                        wrong_line = -1;

                        if (check_sta2_bigger_than_sta1(dt_crossing, col_sta1, col_sta2, start1, ref wrong_line) == false)
                        {
                            MessageBox.Show("error on " + tab_crossings + "\r\nend sta smaller than start sta on row " + wrong_line.ToString() + "\r\noperation aborted");
                            set_enable_true();
                            return;
                        }

                        lista_col = new List<string>();
                        lista_colxl = new List<string>();
                        lista_col.Add(col_sta);
                        lista_col.Add(col_xingid);
                        lista_col.Add(col_xingtype);
                        lista_col.Add(col_descr1);
                        lista_col.Add(col_descr2);
                        lista_col.Add(col_ref_dwg_id);
                        lista_col.Add(col_min_depth);
                        lista_col.Add(col_xing_method);
                        lista_col.Add(col_pipe_type);
                        lista_col.Add(col_pipe_class);
                        lista_col.Add(col_wt);
                        lista_col.Add(col_just);
                        lista_col.Add(col_agen_cvr);

                        lista_colxl.Add(xl_sta);
                        lista_colxl.Add(xl_xingid);
                        lista_colxl.Add(xl_xingtype);
                        lista_colxl.Add(xl_descr1);
                        lista_colxl.Add(xl_descr2);
                        lista_colxl.Add(xl_ref_dwg_id);
                        lista_colxl.Add(xl_min_depth);
                        lista_colxl.Add(xl_xing_method);
                        lista_colxl.Add(xl_pipe_type);
                        lista_colxl.Add(xl_pipe_class);
                        lista_colxl.Add(xl_wt);
                        lista_colxl.Add(xl_just);
                        lista_colxl.Add(xl_agen_cvr);

                        dt_xing = Functions.build_data_table_from_excel(dt_xing, W4,
                                               start1, end1, lista_col, lista_colxl);
                        dt_xing = Functions.Sort_data_table(dt_xing, col_sta);
                        Functions.Round_data_table(dt_xing, 1);
                        //Functions.Transfer_datatable_to_new_excel_spreadsheet(dt_xing, "dt_xing");
                    }

                    Microsoft.Office.Interop.Excel.Worksheet W10 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, tab_geotech);
                    if (W10 != null)
                    {
                        start1 = 4;

                        string xl_notes = "Q";
                        string xl_sta1 = "R";
                        string xl_sta2 = "U";
                        string xl_pipe_type = "X";
                        string xl_wt = "Z";
                        string xl_just = "AA";


                        List<string> lista_col = new List<string>();
                        List<string> lista_colxl = new List<string>();

                        lista_col.Add(col_geotech_sta1);
                        lista_col.Add(col_geotech_descr1);
                        lista_col.Add(col_geotech_sta2);
                        lista_col.Add(col_geotech_descr2);
                        lista_col.Add(col_geotech_class);
                        lista_col.Add(col_geotech_type);
                        lista_col.Add(col_geotech_label);
                        lista_col.Add(col_notes);

                        lista_col.Add(col_sta1);
                        lista_col.Add(col_sta2);
                        lista_col.Add(col_pipe_type);
                        lista_col.Add(col_wt);
                        lista_col.Add(col_just);


                        lista_colxl.Add(xl_geotech_sta1);
                        lista_colxl.Add(xl_geotech_descr1);
                        lista_colxl.Add(xl_geotech_sta2);
                        lista_colxl.Add(xl_geotech_descr2);
                        lista_colxl.Add(xl_geotech_class);
                        lista_colxl.Add(xl_geotech_type);
                        lista_colxl.Add(xl_geotech_label);
                        lista_colxl.Add(xl_notes);

                        lista_colxl.Add(xl_sta1);
                        lista_colxl.Add(xl_sta2);
                        lista_colxl.Add(xl_pipe_type);
                        lista_colxl.Add(xl_wt);
                        lista_colxl.Add(xl_just);



                        dt_geotech = Functions.build_data_table_from_excel(dt_geotech, W10, start1, end1, lista_col, lista_colxl);
                        dt_geotech = Functions.Sort_data_table(dt_geotech, col_geotech_sta1);
                        Functions.Round_data_table(dt_geotech, 1);

                        int wrong_line = -1;

                        if (check_overlaps(dt_geotech, col_sta1, col_sta2, start1, ref wrong_line) == false)
                        {
                            MessageBox.Show("error on " + tab_geotech + "\r\noverlaps on row " + wrong_line.ToString() + "\r\noperation aborted");
                            set_enable_true();
                            return;
                        }

                        wrong_line = -1;

                        if (check_sta2_bigger_than_sta1(dt_geotech, col_sta1, col_sta2, start1, ref wrong_line) == false)
                        {
                            MessageBox.Show("error on " + tab_geotech + "\r\nend sta smaller than start sta on row " + wrong_line.ToString() + "\r\noperation aborted");
                            set_enable_true();
                            return;
                        }

                    }

                    Microsoft.Office.Interop.Excel.Worksheet W44 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, tab_stress);
                    if (W44 != null)
                    {
                        start1 = 4;
                        string xl_pipe_type = "M";
                        string xl_sta1 = "F";
                        string xl_sta2 = "I";
                        string xl_wt = "O";
                        string xl_just = "P";
                        string xl_notes = "Q";

                        List<string> lista_col = new List<string>();
                        List<string> lista_colxl = new List<string>();

                        lista_col.Add(col_sta1);
                        lista_col.Add(col_sta2);
                        lista_col.Add(col_pipe_type);
                        lista_col.Add(col_wt);
                        lista_col.Add(col_just);
                        lista_col.Add(col_notes);

                        lista_colxl.Add(xl_sta1);
                        lista_colxl.Add(xl_sta2);
                        lista_colxl.Add(xl_pipe_type);
                        lista_colxl.Add(xl_wt);
                        lista_colxl.Add(xl_just);
                        lista_colxl.Add(xl_notes);

                        string h1a = "pipe";
                        string h1b = "type";
                        string h2 = "station";
                        string h3 = "w.t.";


                        string value1 = W44.Range["M2"].Value2;
                        string value2 = W44.Range["F3"].Value2;
                        string value22 = W44.Range["I3"].Value2;
                        string value3 = W44.Range["O2"].Value2;

                        if (value1 == null) value1 = "";
                        if (value2 == null) value2 = "";
                        if (value22 == null) value22 = "";
                        if (value3 == null) value3 = "";

                        value1 = value1.ToLower();
                        value2 = value2.ToLower();
                        value22 = value22.ToLower();
                        value3 = value3.ToLower();

                        if (value1.Contains(h1a) == false || value1.Contains(h1b) == false)
                        {
                            MessageBox.Show(h1a + " " + h1b + " is not found. Check M2 on " + tab_stress + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }
                        if (value2.Contains(h2) == false)
                        {
                            MessageBox.Show(h2 + " is not found. Check F3 on " + tab_stress + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }
                        if (value22.Contains(h2) == false)
                        {
                            MessageBox.Show(h2 + " is not found. Check I3 on " + tab_stress + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }

                        if (value3.Contains(h3) == false)
                        {
                            MessageBox.Show(h3 + " is not found. Check O2 on " + tab_stress + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }


                        dt_stress = Functions.build_data_table_from_excel(dt_stress, W44, start1, end1, lista_col, lista_colxl);

                        if (dt_fab.Rows.Count > 0)
                        {
                            for (int i = dt_fab.Rows.Count - 1; i >= 0; --i)
                            {
                                if (dt_fab.Rows[i][col_pipe_type] != DBNull.Value)
                                {
                                    dt_stress.Rows.Add();
                                    dt_stress.Rows[dt_stress.Rows.Count - 1][col_sta1] = dt_fab.Rows[i][col_sta1];
                                    dt_stress.Rows[dt_stress.Rows.Count - 1][col_sta2] = dt_fab.Rows[i][col_sta2];
                                    dt_stress.Rows[dt_stress.Rows.Count - 1][col_pipe_type] = dt_fab.Rows[i][col_pipe_type];
                                    dt_stress.Rows[dt_stress.Rows.Count - 1][col_wt] = dt_fab.Rows[i][col_wt];
                                    dt_stress.Rows[dt_stress.Rows.Count - 1][col_just] = dt_fab.Rows[i][col_just];
                                    dt_stress.Rows[dt_stress.Rows.Count - 1][col_notes] = dt_fab.Rows[i][col_descr1];
                                    dt_fab.Rows[i].Delete();
                                }
                            }
                        }

                        if (dt_geotech.Rows.Count > 0)
                        {
                            for (int i = dt_geotech.Rows.Count - 1; i >= 0; --i)
                            {
                                if (dt_geotech.Rows[i][col_pipe_type] != DBNull.Value && dt_geotech.Rows[i][col_sta1] != DBNull.Value && dt_geotech.Rows[i][col_sta2] != DBNull.Value)
                                {
                                    dt_stress.Rows.Add();
                                    dt_stress.Rows[dt_stress.Rows.Count - 1][col_sta1] = dt_geotech.Rows[i][col_sta1];
                                    dt_stress.Rows[dt_stress.Rows.Count - 1][col_sta2] = dt_geotech.Rows[i][col_sta2];
                                    dt_stress.Rows[dt_stress.Rows.Count - 1][col_pipe_type] = dt_geotech.Rows[i][col_pipe_type];
                                    dt_stress.Rows[dt_stress.Rows.Count - 1][col_wt] = dt_geotech.Rows[i][col_wt];
                                    dt_stress.Rows[dt_stress.Rows.Count - 1][col_just] = dt_geotech.Rows[i][col_just];
                                    dt_stress.Rows[dt_stress.Rows.Count - 1][excel_cell1] = dt_geotech.Rows[i][excel_cell];

                                }
                            }
                        }


                        dt_stress = Functions.Sort_data_table(dt_stress, col_sta1);
                        Functions.Round_data_table(dt_stress, 1);

                        int wrong_line = -1;


                        if (check_wall_thickness(dt_stress, dt_materials, col_pipe_type, col_wt, start1, ref wrong_line) == false)
                        {
                            MessageBox.Show("error on " + tab_stress + " // " + tab_fab + " // " + tab_geotech + "\r\nwall thickness missmatch on row " + wrong_line.ToString() + "\r\noperation aborted");
                            set_enable_true();
                            return;
                        }

                        wrong_line = -1;

                        if (check_overlaps(dt_stress, col_sta1, col_sta2, start1, ref wrong_line) == false)
                        {
                            MessageBox.Show("error on " + tab_stress + " // " + tab_fab + " // " + tab_geotech + "\r\noverlaps on row " + wrong_line.ToString() + "\r\noperation aborted");
                            set_enable_true();
                            return;
                        }

                        wrong_line = -1;

                        if (check_sta2_bigger_than_sta1(dt_stress, col_sta1, col_sta2, start1, ref wrong_line) == false)
                        {
                            MessageBox.Show("error on " + tab_stress + "\r\nend sta smaller than start sta on row " + wrong_line.ToString() + "\r\noperation aborted");
                            set_enable_true();
                            return;
                        }

                    }

                    // if  there is no stress tab but is a fab or geotech tab
                    if (dt_stress.Rows.Count == 0)
                    {
                        if (dt_fab.Rows.Count > 0)
                        {

                            for (int i = dt_fab.Rows.Count - 1; i >= 0; --i)
                            {
                                if (dt_fab.Rows[i][col_pipe_type] != DBNull.Value)
                                {
                                    dt_stress.Rows.Add();
                                    dt_stress.Rows[dt_stress.Rows.Count - 1].ItemArray = dt_fab.Rows[i].ItemArray;
                                    dt_fab.Rows[i].Delete();
                                }
                            }


                            dt_stress = Functions.Sort_data_table(dt_stress, col_sta1);
                            Functions.Round_data_table(dt_stress, 1);



                            int wrong_line = -1;

                            if (check_wall_thickness(dt_stress, dt_materials, col_pipe_type, col_wt, start1, ref wrong_line) == false)
                            {
                                MessageBox.Show("error on " + tab_fab + "\r\nwall thickness missmatch on row " + wrong_line.ToString() + "\r\noperation aborted");
                                set_enable_true();
                                return;
                            }

                            wrong_line = -1;

                            if (check_overlaps(dt_stress, col_sta1, col_sta2, start1, ref wrong_line) == false)
                            {
                                MessageBox.Show("error on " + tab_fab + "\r\noverlaps on row " + wrong_line.ToString() + "\r\noperation aborted");
                                set_enable_true();
                                return;
                            }
                        }

                        if (dt_geotech.Rows.Count > 0)
                        {
                            for (int i = dt_geotech.Rows.Count - 1; i >= 0; --i)
                            {
                                if (dt_geotech.Rows[i][col_pipe_type] != DBNull.Value && dt_geotech.Rows[i][col_sta1] != DBNull.Value && dt_geotech.Rows[i][col_sta2] != DBNull.Value)
                                {
                                    dt_stress.Rows.Add();
                                    dt_stress.Rows[dt_stress.Rows.Count - 1][col_sta1] = dt_geotech.Rows[i][col_sta1];
                                    dt_stress.Rows[dt_stress.Rows.Count - 1][col_sta2] = dt_geotech.Rows[i][col_sta2];
                                    dt_stress.Rows[dt_stress.Rows.Count - 1][col_pipe_type] = dt_geotech.Rows[i][col_pipe_type];
                                    dt_stress.Rows[dt_stress.Rows.Count - 1][col_wt] = dt_geotech.Rows[i][col_wt];
                                    dt_stress.Rows[dt_stress.Rows.Count - 1][col_just] = dt_geotech.Rows[i][col_just];
                                    dt_stress.Rows[dt_stress.Rows.Count - 1][excel_cell1] = dt_geotech.Rows[i][excel_cell];

                                }
                            }

                            dt_stress = Functions.Sort_data_table(dt_stress, col_sta1);
                            Functions.Round_data_table(dt_stress, 1);



                            int wrong_line = -1;

                            if (check_wall_thickness(dt_stress, dt_materials, col_pipe_type, col_wt, start1, ref wrong_line) == false)
                            {
                                MessageBox.Show("error on " + tab_geotech + "\r\nwall thickness missmatch on row " + wrong_line.ToString() + "\r\noperation aborted");
                                set_enable_true();
                                return;
                            }

                            wrong_line = -1;

                            if (check_overlaps(dt_stress, col_sta1, col_sta2, start1, ref wrong_line) == false)
                            {
                                MessageBox.Show("error on " + tab_geotech + "\r\noverlaps on row " + wrong_line.ToString() + "\r\noperation aborted");
                                set_enable_true();
                                return;
                            }


                        }
                    }

                    Microsoft.Office.Interop.Excel.Worksheet W5 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, tab_hydro);
                    if (W5 != null)
                    {
                        start1 = 4;
                        string xl_pipe_type = "L";
                        string xl_sta1 = "F";
                        string xl_sta2 = "I";
                        string xl_wt = "N";
                        string xl_just = "O";
                        string xl_notes = "P";

                        List<string> lista_col = new List<string>();
                        List<string> lista_colxl = new List<string>();

                        lista_col.Add(col_sta1);
                        lista_col.Add(col_sta2);
                        lista_col.Add(col_pipe_type);
                        lista_col.Add(col_wt);
                        lista_col.Add(col_just);
                        lista_col.Add(col_notes);


                        lista_colxl.Add(xl_sta1);
                        lista_colxl.Add(xl_sta2);
                        lista_colxl.Add(xl_pipe_type);
                        lista_colxl.Add(xl_wt);
                        lista_colxl.Add(xl_just);
                        lista_colxl.Add(xl_notes);

                        string h1a = "pipe";
                        string h1b = "type";
                        string h2 = "station";
                        string h3 = "w.t.";


                        string value1 = W5.Range["L2"].Value2;
                        string value2 = W5.Range["F3"].Value2;
                        string value22 = W5.Range["I3"].Value2;
                        string value3 = W5.Range["N2"].Value2;

                        if (value1 == null) value1 = "";
                        if (value2 == null) value2 = "";
                        if (value22 == null) value22 = "";
                        if (value3 == null) value3 = "";

                        value1 = value1.ToLower();
                        value2 = value2.ToLower();
                        value22 = value22.ToLower();
                        value3 = value3.ToLower();

                        if (value1.Contains(h1a) == false || value1.Contains(h1b) == false)
                        {
                            MessageBox.Show(h1a + " " + h1b + " is not found. Check M2 on " + tab_hydro + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }
                        if (value2.Contains(h2) == false)
                        {
                            MessageBox.Show(h2 + " is not found. Check F3 on " + tab_hydro + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }
                        if (value22.Contains(h2) == false)
                        {
                            MessageBox.Show(h2 + " is not found. Check I3 on " + tab_hydro + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }

                        if (value3.Contains(h3) == false)
                        {
                            MessageBox.Show(h3 + " is not found. Check O2 on " + tab_hydro + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }


                        dt_hydro = Functions.build_data_table_from_excel(dt_hydro, W5,
                                                start1, end1, lista_col, lista_colxl);
                        dt_hydro = Functions.Sort_data_table(dt_hydro, col_sta1);
                        Functions.Round_data_table(dt_hydro, 1);

                        int wrong_line = -1;

                        if (check_wall_thickness(dt_hydro, dt_materials, col_pipe_type, col_wt, start1, ref wrong_line) == false)
                        {
                            MessageBox.Show("error on " + tab_hydro + "\r\nwall thickness missmatch on row " + wrong_line.ToString() + "\r\noperation aborted");
                            set_enable_true();
                            return;
                        }


                        wrong_line = -1;

                        if (check_overlaps(dt_hydro, col_sta1, col_sta2, start1, ref wrong_line) == false)
                        {
                            MessageBox.Show("error on " + tab_hydro + "\r\noverlaps on row " + wrong_line.ToString() + "\r\noperation aborted");
                            set_enable_true();
                            return;
                        }

                        wrong_line = -1;

                        if (check_sta2_bigger_than_sta1(dt_hydro, col_sta1, col_sta2, start1, ref wrong_line) == false)
                        {
                            MessageBox.Show("error on " + tab_hydro + "\r\nend sta smaller than start sta on row " + wrong_line.ToString() + "\r\noperation aborted");
                            set_enable_true();
                            return;
                        }

                    }

                    Microsoft.Office.Interop.Excel.Worksheet W6 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, tab_buoyancy);
                    if (W6 != null)
                    {
                        start1 = 3;
                        string xl_feature = "F";
                        string xl_start = "L";
                        string xl_end = "M";
                        string xl_spacing = "G";
                        string xl_count = "K";
                        string xl_descr1 = "Q";
                        string xl_descr2 = "R";
                        string xl_just = "O";
                        string xl_notes = "P";



                        List<string> lista_col = new List<string>();
                        List<string> lista_colxl = new List<string>();

                        lista_col.Add(col_start);
                        lista_col.Add(col_end);
                        lista_col.Add(col_feature);
                        lista_col.Add(col_just);
                        lista_col.Add(col_spacing);
                        lista_col.Add(col_count);
                        lista_col.Add(col_descr1);
                        lista_col.Add(col_descr2);
                        lista_col.Add(col_notes);


                        lista_colxl.Add(xl_start);
                        lista_colxl.Add(xl_end);
                        lista_colxl.Add(xl_feature);
                        lista_colxl.Add(xl_just);
                        lista_colxl.Add(xl_spacing);
                        lista_colxl.Add(xl_count);
                        lista_colxl.Add(xl_descr1);
                        lista_colxl.Add(xl_descr2);
                        lista_colxl.Add(xl_notes);


                        dt_buoy = Functions.build_data_table_from_excel(dt_buoy, W6,
                                                start1, end1, lista_col, lista_colxl);
                        dt_buoy = Functions.Sort_data_table(dt_buoy, col_start);

                        Functions.Round_data_table(dt_buoy, 1);

                        dt_long_strap = dt_buoy.Copy();

                        for (int i = dt_buoy.Rows.Count - 1; i >= 0; --i)
                        {
                            if (dt_buoy.Rows[i][col_feature] != DBNull.Value && dt_buoy.Rows[i][col_start] != DBNull.Value && dt_buoy.Rows[i][col_end] != DBNull.Value)
                            {
                                string feat1 = Convert.ToString(dt_buoy.Rows[i][col_feature]);
                                double sta1 = Convert.ToDouble(dt_buoy.Rows[i][col_start]);
                                double sta2 = Convert.ToDouble(dt_buoy.Rows[i][col_end]);

                                if (sta1 == sta2 && feat1.ToLower().Contains("long") == true && feat1.ToLower().Contains("strap") == true)
                                {
                                    dt_buoy.Rows[i].Delete();
                                }

                            }
                        }

                        for (int i = dt_long_strap.Rows.Count - 1; i >= 0; --i)
                        {
                            if (dt_long_strap.Rows[i][col_feature] != DBNull.Value && dt_long_strap.Rows[i][col_start] != DBNull.Value && dt_long_strap.Rows[i][col_end] != DBNull.Value)
                            {
                                string feat1 = Convert.ToString(dt_long_strap.Rows[i][col_feature]);
                                double sta1 = Convert.ToDouble(dt_long_strap.Rows[i][col_start]);
                                double sta2 = Convert.ToDouble(dt_long_strap.Rows[i][col_end]);

                                if (sta1 != sta2 || (feat1.ToLower().Contains("long") == false && feat1.ToLower().Contains("strap") == false))
                                {
                                    dt_long_strap.Rows[i].Delete();
                                }

                            }
                        }

                        int wrong_line = -1;

                        if (check_overlaps(dt_buoy, col_start, col_end, start1, ref wrong_line) == false)
                        {
                            MessageBox.Show("error on " + tab_buoyancy + "\r\noverlaps on row " + wrong_line.ToString() + "\r\noperation aborted");
                            set_enable_true();
                            return;
                        }
                        wrong_line = -1;

                        if (check_sta2_bigger_than_sta1(dt_buoy, col_start, col_end, start1, ref wrong_line) == false)
                        {
                            MessageBox.Show("error on " + tab_buoyancy + "\r\nend sta smaller than start sta on row " + wrong_line.ToString() + "\r\noperation aborted");
                            set_enable_true();
                            return;
                        }



                    }

                    Microsoft.Office.Interop.Excel.Worksheet W7 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, tab_cpac);
                    if (W7 != null)
                    {
                        start1 = 3;
                        string xl_sta = "E";
                        string xl_descr = "F";
                        string xl_descr2 = "G";

                        string xl_eq1 = "H";
                        string xl_eq2 = "I";
                        string xl_eq3 = "J";
                        string xl_just = "K";
                        string xl_notes = "L";


                        List<string> lista_col = new List<string>();
                        List<string> lista_colxl = new List<string>();

                        lista_col.Add(col_sta);
                        lista_col.Add(col_descr);
                        lista_col.Add(col_eq1);
                        lista_col.Add(col_just);
                        lista_col.Add(col_eq2);
                        lista_col.Add(col_eq3);
                        lista_col.Add(col_descr2);
                        lista_col.Add(col_notes);



                        lista_colxl.Add(xl_sta);
                        lista_colxl.Add(xl_descr);
                        lista_colxl.Add(xl_eq1);
                        lista_colxl.Add(xl_just);
                        lista_colxl.Add(xl_eq2);
                        lista_colxl.Add(xl_eq3);
                        lista_colxl.Add(xl_descr2);
                        lista_colxl.Add(xl_notes);


                        dt_cpac = Functions.build_data_table_from_excel(dt_cpac, W7,
                                                start1, end1, lista_col, lista_colxl);
                        dt_cpac = Functions.Sort_data_table(dt_cpac, col_sta);
                        Functions.Round_data_table(dt_cpac, 1);
                    }

                    Microsoft.Office.Interop.Excel.Worksheet W8 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, tab_es);
                    if (W8 != null)
                    {
                        start1 = 4;
                        string xl_sta = "F";
                        string xl_ditchplug = "J";
                        string xl_notes = "K";
                        string xl_spacing = "I";


                        List<string> lista_col = new List<string>();
                        List<string> lista_colxl = new List<string>();

                        lista_col.Add(col_sta);
                        lista_col.Add(col_ditchplug);
                        lista_col.Add(col_notes);
                        lista_col.Add(col_es_spacing);


                        lista_colxl.Add(xl_sta);
                        lista_colxl.Add(xl_ditchplug);
                        lista_colxl.Add(xl_notes);
                        lista_colxl.Add(xl_spacing);


                        dt_es = Functions.build_data_table_from_excel(dt_es, W8,
                                                start1, end1, lista_col, lista_colxl);
                        dt_es = Functions.Sort_data_table(dt_es, col_sta);
                        Functions.Round_data_table(dt_es, 1);
                    }

                    Microsoft.Office.Interop.Excel.Worksheet W9 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, tab_hydrotest);
                    if (W9 != null)
                    {
                        start1 = 4;

                        List<string> lista_col = new List<string>();
                        List<string> lista_colxl = new List<string>();

                        lista_col.Add(col_sta1);
                        lista_col.Add(col_descr1);
                        lista_col.Add(col_sta2);
                        lista_col.Add(col_descr2);
                        lista_col.Add(col_tst_sec);

                        lista_colxl.Add(xl_hydrotest_sta1);
                        lista_colxl.Add(xl_hydrotest_descr1);
                        lista_colxl.Add(xl_hydrotest_sta2);
                        lista_colxl.Add(xl_hydrotest_descr2);
                        lista_colxl.Add("E");

                        string h1a = "test";
                        string h1b = "section";
                        string h2 = "station";

                        string value1 = W9.Range["E2"].Value2;
                        string value2 = W9.Range["F3"].Value2;
                        string value22 = W9.Range["K3"].Value2;

                        if (value1 == null) value1 = "";
                        if (value2 == null) value2 = "";
                        if (value22 == null) value22 = "";

                        value1 = value1.ToLower();
                        value2 = value2.ToLower();
                        value22 = value22.ToLower();

                        if (value1.Contains(h1a) == false || value1.Contains(h1b) == false)
                        {
                            MessageBox.Show(h1a + " " + h1b + " is not found. Check E2 on " + tab_hydrotest + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }
                        if (value2.Contains(h2) == false)
                        {
                            MessageBox.Show(h2 + " is not found. Check F3 on " + tab_hydrotest + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }
                        if (value22.Contains(h2) == false)
                        {
                            MessageBox.Show(h2 + " is not found. Check K3 on " + tab_hydrotest + " tab\r\nOperation aborted");
                            set_enable_true();
                            return;
                        }

                        dt_hydrotest = Functions.build_data_table_from_excel(dt_hydrotest, W9,
                                                start1, end1, lista_col, lista_colxl);
                        dt_hydrotest = Functions.Sort_data_table(dt_hydrotest, col_sta1);
                        Functions.Round_data_table(dt_hydrotest, 1);
                    }

                    Microsoft.Office.Interop.Excel.Worksheet W111 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, tab_doc);
                    if (W111 != null)
                    {
                        start1 = 4;

                        string xl_just = "R";
                        string xl_notes = "S";
                        string xl_descr1 = "T";
                        string xl_descr2 = "U";


                        List<string> lista_col = new List<string>();
                        List<string> lista_colxl = new List<string>();

                        lista_col.Add(col_doc_sta1);
                        lista_col.Add(col_doc_sta2);
                        lista_col.Add(col_doc_min_cvr);
                        lista_col.Add(col_notes);
                        lista_col.Add(col_just);
                        lista_col.Add(col_descr1);
                        lista_col.Add(col_descr2);

                        lista_colxl.Add(xl_doc_sta1);
                        lista_colxl.Add(xl_doc_sta2);
                        lista_colxl.Add(xl_doc_min_cvr);
                        lista_colxl.Add(xl_notes);
                        lista_colxl.Add(xl_just);
                        lista_colxl.Add(xl_descr1);
                        lista_colxl.Add(xl_descr2);


                        dt_doc = Functions.build_data_table_from_excel(dt_doc, W111, start1, end1, lista_col, lista_colxl);
                        dt_doc = Functions.Sort_data_table(dt_doc, col_doc_sta1);
                        Functions.Round_data_table(dt_doc, 1);
                    }

                    Microsoft.Office.Interop.Excel.Worksheet W12 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, tab_muskeg);
                    if (W12 != null)
                    {
                        start1 = 4;



                        List<string> lista_col = new List<string>();
                        List<string> lista_colxl = new List<string>();

                        lista_col.Add(col_muskeg_sta1);
                        lista_col.Add(col_muskeg_descr1);
                        lista_col.Add(col_muskeg_sta2);
                        lista_col.Add(col_muskeg_descr2);
                        lista_col.Add(col_muskeg_label);


                        lista_colxl.Add(xl_muskeg_sta1);
                        lista_colxl.Add(xl_muskeg_descr1);
                        lista_colxl.Add(xl_muskeg_sta2);
                        lista_colxl.Add(xl_muskeg_descr2);
                        lista_colxl.Add(xl_muskeg_label);



                        dt_muskeg = Functions.build_data_table_from_excel(dt_muskeg, W12, start1, end1, lista_col, lista_colxl);
                        dt_muskeg = Functions.Sort_data_table(dt_muskeg, col_muskeg_sta1);
                        Functions.Round_data_table(dt_muskeg, 1);
                    }

                    for (int i = 0; i < dt_class.Rows.Count; ++i)
                    {
                        if (dt_class.Rows[i][col_sta1] != DBNull.Value && dt_class.Rows[i][col_sta2] != DBNull.Value)
                        {
                            dt1.Rows.Add();
                            dt1.Rows[dt1.Rows.Count - 1][col_sta1] = dt_class.Rows[i][col_sta1];
                            dt1.Rows[dt1.Rows.Count - 1][col_sta2] = dt_class.Rows[i][col_sta2];
                            dt1.Rows[dt1.Rows.Count - 1][col_pipe_type] = dt_class.Rows[i][col_pipe_type];
                            dt1.Rows[dt1.Rows.Count - 1][col_wt] = dt_class.Rows[i][col_wt];
                            dt1.Rows[dt1.Rows.Count - 1][col_just] = dt_class.Rows[i][col_descr];
                            dt1.Rows[dt1.Rows.Count - 1][col_src] = "class";
                        }
                    }

                    if (dt_elbow.Rows.Count > 0)
                    {
                        System.Data.DataTable dt_ell = new System.Data.DataTable();
                        dt_ell = dt_elbow.Copy();

                        insert_adjacent(ref dt1, dt_ell, dt_materials);
                    }



                    if (dt_stress.Rows.Count > 0)
                    {
                        #region split dt_stress by dt_elbow
                        if (dt_elbow.Rows.Count > 0)
                        {
                            List<int> lista_del = new List<int>();
                            for (int k = 0; k < dt_elbow.Rows.Count; ++k)
                            {
                                if (dt_elbow.Rows[k][col_sta1] != DBNull.Value && dt_elbow.Rows[k][col_sta2] != DBNull.Value)
                                {
                                    double sta1_elbow = Convert.ToDouble(dt_elbow.Rows[k][col_sta1]);
                                    double sta2_elbow = Convert.ToDouble(dt_elbow.Rows[k][col_sta2]);

                                    string just_elbow = "";

                                    if (dt_elbow.Rows[k][col_just] != DBNull.Value)
                                    {
                                        just_elbow = Convert.ToString(dt_elbow.Rows[k][col_just]);
                                    }

                                    for (int j = 0; j < dt_stress.Rows.Count; ++j)
                                    {
                                        if (dt_stress.Rows[j][col_sta1] != DBNull.Value && dt_stress.Rows[j][col_sta2] != DBNull.Value)
                                        {
                                            double sta1 = Convert.ToDouble(dt_stress.Rows[j][col_sta1]);
                                            double sta2 = Convert.ToDouble(dt_stress.Rows[j][col_sta2]);

                                            if (sta1 < sta1_elbow && sta2_elbow < sta2)
                                            {
                                                dt_stress.Rows[j][col_sta2] = sta1_elbow;

                                                dt_stress.Rows.Add();
                                                dt_stress.Rows[dt_stress.Rows.Count - 1][col_sta1] = sta2_elbow;
                                                dt_stress.Rows[dt_stress.Rows.Count - 1][col_sta2] = sta2;

                                                for (int m = 2; m < dt_stress.Columns.Count; ++m)
                                                {
                                                    dt_stress.Rows[dt_stress.Rows.Count - 1][m] = dt_stress.Rows[j][m];
                                                }

                                            }
                                            else if (sta1 == sta1_elbow && sta2_elbow < sta2)
                                            {
                                                dt_stress.Rows[j][col_sta1] = sta2_elbow;
                                            }
                                            else if (sta2 == sta2_elbow && sta1_elbow < sta2 && sta1_elbow > sta1)
                                            {
                                                dt_stress.Rows[j][col_sta2] = sta1_elbow;
                                            }

                                            else if (sta1 > sta1_elbow && sta2_elbow < sta2 && sta2_elbow > sta1)
                                            {
                                                dt_stress.Rows[j][col_sta1] = sta2_elbow;
                                            }

                                            else if (sta2 < sta2_elbow && sta1_elbow < sta2 && sta1_elbow > sta1)
                                            {
                                                dt_stress.Rows[j][col_sta2] = sta1_elbow;
                                            }

                                            else if ((sta2 == sta2_elbow && sta1_elbow == sta1) || (sta2 < sta2_elbow && sta1_elbow < sta1))
                                            {
                                                if (lista_del.Contains(j) == false) lista_del.Add(j);
                                            }

                                        }
                                    }
                                }
                            }

                            if (lista_del.Count > 0)
                            {
                                for (int i = lista_del.Count - 1; i > +0; --i)
                                {
                                    dt_stress.Rows[lista_del[i]].Delete();
                                }
                            }

                            dt_stress = Functions.Sort_data_table(dt_stress, col_sta1);




                            List<int> lista_stress = new List<int>();
                            for (int i = 0; i < dt1.Rows.Count; ++i)
                            {
                                if (dt1.Rows[i][col_sta1] != DBNull.Value && dt1.Rows[i][col_sta2] != DBNull.Value)
                                {
                                    double sta1 = Convert.ToDouble(dt1.Rows[i][col_sta1]);
                                    double sta2 = Convert.ToDouble(dt1.Rows[i][col_sta2]);
                                    if (dt1.Rows[i][col_wt] != DBNull.Value && dt1.Rows[i][col_pipe_type] != DBNull.Value)
                                    {
                                        double wt0 = Convert.ToDouble(dt1.Rows[i][col_wt]);
                                        string mat0 = Convert.ToString(dt1.Rows[i][col_pipe_type]);
                                        string coating0 = get_coating(dt_materials, mat0);
                                        bool sta1_sta2_processed = false;
                                        bool is_gap_start = true;
                                        double last_sta = -1;
                                        for (int j = 0; j < dt_stress.Rows.Count; ++j)
                                        {
                                            if (dt_stress.Rows[j][col_sta1] != DBNull.Value && dt_stress.Rows[j][col_sta2] != DBNull.Value && lista_stress.Contains(j) == false)
                                            {
                                                double sta1_stress = Convert.ToDouble(dt_stress.Rows[j][col_sta1]);
                                                double sta2_stress = Convert.ToDouble(dt_stress.Rows[j][col_sta2]);
                                                if (last_sta > -1)
                                                {
                                                    if (last_sta > sta1_stress && last_sta < sta2_stress)
                                                    {
                                                        sta1_stress = last_sta;
                                                    }
                                                }

                                                if (dt_stress.Rows[j][col_wt] != DBNull.Value && dt_stress.Rows[j][col_pipe_type] != DBNull.Value)
                                                {
                                                    double wt1 = Convert.ToDouble(dt_stress.Rows[j][col_wt]);
                                                    string mat1 = Convert.ToString(dt_stress.Rows[j][col_pipe_type]);
                                                    string coating1 = get_coating(dt_materials, mat1);

                                                    bool proceseaza = false;
                                                    if (wt0 < wt1 || (coating0 == "FBE" && coating1 != coating0))
                                                    {
                                                        proceseaza = true;
                                                    }

                                                    if (proceseaza == true)
                                                    {
                                                        #region sta1_stress >= sta1 && sta2_stress <= sta2
                                                        if (sta1_stress >= sta1 && sta2_stress <= sta2)
                                                        {
                                                            sta1_sta2_processed = true;
                                                            if (is_gap_start == true)
                                                            {
                                                                if (sta1_stress > sta1)
                                                                {
                                                                    dt2.Rows.Add();
                                                                    dt2.Rows[dt2.Rows.Count - 1][col_sta1] = sta1;
                                                                    dt2.Rows[dt2.Rows.Count - 1][col_sta2] = sta1_stress;
                                                                    dt2.Rows[dt2.Rows.Count - 1][col_pipe_type] = dt1.Rows[i][col_pipe_type];
                                                                    dt2.Rows[dt2.Rows.Count - 1][col_just] = dt1.Rows[i][col_just];
                                                                    dt2.Rows[dt2.Rows.Count - 1][col_notes] = dt1.Rows[i][col_notes];

                                                                    dt2.Rows[dt2.Rows.Count - 1][col_wt] = dt1.Rows[i][col_wt];


                                                                }
                                                                is_gap_start = false;
                                                            }

                                                            dt2.Rows.Add();
                                                            dt2.Rows[dt2.Rows.Count - 1][col_sta1] = sta1_stress;
                                                            dt2.Rows[dt2.Rows.Count - 1][col_sta2] = sta2_stress;
                                                            dt2.Rows[dt2.Rows.Count - 1][col_pipe_type] = dt_stress.Rows[j][col_pipe_type];
                                                            dt2.Rows[dt2.Rows.Count - 1][col_just] = dt_stress.Rows[j][col_just];
                                                            dt2.Rows[dt2.Rows.Count - 1][col_notes] = dt_stress.Rows[j][col_notes];
                                                            dt2.Rows[dt2.Rows.Count - 1][col_wt] = wt1;


                                                            last_sta = sta2_stress;
                                                            lista_stress.Add(j);

                                                            if (j < dt_stress.Rows.Count - 1)
                                                            {
                                                                for (int k = j + 1; k < dt_stress.Rows.Count; ++k)
                                                                {
                                                                    if (dt_stress.Rows[k][col_sta1] != DBNull.Value && dt_stress.Rows[k][col_sta2] != DBNull.Value && lista_stress.Contains(k) == false)
                                                                    {
                                                                        double sta1_stress_next = Convert.ToDouble(dt_stress.Rows[k][col_sta1]);
                                                                        double sta2_stress_next = Convert.ToDouble(dt_stress.Rows[k][col_sta2]);
                                                                        if (dt_stress.Rows[k][col_wt] != DBNull.Value)
                                                                        {
                                                                            double WT_stress_next = Convert.ToDouble(dt_stress.Rows[k][col_wt]);
                                                                            if (wt0 < WT_stress_next)
                                                                            {
                                                                                if (sta1_stress_next > last_sta && sta1_stress_next < sta2)
                                                                                {
                                                                                    dt2.Rows.Add();
                                                                                    dt2.Rows[dt2.Rows.Count - 1][col_sta1] = last_sta;
                                                                                    dt2.Rows[dt2.Rows.Count - 1][col_sta2] = sta1_stress_next;
                                                                                    dt2.Rows[dt2.Rows.Count - 1][col_pipe_type] = dt1.Rows[i][col_pipe_type];
                                                                                    dt2.Rows[dt2.Rows.Count - 1][col_just] = dt1.Rows[i][col_just];
                                                                                    dt2.Rows[dt2.Rows.Count - 1][col_notes] = dt1.Rows[i][col_notes];

                                                                                    dt2.Rows[dt2.Rows.Count - 1][col_wt] = dt1.Rows[i][col_wt];



                                                                                }
                                                                                if (sta1_stress_next < sta2)
                                                                                {
                                                                                    if (sta2_stress_next > sta2)
                                                                                    {
                                                                                        dt2.Rows.Add();
                                                                                        dt2.Rows[dt2.Rows.Count - 1][col_sta1] = sta1_stress_next;
                                                                                        dt2.Rows[dt2.Rows.Count - 1][col_sta2] = sta2;
                                                                                        dt2.Rows[dt2.Rows.Count - 1][col_pipe_type] = dt_stress.Rows[k][col_pipe_type];
                                                                                        dt2.Rows[dt2.Rows.Count - 1][col_just] = dt_stress.Rows[k][col_just];
                                                                                        dt2.Rows[dt2.Rows.Count - 1][col_notes] = dt_stress.Rows[k][col_notes];

                                                                                        dt2.Rows[dt2.Rows.Count - 1][col_wt] = WT_stress_next;

                                                                                        last_sta = sta2;
                                                                                    }
                                                                                    else
                                                                                    {
                                                                                        dt2.Rows.Add();
                                                                                        dt2.Rows[dt2.Rows.Count - 1][col_sta1] = sta1_stress_next;
                                                                                        dt2.Rows[dt2.Rows.Count - 1][col_sta2] = sta2_stress_next;
                                                                                        dt2.Rows[dt2.Rows.Count - 1][col_pipe_type] = dt_stress.Rows[k][col_pipe_type];
                                                                                        dt2.Rows[dt2.Rows.Count - 1][col_just] = dt_stress.Rows[k][col_just];
                                                                                        dt2.Rows[dt2.Rows.Count - 1][col_notes] = dt_stress.Rows[k][col_notes];

                                                                                        dt2.Rows[dt2.Rows.Count - 1][col_wt] = WT_stress_next;

                                                                                        last_sta = sta2_stress_next;
                                                                                        lista_stress.Add(k);
                                                                                    }
                                                                                }
                                                                                if (sta1_stress_next >= sta2) k = dt_stress.Rows.Count;
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }

                                                            if (last_sta > -1 && last_sta < sta2)
                                                            {
                                                                dt2.Rows.Add();
                                                                dt2.Rows[dt2.Rows.Count - 1][col_sta1] = last_sta;
                                                                dt2.Rows[dt2.Rows.Count - 1][col_sta2] = sta2;
                                                                dt2.Rows[dt2.Rows.Count - 1][col_pipe_type] = dt1.Rows[i][col_pipe_type];
                                                                dt2.Rows[dt2.Rows.Count - 1][col_just] = dt1.Rows[i][col_just];
                                                                dt2.Rows[dt2.Rows.Count - 1][col_notes] = dt1.Rows[i][col_notes];

                                                                dt2.Rows[dt2.Rows.Count - 1][col_wt] = dt1.Rows[i][col_wt];

                                                                last_sta = sta2;
                                                            }
                                                        }
                                                        #endregion

                                                        #region sta1_stress >= sta1 && sta1_stress < sta2 &&  sta2_stress > sta2
                                                        else if (sta1_stress >= sta1 && sta1_stress < sta2 && sta2_stress > sta2)
                                                        {
                                                            sta1_sta2_processed = true;

                                                            if (is_gap_start == true)
                                                            {
                                                                if (sta1_stress > sta1)
                                                                {
                                                                    dt2.Rows.Add();
                                                                    dt2.Rows[dt2.Rows.Count - 1][col_sta1] = sta1;
                                                                    dt2.Rows[dt2.Rows.Count - 1][col_sta2] = sta1_stress;
                                                                    dt2.Rows[dt2.Rows.Count - 1][col_pipe_type] = dt1.Rows[i][col_pipe_type];
                                                                    dt2.Rows[dt2.Rows.Count - 1][col_just] = dt1.Rows[i][col_just];
                                                                    dt2.Rows[dt2.Rows.Count - 1][col_notes] = dt1.Rows[i][col_notes];

                                                                    dt2.Rows[dt2.Rows.Count - 1][col_wt] = dt1.Rows[i][col_wt];



                                                                }
                                                                is_gap_start = false;
                                                            }

                                                            dt2.Rows.Add();
                                                            dt2.Rows[dt2.Rows.Count - 1][col_sta1] = sta1_stress;
                                                            dt2.Rows[dt2.Rows.Count - 1][col_sta2] = sta2;
                                                            dt2.Rows[dt2.Rows.Count - 1][col_pipe_type] = dt_stress.Rows[j][col_pipe_type];
                                                            dt2.Rows[dt2.Rows.Count - 1][col_just] = dt_stress.Rows[j][col_just];
                                                            dt2.Rows[dt2.Rows.Count - 1][col_notes] = dt_stress.Rows[j][col_notes];
                                                            dt2.Rows[dt2.Rows.Count - 1][col_wt] = wt1;


                                                            last_sta = sta2;
                                                        }
                                                        #endregion


                                                        #region sta1_stress < sta1 && sta2_stress > sta1 && sta2_stress <= sta2
                                                        else if (sta1_stress < sta1 && sta2_stress > sta1 && sta2_stress <= sta2)
                                                        {
                                                            sta1_sta2_processed = true;
                                                            dt2.Rows.Add();
                                                            dt2.Rows[dt2.Rows.Count - 1][col_sta1] = sta1;
                                                            dt2.Rows[dt2.Rows.Count - 1][col_sta2] = sta2_stress;
                                                            dt2.Rows[dt2.Rows.Count - 1][col_pipe_type] = dt_stress.Rows[j][col_pipe_type];
                                                            dt2.Rows[dt2.Rows.Count - 1][col_just] = dt_stress.Rows[j][col_just];
                                                            dt2.Rows[dt2.Rows.Count - 1][col_notes] = dt_stress.Rows[j][col_notes];
                                                            dt2.Rows[dt2.Rows.Count - 1][col_wt] = wt1;



                                                            last_sta = sta2_stress;
                                                            lista_stress.Add(j);

                                                            if (j < dt_stress.Rows.Count - 1)
                                                            {
                                                                for (int k = j + 1; k < dt_stress.Rows.Count; ++k)
                                                                {
                                                                    if (dt_stress.Rows[k][col_sta1] != DBNull.Value && dt_stress.Rows[k][col_sta2] != DBNull.Value && lista_stress.Contains(k) == false)
                                                                    {
                                                                        double sta1_stress_next = Convert.ToDouble(dt_stress.Rows[k][col_sta1]);
                                                                        double sta2_stress_next = Convert.ToDouble(dt_stress.Rows[k][col_sta2]);
                                                                        if (dt_stress.Rows[k][col_wt] != DBNull.Value)
                                                                        {
                                                                            double WT_stress_next = Convert.ToDouble(dt_stress.Rows[k][col_wt]);
                                                                            if (wt0 < WT_stress_next)
                                                                            {
                                                                                if (sta1_stress_next > last_sta && sta1_stress_next <= sta2)
                                                                                {
                                                                                    dt2.Rows.Add();
                                                                                    dt2.Rows[dt2.Rows.Count - 1][col_sta1] = last_sta;
                                                                                    dt2.Rows[dt2.Rows.Count - 1][col_sta2] = sta1_stress_next;
                                                                                    dt2.Rows[dt2.Rows.Count - 1][col_pipe_type] = dt1.Rows[i][col_pipe_type];
                                                                                    dt2.Rows[dt2.Rows.Count - 1][col_just] = dt1.Rows[i][col_just];
                                                                                    dt2.Rows[dt2.Rows.Count - 1][col_notes] = dt1.Rows[i][col_notes];

                                                                                    dt2.Rows[dt2.Rows.Count - 1][col_wt] = dt1.Rows[i][col_wt];

                                                                                    last_sta = sta1_stress_next;
                                                                                }
                                                                                if (sta1_stress_next < sta2 && sta2_stress_next <= sta2)
                                                                                {
                                                                                    dt2.Rows.Add();
                                                                                    dt2.Rows[dt2.Rows.Count - 1][col_sta1] = sta1_stress_next;
                                                                                    dt2.Rows[dt2.Rows.Count - 1][col_sta2] = sta2_stress_next;
                                                                                    dt2.Rows[dt2.Rows.Count - 1][col_pipe_type] = dt_stress.Rows[k][col_pipe_type];
                                                                                    dt2.Rows[dt2.Rows.Count - 1][col_just] = dt_stress.Rows[k][col_just];
                                                                                    dt2.Rows[dt2.Rows.Count - 1][col_notes] = dt_stress.Rows[k][col_notes];
                                                                                    dt2.Rows[dt2.Rows.Count - 1][col_wt] = WT_stress_next;



                                                                                    last_sta = sta2_stress_next;
                                                                                    lista_stress.Add(k);
                                                                                }
                                                                                if (sta1_stress_next >= sta2) k = dt_stress.Rows.Count;
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }

                                                            if (last_sta > -1 && last_sta < sta2)
                                                            {
                                                                dt2.Rows.Add();
                                                                dt2.Rows[dt2.Rows.Count - 1][col_sta1] = last_sta;
                                                                dt2.Rows[dt2.Rows.Count - 1][col_sta2] = sta2;
                                                                dt2.Rows[dt2.Rows.Count - 1][col_pipe_type] = dt1.Rows[i][col_pipe_type];
                                                                dt2.Rows[dt2.Rows.Count - 1][col_just] = dt1.Rows[i][col_just];
                                                                dt2.Rows[dt2.Rows.Count - 1][col_notes] = dt1.Rows[i][col_notes];
                                                                dt2.Rows[dt2.Rows.Count - 1][col_wt] = dt1.Rows[i][col_wt];

                                                                last_sta = sta2;
                                                            }
                                                        }
                                                        #endregion

                                                        #region sta1_stress < sta1 && sta2_stress >= sta2
                                                        else if (sta1_stress < sta1 && sta2_stress >= sta2)
                                                        {
                                                            sta1_sta2_processed = true;
                                                            dt2.Rows.Add();
                                                            dt2.Rows[dt2.Rows.Count - 1][col_sta1] = sta1;
                                                            dt2.Rows[dt2.Rows.Count - 1][col_sta2] = sta2;
                                                            dt2.Rows[dt2.Rows.Count - 1][col_pipe_type] = dt_stress.Rows[j][col_pipe_type];
                                                            dt2.Rows[dt2.Rows.Count - 1][col_just] = dt_stress.Rows[j][col_just];
                                                            dt2.Rows[dt2.Rows.Count - 1][col_notes] = dt_stress.Rows[j][col_notes];
                                                            dt2.Rows[dt2.Rows.Count - 1][col_wt] = wt1;


                                                            last_sta = sta2;
                                                            if (sta2_stress == sta2)
                                                            {
                                                                lista_stress.Add(j);
                                                            }
                                                            else
                                                            {

                                                            }
                                                        }
                                                        #endregion
                                                    }
                                                }
                                            }
                                        }

                                        if (sta1_sta2_processed == false)
                                        {
                                            dt2.Rows.Add();
                                            dt2.Rows[dt2.Rows.Count - 1][col_sta1] = dt1.Rows[i][col_sta1];
                                            dt2.Rows[dt2.Rows.Count - 1][col_sta2] = dt1.Rows[i][col_sta2];
                                            dt2.Rows[dt2.Rows.Count - 1][col_pipe_type] = dt1.Rows[i][col_pipe_type];
                                            dt2.Rows[dt2.Rows.Count - 1][col_just] = dt1.Rows[i][col_just];
                                            dt2.Rows[dt2.Rows.Count - 1][col_notes] = dt1.Rows[i][col_notes];

                                            dt2.Rows[dt2.Rows.Count - 1][col_wt] = dt1.Rows[i][col_wt];


                                        }
                                    }
                                }
                            }
                        }
                        #endregion
                    }
                    else
                    {
                        dt2 = dt1.Copy();
                    }
                    dt_compiled = Functions.Sort_data_table(dt2, col_sta1);


                    if (dt_crossing.Rows.Count > 0)
                    {
                        #region split dt_crossing by dt_elbow
                        if (dt_elbow.Rows.Count > 0)
                        {

                            for (int k = 0; k < dt_elbow.Rows.Count; ++k)
                            {
                                if (dt_elbow.Rows[k][col_sta1] != DBNull.Value && dt_elbow.Rows[k][col_sta2] != DBNull.Value)
                                {
                                    double sta1_elbow = Convert.ToDouble(dt_elbow.Rows[k][col_sta1]);
                                    double sta2_elbow = Convert.ToDouble(dt_elbow.Rows[k][col_sta2]);


                                    for (int j = 0; j < dt_crossing.Rows.Count; ++j)
                                    {
                                        if (dt_crossing.Rows[j][col_sta1] != DBNull.Value && dt_crossing.Rows[j][col_sta2] != DBNull.Value)
                                        {
                                            double sta1 = Convert.ToDouble(dt_crossing.Rows[j][col_sta1]);
                                            double sta2 = Convert.ToDouble(dt_crossing.Rows[j][col_sta2]);

                                            if (sta1 < sta1_elbow && sta2_elbow < sta2)
                                            {
                                                dt_crossing.Rows[j][col_sta2] = sta1_elbow;

                                                dt_crossing.Rows.Add();
                                                dt_crossing.Rows[dt_crossing.Rows.Count - 1][col_sta1] = sta2_elbow;
                                                dt_crossing.Rows[dt_crossing.Rows.Count - 1][col_sta2] = sta2;

                                                for (int m = 2; m < dt_crossing.Columns.Count; ++m)
                                                {
                                                    dt_crossing.Rows[dt_crossing.Rows.Count - 1][m] = dt_crossing.Rows[j][m];
                                                }

                                            }
                                            if (sta1 == sta1_elbow && sta2_elbow < sta2)
                                            {
                                                dt_crossing.Rows[j][col_sta1] = sta2_elbow;
                                            }
                                            if (sta2 == sta2_elbow && sta1_elbow < sta2 && sta1_elbow > sta1)
                                            {
                                                dt_crossing.Rows[j][col_sta2] = sta1_elbow;
                                            }

                                            if (sta1 > sta1_elbow && sta2_elbow < sta2 && sta2_elbow > sta1)
                                            {
                                                dt_crossing.Rows[j][col_sta1] = sta2_elbow;
                                            }

                                            if (sta2 < sta2_elbow && sta1_elbow < sta2 && sta1_elbow > sta1)
                                            {
                                                dt_crossing.Rows[j][col_sta2] = sta1_elbow;
                                            }

                                            if ((sta2 == sta2_elbow && sta1_elbow == sta1) || (sta2 < sta2_elbow && sta1_elbow < sta1))
                                            {
                                                dt_crossing.Rows[j].Delete();
                                            }

                                        }
                                    }
                                }
                            }



                            dt_crossing = Functions.Sort_data_table(dt_crossing, col_sta1);

                        }
                        #endregion


                        insert_dt_into_dt_compiledV2(ref dt_compiled, dt_crossing, dt_materials);


                    }


                    // dt_compiled = Functions.Sort_data_table(dt_compiled, col_sta1);

                    //Functions.Transfer_datatable_to_new_excel_spreadsheet(dt_compiled, "General");


                    if (dt_hydro.Rows.Count > 0)
                    {
                        #region split dt_hydro by dt_elbow
                        if (dt_elbow.Rows.Count > 0)
                        {

                            for (int k = 0; k < dt_elbow.Rows.Count; ++k)
                            {
                                if (dt_elbow.Rows[k][col_sta1] != DBNull.Value && dt_elbow.Rows[k][col_sta2] != DBNull.Value)
                                {
                                    double sta1_elbow = Convert.ToDouble(dt_elbow.Rows[k][col_sta1]);
                                    double sta2_elbow = Convert.ToDouble(dt_elbow.Rows[k][col_sta2]);

                                    for (int j = 0; j < dt_hydro.Rows.Count; ++j)
                                    {
                                        if (dt_hydro.Rows[j][col_sta1] != DBNull.Value && dt_hydro.Rows[j][col_sta2] != DBNull.Value)
                                        {
                                            double sta1 = Convert.ToDouble(dt_hydro.Rows[j][col_sta1]);
                                            double sta2 = Convert.ToDouble(dt_hydro.Rows[j][col_sta2]);

                                            if (sta1 < sta1_elbow && sta2_elbow < sta2)
                                            {
                                                dt_hydro.Rows[j][col_sta2] = sta1_elbow;

                                                dt_hydro.Rows.Add();
                                                dt_hydro.Rows[dt_hydro.Rows.Count - 1][col_sta1] = sta2_elbow;
                                                dt_hydro.Rows[dt_hydro.Rows.Count - 1][col_sta2] = sta2;

                                                for (int m = 2; m < dt_hydro.Columns.Count; ++m)
                                                {
                                                    dt_hydro.Rows[dt_hydro.Rows.Count - 1][m] = dt_hydro.Rows[j][m];
                                                }

                                            }
                                            else if (sta1 == sta1_elbow && sta2_elbow < sta2)
                                            {
                                                dt_hydro.Rows[j][col_sta1] = sta2_elbow;
                                            }
                                            else if (sta2 == sta2_elbow && sta1_elbow < sta2 && sta1_elbow > sta1)
                                            {
                                                dt_hydro.Rows[j][col_sta2] = sta1_elbow;
                                            }

                                            else if (sta1 > sta1_elbow && sta2_elbow < sta2 && sta2_elbow > sta1)
                                            {
                                                dt_hydro.Rows[j][col_sta1] = sta2_elbow;
                                            }

                                            else if (sta2 < sta2_elbow && sta1_elbow < sta2 && sta1_elbow > sta1)
                                            {
                                                dt_hydro.Rows[j][col_sta2] = sta1_elbow;
                                            }

                                            else if ((sta2 == sta2_elbow && sta1_elbow == sta1) || (sta2 < sta2_elbow && sta1_elbow < sta1))
                                            {
                                                dt_hydro.Rows[j].Delete();
                                            }

                                        }
                                    }
                                    dt_hydro = Functions.Sort_data_table(dt_hydro, col_sta1);

                                }
                            }
                        }
                        #endregion




                        insert_dt_into_dt_compiledV2(ref dt_compiled, dt_hydro, dt_materials);



                        // Functions.Transfer_datatable_to_new_excel_spreadsheet(dt_compiled, "General");

                    }

                    dt_compiled = Functions.Sort_data_table(dt_compiled, col_sta1);


                    //Functions.Transfer_datatable_to_new_excel_spreadsheet(dt_compiled, "General");

                    System.Data.DataTable dt_el_copy = new System.Data.DataTable();
                    dt_el_copy = dt_elbow.Copy();

                    insert_elbows(ref dt_compiled, dt_el_copy);
                    insert_fab(ref dt_compiled, dt_fab, dt_pre_existing);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                set_enable_true();
                return;
            }

            for (int i = 0; i < dt_compiled.Rows.Count; ++i)
            {
                string mat1 = Convert.ToString(dt_compiled.Rows[i][col_pipe_type]);
                for (int j = 0; j < dt_materials.Rows.Count; ++j)
                {
                    string mat2 = Convert.ToString(dt_materials.Rows[j][col_pipe_type]);
                    if (mat1 == mat2)
                    {
                        dt_compiled.Rows[i][col_coating] = dt_materials.Rows[j][col_coating];
                        if (dt_compiled.Rows[i][col_descr] == DBNull.Value) dt_compiled.Rows[i][col_descr] = dt_materials.Rows[j][col_descr];
                        dt_compiled.Rows[i][col_pipe_class] = dt_materials.Rows[j][col_pipe_class];
                        //col_pipe_class;
                    }
                }
            }



            System.Data.DataTable dt_ps = null;

            #region pipe summary
            try
            {
                string tab_summary = "Pipe-Summary";
                Microsoft.Office.Interop.Excel.Worksheet W1 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, tab_summary);
                if (W1 != null)
                {
                    W1.Range["B4:L20000"].ClearContents();
                    W1.Range["E4:L20000"].ClearFormats();


                    dt_ps = new System.Data.DataTable();
                    dt_ps.Columns.Add(col_pipe_type, typeof(string));
                    dt_ps.Columns.Add(col_pipe_class, typeof(string));
                    dt_ps.Columns.Add(col_wt, typeof(double));
                    dt_ps.Columns.Add(col_sta1, typeof(double));
                    dt_ps.Columns.Add(col_sta2, typeof(double));
                    dt_ps.Columns.Add(col_len, typeof(double));
                    dt_ps.Columns.Add(col_just, typeof(string));

                    for (int i = 0; i < dt_compiled.Rows.Count; ++i)
                    {
                        dt_ps.Rows.Add();
                        dt_ps.Rows[dt_ps.Rows.Count - 1][col_pipe_type] = dt_compiled.Rows[i][col_pipe_type];
                        dt_ps.Rows[dt_ps.Rows.Count - 1][col_pipe_class] = dt_compiled.Rows[i][col_pipe_class];
                        dt_ps.Rows[dt_ps.Rows.Count - 1][col_pipe_type] = dt_compiled.Rows[i][col_pipe_type];
                        dt_ps.Rows[dt_ps.Rows.Count - 1][col_wt] = dt_compiled.Rows[i][col_wt];
                        dt_ps.Rows[dt_ps.Rows.Count - 1][col_sta1] = dt_compiled.Rows[i][col_sta1];
                        dt_ps.Rows[dt_ps.Rows.Count - 1][col_sta2] = dt_compiled.Rows[i][col_sta2];
                        dt_ps.Rows[dt_ps.Rows.Count - 1][col_just] = dt_compiled.Rows[i][col_just];

                        if (dt_compiled.Rows[i][col_sta1] != DBNull.Value && dt_compiled.Rows[i][col_sta2] != DBNull.Value)
                        {

                            double sta1 = Convert.ToDouble(dt_compiled.Rows[i][col_sta1]);
                            double sta2 = Convert.ToDouble(dt_compiled.Rows[i][col_sta2]);
                            double extra1 = 0;


                            if (dt_eq != null && dt_eq.Rows.Count > 0)
                            {
                                for (int k = 0; k < dt_eq.Rows.Count; ++k)
                                {
                                    if (dt_eq.Rows[k][col_back] != DBNull.Value && dt_eq.Rows[k][col_ahead] != DBNull.Value)
                                    {
                                        double back1 = Convert.ToDouble(dt_eq.Rows[k][col_back]);
                                        double ahead1 = Convert.ToDouble(dt_eq.Rows[k][col_ahead]);


                                        if (sta1 < back1 && ahead1 < sta2 && sta2 > back1)
                                        {
                                            extra1 = extra1 + ahead1 - back1;
                                        }

                                        if (sta1 == back1) sta1 = ahead1;

                                    }


                                }
                            }


                            dt_ps.Rows[dt_ps.Rows.Count - 1][col_len] = Math.Round(Convert.ToDecimal(sta2) - Convert.ToDecimal(sta1) - Convert.ToDecimal(extra1), 1);

                            if (dt_counts != null && dt_counts.Rows.Count > 0)
                            {
                                bool repeat1 = true;

                                do
                                {

                                    for (int k = 0; k < dt_counts.Rows.Count; ++k)
                                    {
                                        bool assigned_this_time = false;

                                        if (dt_counts.Rows[k][col_m1] != DBNull.Value)
                                        {
                                            double m1 = Convert.ToDouble(dt_counts.Rows[k][col_m1]);

                                            if (sta1 < m1 && sta2 > m1)
                                            {

                                                dt_ps.Rows[dt_ps.Rows.Count - 1][col_sta2] = m1;

                                                #region station eq
                                                double extra2 = 0;

                                                if (dt_eq != null && dt_eq.Rows.Count > 0)
                                                {
                                                    for (int n = 0; n < dt_eq.Rows.Count; ++n)
                                                    {
                                                        if (dt_eq.Rows[n][col_back] != DBNull.Value && dt_eq.Rows[n][col_ahead] != DBNull.Value)
                                                        {
                                                            double back1 = Convert.ToDouble(dt_eq.Rows[n][col_back]);
                                                            double ahead1 = Convert.ToDouble(dt_eq.Rows[n][col_ahead]);

                                                            if (sta1 < back1 && ahead1 < m1 && sta2 > back1)
                                                            {
                                                                extra2 = extra2 + ahead1 - back1;
                                                            }


                                                        }
                                                    }
                                                }
                                                #endregion

                                                dt_ps.Rows[dt_ps.Rows.Count - 1][col_len] = Math.Round(Convert.ToDecimal(m1) - Convert.ToDecimal(sta1) - Convert.ToDecimal(extra2), 1);

                                                System.Data.DataRow row1 = dt_ps.NewRow();
                                                row1[col_pipe_type] = dt_compiled.Rows[i][col_pipe_type];
                                                row1[col_pipe_class] = dt_compiled.Rows[i][col_pipe_class];
                                                row1[col_pipe_type] = dt_compiled.Rows[i][col_pipe_type];
                                                row1[col_wt] = dt_compiled.Rows[i][col_wt];
                                                row1[col_sta1] = m1;
                                                row1[col_sta2] = dt_compiled.Rows[i][col_sta2];
                                                row1[col_just] = dt_compiled.Rows[i][col_just];



                                                #region station eq
                                                double extra3 = 0;

                                                if (dt_eq != null && dt_eq.Rows.Count > 0)
                                                {
                                                    for (int n = 0; n < dt_eq.Rows.Count; ++n)
                                                    {
                                                        if (dt_eq.Rows[n][col_back] != DBNull.Value && dt_eq.Rows[n][col_ahead] != DBNull.Value)
                                                        {
                                                            double back1 = Convert.ToDouble(dt_eq.Rows[n][col_back]);
                                                            double ahead1 = Convert.ToDouble(dt_eq.Rows[n][col_ahead]);

                                                            if (m1 < back1 && ahead1 < sta2 && sta2 > back1)
                                                            {
                                                                extra3 = extra3 + ahead1 - back1;
                                                            }


                                                        }
                                                    }
                                                }
                                                #endregion

                                                sta1 = m1;
                                                row1[col_len] = Math.Round(Convert.ToDecimal(sta2) - Convert.ToDecimal(m1) - Convert.ToDecimal(extra3), 1);
                                                dt_ps.Rows.Add(row1);
                                                repeat1 = true;
                                                assigned_this_time = true;
                                            }
                                        }
                                        if (k == dt_counts.Rows.Count - 1)
                                        {
                                            if (assigned_this_time == false && repeat1 == true) repeat1 = false;
                                        }
                                    }
                                } while (repeat1 == true);

                            }



                        }
                    }

                    int maxRows = dt_ps.Rows.Count;
                    int col_no = dt_ps.Columns.Count;
                    string last_col = Functions.get_excel_column_letter(col_no + 4);
                    Microsoft.Office.Interop.Excel.Range range1 = W1.Range["E4:" + last_col + (maxRows + 4 - 1).ToString()];

                    object[,] values1 = new object[maxRows, col_no];

                    for (int i = 0; i < maxRows; ++i)
                    {
                        for (int j = 0; j < col_no; ++j)
                        {
                            if (dt_ps.Rows[i][j] != DBNull.Value)
                            {
                                values1[i, j] = Convert.ToString(dt_ps.Rows[i][j]);
                            }
                        }
                    }
                    range1.Value2 = values1;

                    if (checkBox_prof_smys.Checked == true)
                    {
                        using (System.Data.DataTable dt_pipe_summary_simplified = dt_ps.Copy())
                        {

                            for (int i = dt_pipe_summary_simplified.Rows.Count - 1; i > 0; --i)
                            {
                                double sta2 = Convert.ToDouble(dt_pipe_summary_simplified.Rows[i][col_sta2]);
                                string mat1 = Convert.ToString(dt_pipe_summary_simplified.Rows[i][col_pipe_type]);


                                string mat2 = Convert.ToString(dt_pipe_summary_simplified.Rows[i - 1][col_pipe_type]);

                                if (mat1 == mat2)
                                {
                                    dt_pipe_summary_simplified.Rows[i - 1][col_sta2] = sta2;
                                    dt_pipe_summary_simplified.Rows[i].Delete();
                                }


                            }



                            if (dt_pipe_summary_simplified != null && dt_pipe_summary_simplified.Rows.Count > 0)
                            {
                                if (checkBox_TOP.Checked == false)
                                {

                                    #region hydro profile old code - update it!

                                    System.Data.DataTable dt_prof = new System.Data.DataTable();
                                    dt_prof.Columns.Add(col_sta, typeof(double));
                                    dt_prof.Columns.Add(col_x, typeof(double));
                                    dt_prof.Columns.Add(col_y, typeof(double));
                                    dt_prof.Columns.Add(col_z, typeof(double));
                                    dt_prof.Columns.Add(col_mat, typeof(string));
                                    dt_prof.Columns.Add(col_wt, typeof(double));
                                    dt_prof.Columns.Add(col_descr, typeof(string));

                                    if (dt_cl != null && dt_cl.Rows.Count > 1)
                                    {
                                        double max_len = -1;

                                        for (int i = 0; i < dt_cl.Rows.Count; ++i)
                                        {
                                            if (dt_cl.Rows[i][col_x] != DBNull.Value && dt_cl.Rows[i][col_y] != DBNull.Value && dt_cl.Rows[i][col_z] != DBNull.Value && dt_cl.Rows[i][Col_3DSta] != DBNull.Value)
                                            {
                                                dt_prof.Rows.Add();
                                                dt_prof.Rows[dt_prof.Rows.Count - 1][col_x] = dt_cl.Rows[i][col_x];
                                                dt_prof.Rows[dt_prof.Rows.Count - 1][col_y] = dt_cl.Rows[i][col_y];
                                                dt_prof.Rows[dt_prof.Rows.Count - 1][col_z] = dt_cl.Rows[i][col_z];
                                                dt_prof.Rows[dt_prof.Rows.Count - 1][col_sta] = dt_cl.Rows[i][Col_3DSta];
                                                dt_prof.Rows[dt_prof.Rows.Count - 1][col_descr] = "CL";

                                                if (i == dt_cl.Rows.Count - 1) max_len = Convert.ToDouble(dt_cl.Rows[i][Col_3DSta]);

                                            }
                                        }



                                        for (int i = 0; i < dt_pipe_summary_simplified.Rows.Count; ++i)
                                        {
                                            if (dt_pipe_summary_simplified.Rows[i][col_sta1] != DBNull.Value && dt_pipe_summary_simplified.Rows[i][col_sta2] != DBNull.Value)
                                            {
                                                double sta1_h = Convert.ToDouble(dt_pipe_summary_simplified.Rows[i][col_sta1]);
                                                double sta2_h = Convert.ToDouble(dt_pipe_summary_simplified.Rows[i][col_sta2]);

                                                if (sta2_h > max_len) sta2_h = max_len;

                                                for (int j = 0; j < dt_cl.Rows.Count - 1; ++j)
                                                {
                                                    if (dt_cl.Rows[j]["3DSta"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", "")) == true &&
                                                       dt_cl.Rows[j + 1]["3DSta"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", "")) == true &&
                                                       dt_cl.Rows[j]["X"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["X"])) == true &&
                                                       dt_cl.Rows[j]["Y"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Y"])) == true &&
                                                        dt_cl.Rows[j]["Z"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Z"])) == true &&
                                                        dt_cl.Rows[j]["X"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["X"])) == true &&
                                                       dt_cl.Rows[j + 1]["Y"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Y"])) == true &&
                                                        dt_cl.Rows[j + 1]["Z"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Z"])) == true)
                                                    {
                                                        double sta1 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", ""));
                                                        double sta2 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", ""));


                                                        double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                        double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);
                                                        double z1 = Convert.ToDouble(dt_cl.Rows[j]["Z"]);
                                                        double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                        double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);
                                                        double z2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Z"]);

                                                        if (i == 0 && j == 0 && sta1 == sta1_h)
                                                        {
                                                            dt_prof.Rows[0][col_mat] = dt_pipe_summary_simplified.Rows[i][col_pipe_type];
                                                            dt_prof.Rows[0][col_wt] = dt_pipe_summary_simplified.Rows[i][col_wt];
                                                            dt_prof.Rows[0][col_descr] = "H_start";
                                                        }
                                                        else
                                                        {
                                                            if (sta1_h >= sta1 && sta1_h <= sta2)
                                                            {

                                                                double x = x1 + (x2 - x1) * (sta1_h - sta1) / (sta2 - sta1);
                                                                double y = y1 + (y2 - y1) * (sta1_h - sta1) / (sta2 - sta1);
                                                                double z = z1 + (z2 - z1) * (sta1_h - sta1) / (sta2 - sta1);

                                                                System.Data.DataRow row1 = dt_prof.NewRow();
                                                                row1[col_sta] = sta1_h;
                                                                row1[col_x] = x;
                                                                row1[col_y] = y;
                                                                row1[col_z] = z;
                                                                row1[col_mat] = dt_pipe_summary_simplified.Rows[i][col_pipe_type];
                                                                row1[col_wt] = dt_pipe_summary_simplified.Rows[i][col_wt];
                                                                row1[col_descr] = "H_start";
                                                                dt_prof.Rows.Add(row1);

                                                            }

                                                            if (sta2_h >= sta1 && sta2_h <= sta2)
                                                            {

                                                                double x = x1 + (x2 - x1) * (sta2_h - sta1) / (sta2 - sta1);
                                                                double y = y1 + (y2 - y1) * (sta2_h - sta1) / (sta2 - sta1);
                                                                double z = z1 + (z2 - z1) * (sta2_h - sta1) / (sta2 - sta1);

                                                                System.Data.DataRow row1 = dt_prof.NewRow();
                                                                row1[col_sta] = sta2_h;
                                                                row1[col_x] = x;
                                                                row1[col_y] = y;
                                                                row1[col_z] = z;
                                                                row1[col_descr] = "H_end";
                                                                row1[col_mat] = dt_pipe_summary_simplified.Rows[i][col_pipe_type];
                                                                row1[col_wt] = dt_pipe_summary_simplified.Rows[i][col_wt];
                                                                dt_prof.Rows.Add(row1);

                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                        dt_prof = Functions.Sort_data_table(dt_prof, col_sta);

                                        for (int i = 0; i < dt_prof.Rows.Count; ++i)
                                        {
                                            double sta = Convert.ToDouble(dt_prof.Rows[i][col_sta]);

                                            if (dt_prof.Rows[i][col_mat] == DBNull.Value)
                                            {

                                                for (int j = 0; j < dt_pipe_summary_simplified.Rows.Count; ++j)
                                                {
                                                    if (dt_pipe_summary_simplified.Rows[j][col_sta1] != DBNull.Value &&
                                                        dt_pipe_summary_simplified.Rows[j][col_sta2] != DBNull.Value && dt_pipe_summary_simplified.Rows[j][col_pipe_type] != DBNull.Value && dt_pipe_summary_simplified.Rows[j][col_wt] != DBNull.Value)
                                                    {
                                                        double sta1_h = Convert.ToDouble(dt_pipe_summary_simplified.Rows[j][col_sta1]);
                                                        double sta2_h = Convert.ToDouble(dt_pipe_summary_simplified.Rows[j][col_sta2]);
                                                        double wt1 = Convert.ToDouble(dt_pipe_summary_simplified.Rows[j][col_wt]);
                                                        string mat1 = Convert.ToString(dt_pipe_summary_simplified.Rows[j][col_pipe_type]);

                                                        if (sta >= sta1_h && sta <= sta2_h)
                                                        {
                                                            dt_prof.Rows[i][col_mat] = dt_pipe_summary_simplified.Rows[j][col_pipe_type];
                                                            dt_prof.Rows[i][col_wt] = dt_pipe_summary_simplified.Rows[j][col_wt];

                                                        }
                                                    }
                                                }
                                            }
                                        }

                                        Functions.Transfer_datatable_to_new_excel_spreadsheet(dt_prof, "ProfHydro");
                                    }
                                    #endregion
                                }
                                if (checkBox_TOP.Checked == true)
                                {

                                    #region hydro top
                                    System.Data.DataTable dt_prof = new System.Data.DataTable();
                                    dt_prof.Columns.Add(col_sta, typeof(double));
                                    dt_prof.Columns.Add(col_z, typeof(double));
                                    dt_prof.Columns.Add(col_mat, typeof(string));
                                    dt_prof.Columns.Add(col_wt, typeof(double));
                                    dt_prof.Columns.Add(col_descr, typeof(string));

                                    dt_top = Functions.Sort_data_table(dt_top, col_3dsta);

                                    if (dt_top != null && dt_top.Rows.Count > 1)
                                    {
                                        double max_len = -1;
                                        double sta0 = Convert.ToDouble(dt_top.Rows[0][col_3dsta]) - 1;

                                        for (int i = 0; i < dt_top.Rows.Count; ++i)
                                        {
                                            if (dt_top.Rows[i][col_z] != DBNull.Value && dt_top.Rows[i][col_3dsta] != DBNull.Value)
                                            {
                                                double sta1 = Convert.ToDouble(dt_top.Rows[i][col_3dsta]);
                                                double z1 = Convert.ToDouble(dt_top.Rows[i][col_z]);
                                                if (Math.Round(sta0, 3) < Math.Round(sta1, 3))
                                                {
                                                    dt_prof.Rows.Add();
                                                    dt_prof.Rows[dt_prof.Rows.Count - 1][col_z] = z1;
                                                    dt_prof.Rows[dt_prof.Rows.Count - 1][col_sta] = sta1;
                                                    dt_prof.Rows[dt_prof.Rows.Count - 1][col_descr] = "TOP";

                                                    for (int j = 0; j < dt_pipe_summary_simplified.Rows.Count; ++j)
                                                    {
                                                        if (dt_pipe_summary_simplified.Rows[j][col_sta1] != DBNull.Value &&
                                                            dt_pipe_summary_simplified.Rows[j][col_sta2] != DBNull.Value &&
                                                            dt_pipe_summary_simplified.Rows[j][col_pipe_type] != DBNull.Value &&
                                                            dt_pipe_summary_simplified.Rows[j][col_wt] != DBNull.Value)
                                                        {
                                                            double sta1_h = Convert.ToDouble(dt_pipe_summary_simplified.Rows[j][col_sta1]);
                                                            double sta2_h = Convert.ToDouble(dt_pipe_summary_simplified.Rows[j][col_sta2]);
                                                            double wt1 = Convert.ToDouble(dt_pipe_summary_simplified.Rows[j][col_wt]);
                                                            string mat1 = Convert.ToString(dt_pipe_summary_simplified.Rows[j][col_pipe_type]);

                                                            if (sta1 >= sta1_h && sta1 <= sta2_h)
                                                            {
                                                                dt_prof.Rows[dt_prof.Rows.Count - 1][col_mat] = dt_pipe_summary_simplified.Rows[j][col_pipe_type];
                                                                dt_prof.Rows[dt_prof.Rows.Count - 1][col_wt] = dt_pipe_summary_simplified.Rows[j][col_wt];
                                                                j = dt_pipe_summary_simplified.Rows.Count;
                                                            }
                                                        }
                                                    }
                                                }
                                                if (i == dt_top.Rows.Count - 1) max_len = Convert.ToDouble(dt_top.Rows[i][col_3dsta]);
                                            }
                                        }






                                        using (System.Data.DataTable dt_temp_prof = dt_prof.Clone())
                                        {
                                            for (int i = 0; i < dt_pipe_summary_simplified.Rows.Count; ++i)
                                            {
                                                if (dt_pipe_summary_simplified.Rows[i][col_sta1] != DBNull.Value && dt_pipe_summary_simplified.Rows[i][col_sta2] != DBNull.Value)
                                                {
                                                    double sta1_h = Convert.ToDouble(dt_pipe_summary_simplified.Rows[i][col_sta1]);
                                                    double sta2_h = Convert.ToDouble(dt_pipe_summary_simplified.Rows[i][col_sta2]);
                                                    string mat = "Facility";
                                                    double wt = 100;

                                                    if (dt_pipe_summary_simplified.Rows[i][col_pipe_type] != DBNull.Value)
                                                    {
                                                        mat = Convert.ToString(dt_pipe_summary_simplified.Rows[i][col_pipe_type]);
                                                    }


                                                    if (dt_pipe_summary_simplified.Rows[i][col_wt] != DBNull.Value)
                                                    {
                                                        wt = Convert.ToDouble(dt_pipe_summary_simplified.Rows[i][col_wt]);
                                                    }

                                                    if (sta2_h > max_len) sta2_h = max_len;

                                                    bool pr1 = false;
                                                    bool pr2 = false;

                                                    for (int j = 0; j < dt_prof.Rows.Count - 1; ++j)
                                                    {



                                                        if (dt_prof.Rows[j][col_sta] != DBNull.Value &&
                                                            dt_prof.Rows[j + 1][col_sta] != DBNull.Value &&
                                                            dt_prof.Rows[j][col_z] != DBNull.Value &&
                                                            dt_prof.Rows[j + 1][col_z] != DBNull.Value)
                                                        {
                                                            double sta1 = Convert.ToDouble(dt_prof.Rows[j][col_sta]);
                                                            double sta2 = Convert.ToDouble(dt_prof.Rows[j + 1][col_sta]);


                                                            double z1 = Convert.ToDouble(dt_prof.Rows[j][col_z]);
                                                            double z2 = Convert.ToDouble(dt_prof.Rows[j + 1][col_z]);



                                                            if (pr1 == false && sta1 <= sta1_h && sta1_h <= sta2)
                                                            {
                                                                double z = z1 + (z2 - z1) * (sta1_h - sta1) / (sta2 - sta1);

                                                                dt_temp_prof.Rows.Add();
                                                                dt_temp_prof.Rows[dt_temp_prof.Rows.Count - 1][col_sta] = sta1_h;
                                                                dt_temp_prof.Rows[dt_temp_prof.Rows.Count - 1][col_z] = z;
                                                                dt_temp_prof.Rows[dt_temp_prof.Rows.Count - 1][col_descr] = "H_start";
                                                                dt_temp_prof.Rows[dt_temp_prof.Rows.Count - 1][col_mat] = mat;
                                                                dt_temp_prof.Rows[dt_temp_prof.Rows.Count - 1][col_wt] = wt;

                                                                pr1 = true;
                                                            }


                                                            if (pr2 == false && sta1 <= sta2_h && sta2_h <= sta2)
                                                            {
                                                                double z = z1 + (z2 - z1) * (sta2_h - sta1) / (sta2 - sta1);

                                                                dt_temp_prof.Rows.Add();
                                                                dt_temp_prof.Rows[dt_temp_prof.Rows.Count - 1][col_sta] = sta2_h;
                                                                dt_temp_prof.Rows[dt_temp_prof.Rows.Count - 1][col_z] = z;
                                                                dt_temp_prof.Rows[dt_temp_prof.Rows.Count - 1][col_descr] = "H_end";
                                                                dt_temp_prof.Rows[dt_temp_prof.Rows.Count - 1][col_mat] = mat;
                                                                dt_temp_prof.Rows[dt_temp_prof.Rows.Count - 1][col_wt] = wt;
                                                                pr2 = true;

                                                            }
                                                        }


                                                        if (pr1 == true && pr2 == true)
                                                        {
                                                            j = dt_prof.Rows.Count;
                                                        }

                                                    }


                                                }
                                            }

                                            dt_prof.Rows[0][col_descr] = dt_temp_prof.Rows[0][col_descr];


                                            if (dt_hydrotest.Rows.Count > 1)
                                            {
                                                using (System.Data.DataTable dt_tst_sect_temp = dt_hydrotest.Copy())
                                                {
                                                    dt_tst_sect_temp.Rows[0].Delete();
                                                    dt_tst_sect_temp.Columns.Remove(col_sta2);
                                                    dt_tst_sect_temp.Columns.Remove(col_descr1);
                                                    dt_tst_sect_temp.Columns.Remove(col_descr2);
                                                    dt_tst_sect_temp.Columns.Remove(excel_cell);
                                                    dt_tst_sect_temp.Columns.Remove("x1");
                                                    dt_tst_sect_temp.Columns.Remove("y1");
                                                    dt_tst_sect_temp.Columns.Remove("x2");
                                                    dt_tst_sect_temp.Columns.Remove("y2");
                                                    dt_tst_sect_temp.Columns.Remove("layer");
                                                    dt_tst_sect_temp.Columns.Remove("ci");
                                                    for (int i = 0; i < dt_tst_sect_temp.Rows.Count; ++i)
                                                    {
                                                        if (dt_tst_sect_temp.Rows[i][col_sta1] != DBNull.Value)
                                                        {
                                                            double tst1 = Convert.ToDouble(dt_tst_sect_temp.Rows[i][col_sta1]);
                                                            if (tst1 > max_len) tst1 = max_len;

                                                            bool pr1 = false;

                                                            for (int j = 0; j < dt_prof.Rows.Count - 1; ++j)
                                                            {
                                                                if (dt_prof.Rows[j][col_sta] != DBNull.Value &&
                                                                   dt_prof.Rows[j + 1][col_sta] != DBNull.Value &&
                                                                   dt_prof.Rows[j][col_z] != DBNull.Value &&
                                                                   dt_prof.Rows[j + 1][col_z] != DBNull.Value)
                                                                {
                                                                    double sta1 = Convert.ToDouble(dt_prof.Rows[j][col_sta]);
                                                                    double sta2 = Convert.ToDouble(dt_prof.Rows[j + 1][col_sta]);


                                                                    double z1 = Convert.ToDouble(dt_prof.Rows[j][col_z]);
                                                                    double z2 = Convert.ToDouble(dt_prof.Rows[j + 1][col_z]);


                                                                    if (pr1 == false && sta1 <= tst1 && tst1 <= sta2)
                                                                    {
                                                                        double z = z1 + (z2 - z1) * (tst1 - sta1) / (sta2 - sta1);

                                                                        System.Data.DataRow row2 = dt_prof.NewRow();
                                                                        row2[col_sta] = tst1;
                                                                        row2[col_z] = z;
                                                                        row2[col_descr] = "TST" + Convert.ToString(dt_tst_sect_temp.Rows[i][col_tst_sec]);
                                                                        row2[col_mat] = dt_prof.Rows[j][col_mat];
                                                                        row2[col_wt] = dt_prof.Rows[j][col_wt];

                                                                        dt_prof.Rows.InsertAt(row2, j + 1);
                                                                        j = dt_prof.Rows.Count;
                                                                        pr1 = true;
                                                                    }

                                                                }
                                                            }
                                                        }

                                                    }
                                                }
                                            }


                                            for (int i = 1; i < dt_temp_prof.Rows.Count - 1; i += 2)
                                            {
                                                if (dt_temp_prof.Rows[i][col_sta] != DBNull.Value &&
                                                    dt_temp_prof.Rows[i][col_z] != DBNull.Value &&
                                                    dt_temp_prof.Rows[i][col_descr] != DBNull.Value &&
                                                    dt_temp_prof.Rows[i][col_mat] != DBNull.Value &&
                                                    dt_temp_prof.Rows[i][col_wt] != DBNull.Value &&
                                                    dt_temp_prof.Rows[i + 1][col_sta] != DBNull.Value &&
                                                    dt_temp_prof.Rows[i + 1][col_z] != DBNull.Value &&
                                                    dt_temp_prof.Rows[i + 1][col_descr] != DBNull.Value &&
                                                    dt_temp_prof.Rows[i + 1][col_mat] != DBNull.Value &&
                                                    dt_temp_prof.Rows[i + 1][col_wt] != DBNull.Value)
                                                {
                                                    double sta1 = Convert.ToDouble(dt_temp_prof.Rows[i][col_sta]);
                                                    double z1 = Convert.ToDouble(dt_temp_prof.Rows[i][col_z]);
                                                    string descr1 = Convert.ToString(dt_temp_prof.Rows[i][col_descr]);
                                                    string mat1 = Convert.ToString(dt_temp_prof.Rows[i][col_mat]);
                                                    double wt1 = Convert.ToDouble(dt_temp_prof.Rows[i][col_wt]);

                                                    double sta2 = Convert.ToDouble(dt_temp_prof.Rows[i + 1][col_sta]);
                                                    double z2 = Convert.ToDouble(dt_temp_prof.Rows[i + 1][col_z]);
                                                    string descr2 = Convert.ToString(dt_temp_prof.Rows[i + 1][col_descr]);
                                                    string mat2 = Convert.ToString(dt_temp_prof.Rows[i + 1][col_mat]);
                                                    double wt2 = Convert.ToDouble(dt_temp_prof.Rows[i + 1][col_wt]);

                                                    bool ins1 = false;
                                                    bool ins2 = false;

                                                    for (int j = 0; j < dt_prof.Rows.Count; ++j)
                                                    {
                                                        if (dt_prof.Rows[j][col_sta] != DBNull.Value)
                                                        {
                                                            double sta3 = Convert.ToDouble(dt_prof.Rows[j][col_sta]);

                                                            if (ins1 == false)
                                                            {
                                                                if (sta1 < sta3)
                                                                {

                                                                    System.Data.DataRow row2 = dt_prof.NewRow();

                                                                    row2[col_sta] = sta2;
                                                                    row2[col_z] = z2;
                                                                    row2[col_descr] = descr2;
                                                                    row2[col_mat] = mat2;
                                                                    row2[col_wt] = wt2;

                                                                    dt_prof.Rows.InsertAt(row2, j);
                                                                    ins2 = true;


                                                                    System.Data.DataRow row1 = dt_prof.NewRow();

                                                                    row1[col_sta] = sta1;
                                                                    row1[col_z] = z1;
                                                                    row1[col_descr] = descr1;
                                                                    row1[col_mat] = mat1;
                                                                    row1[col_wt] = wt1;
                                                                    dt_prof.Rows.InsertAt(row1, j);
                                                                    ins1 = true;

                                                                }
                                                                else if (sta1 == sta3)
                                                                {
                                                                    dt_prof.Rows[j][col_sta] = sta1;
                                                                    dt_prof.Rows[j][col_z] = z1;
                                                                    dt_prof.Rows[j][col_descr] = descr1;
                                                                    dt_prof.Rows[j][col_mat] = mat1;
                                                                    dt_prof.Rows[j][col_wt] = wt1;
                                                                    ins1 = true;

                                                                    System.Data.DataRow row2 = dt_prof.NewRow();

                                                                    row2[col_sta] = sta2;
                                                                    row2[col_z] = z2;
                                                                    row2[col_descr] = descr2;
                                                                    row2[col_mat] = mat2;
                                                                    row2[col_wt] = wt2;

                                                                    dt_prof.Rows.InsertAt(row2, j + 1);
                                                                    ins2 = true;
                                                                }
                                                            }



                                                            if (ins1 == true && ins2 == true)
                                                            {
                                                                j = dt_prof.Rows.Count + 200;

                                                            }
                                                        }
                                                    }




                                                }
                                            }






                                            dt_prof.Rows[dt_prof.Rows.Count - 1][col_descr] = dt_temp_prof.Rows[dt_temp_prof.Rows.Count - 1][col_descr];
                                        }

                                        Microsoft.Office.Interop.Excel.Worksheet W_hydro = Functions.Transfer_datatable_to_new_excel_spreadsheet(dt_prof, "ProfHydroTOP");
                                    }
                                    #endregion
                                }

                            }
                        }

                    }


                    System.Data.DataTable dt88 = new System.Data.DataTable();
                    dt88.Columns.Add("BY", typeof(string));
                    dt88.Columns.Add("EMPTY", typeof(string));
                    dt88.Columns.Add("REVDATE", typeof(string));
                    for (int i = 0; i < dt_ps.Rows.Count; ++i)
                    {
                        dt88.Rows.Add();
                        dt88.Rows[i]["BY"] = Functions.get_initial();
                        dt88.Rows[i]["REVDATE"] = DateTime.Now.Month + "/" + DateTime.Now.Day + "/" + DateTime.Now.Year;
                    }
                    range1 = W1.Range["B4:D" + (maxRows + 4 - 1).ToString()];

                    values1 = new object[maxRows, 3];

                    for (int i = 0; i < maxRows; ++i)
                    {
                        for (int j = 0; j < 3; ++j)
                        {
                            if (dt88.Rows[i][j] != DBNull.Value)
                            {
                                values1[i, j] = Convert.ToString(dt88.Rows[i][j]);
                            }
                        }
                    }

                    range1.Value2 = values1;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            #endregion


            #region Material counts
            if (dt_compiled.Rows.Count > 0)
            {
                for (int i = 0; i < dt_compiled.Rows.Count; ++i)
                {
                    if (dt_compiled.Rows[i][col_sta1] != DBNull.Value && dt_compiled.Rows[i][col_sta2] != DBNull.Value && dt_compiled.Rows[i][col_pipe_type] != DBNull.Value)
                    {
                        double sta1 = Convert.ToDouble(dt_compiled.Rows[i][col_sta1]);
                        double sta2 = Convert.ToDouble(dt_compiled.Rows[i][col_sta2]);
                        string mat1 = Convert.ToString(dt_compiled.Rows[i][col_pipe_type]);
                        string just1 = "";

                        if (dt_compiled.Rows[i][col_just] != DBNull.Value)
                        {
                            just1 = Convert.ToString(dt_compiled.Rows[i][col_just]);
                        }

                        if (just1.Contains("**elbow") == false)
                        {
                            for (int j = 0; j < dt_materials.Rows.Count; ++j)
                            {
                                if (dt_materials.Rows[j][col_pipe_type] != DBNull.Value)
                                {
                                    string mat2 = Convert.ToString(dt_materials.Rows[j][col_pipe_type]);
                                    if (mat1 == mat2)
                                    {

                                        double extra1 = 0;

                                        if (dt_eq != null && dt_eq.Rows.Count > 0)
                                        {
                                            for (int k = 0; k < dt_eq.Rows.Count; ++k)
                                            {
                                                if (dt_eq.Rows[k][col_back] != DBNull.Value && dt_eq.Rows[k][col_ahead] != DBNull.Value)
                                                {
                                                    double back1 = Convert.ToDouble(dt_eq.Rows[k][col_back]);
                                                    double ahead1 = Convert.ToDouble(dt_eq.Rows[k][col_ahead]);

                                                    if (sta1 < back1 && ahead1 < sta2 && sta2 > back1)
                                                    {
                                                        extra1 = extra1 + ahead1 - back1;
                                                    }

                                                }
                                            }
                                        }


                                        double existing_len = 0;
                                        if (dt_materials.Rows[j]["pipe"] != DBNull.Value)
                                        {
                                            existing_len = Convert.ToDouble(dt_materials.Rows[j]["pipe"]);
                                        }
                                        dt_materials.Rows[j]["pipe"] = Math.Round(Convert.ToDecimal(existing_len) + Convert.ToDecimal(sta2) - Convert.ToDecimal(sta1) - Convert.ToDecimal(extra1), 1);
                                    }
                                }
                            }
                        }
                    }
                }


                if (dt_counts != null && dt_counts.Rows.Count > 0 && dt_ps != null && dt_ps.Rows.Count > 0 && dt_materials != null && dt_materials.Rows.Count > 0)
                {

                    for (int j = 0; j < dt_materials.Rows.Count; ++j)
                    {
                        if (dt_materials.Rows[j][col_pipe_type] != DBNull.Value)
                        {
                            string mat2 = Convert.ToString(dt_materials.Rows[j][col_pipe_type]);
                            if (dt_counts.Columns.Contains(mat2.ToUpper()) == false)
                            {
                                dt_counts.Columns.Add(mat2, typeof(double));
                            }

                        }
                    }

                    dt_counts.Columns.Add("ELBOW", typeof(double));



                    for (int i = 0; i < dt_ps.Rows.Count; ++i)
                    {
                        if (dt_ps.Rows[i][col_sta1] != DBNull.Value && dt_ps.Rows[i][col_sta2] != DBNull.Value && dt_ps.Rows[i][col_pipe_type] != DBNull.Value)
                        {
                            double sta1 = Convert.ToDouble(dt_ps.Rows[i][col_sta1]);
                            double sta2 = Convert.ToDouble(dt_ps.Rows[i][col_sta2]);
                            string mat1 = Convert.ToString(dt_ps.Rows[i][col_pipe_type]);
                            string just1 = "";
                            if (dt_ps.Rows[i][col_just] != DBNull.Value)
                            {
                                just1 = Convert.ToString(dt_ps.Rows[i][col_just]);
                            }
                            if (just1.Contains("**elbow") == true)
                            {
                                mat1 = "ELBOW";
                            }

                            for (int j = 0; j < dt_counts.Rows.Count; ++j)
                            {
                                if (dt_counts.Rows[j][col_m1] != DBNull.Value && dt_counts.Rows[j][col_m2] != DBNull.Value)
                                {
                                    double m1 = Convert.ToDouble(dt_counts.Rows[j][col_m1]);
                                    double m2 = Convert.ToDouble(dt_counts.Rows[j][col_m2]);

                                    double existing_len = 0;
                                    if (dt_counts.Rows[j][mat1] != DBNull.Value)
                                    {
                                        existing_len = Convert.ToDouble(dt_counts.Rows[j][mat1]);
                                    }

                                    if (m1 <= sta1 && m2 >= sta2)
                                    {
                                        double extra1 = 0;

                                        if (dt_eq != null && dt_eq.Rows.Count > 0)
                                        {
                                            for (int k = 0; k < dt_eq.Rows.Count; ++k)
                                            {
                                                if (dt_eq.Rows[k][col_back] != DBNull.Value && dt_eq.Rows[k][col_ahead] != DBNull.Value)
                                                {
                                                    double back1 = Convert.ToDouble(dt_eq.Rows[k][col_back]);
                                                    double ahead1 = Convert.ToDouble(dt_eq.Rows[k][col_ahead]);

                                                    if (sta1 < back1 && ahead1 < sta2 && sta2 > back1)
                                                    {
                                                        extra1 = extra1 + ahead1 - back1;
                                                    }

                                                    if (sta1 == back1) sta1 = ahead1;

                                                }
                                            }
                                        }

                                        double new_len = Convert.ToDouble(Convert.ToDecimal(existing_len) + Convert.ToDecimal(sta2) - Convert.ToDecimal(sta1) - Convert.ToDecimal(extra1));
                                        dt_counts.Rows[j][mat1] = Math.Round(new_len, 1);
                                        j = dt_counts.Rows.Count;
                                    }
                                    else if (m1 <= sta1 && m2 < sta2 && m2 > sta1)
                                    {

                                        double extra1 = 0;

                                        if (dt_eq != null && dt_eq.Rows.Count > 0)
                                        {
                                            for (int k = 0; k < dt_eq.Rows.Count; ++k)
                                            {
                                                if (dt_eq.Rows[k][col_back] != DBNull.Value && dt_eq.Rows[k][col_ahead] != DBNull.Value)
                                                {
                                                    double back1 = Convert.ToDouble(dt_eq.Rows[k][col_back]);
                                                    double ahead1 = Convert.ToDouble(dt_eq.Rows[k][col_ahead]);

                                                    if (sta1 < back1 && ahead1 < m2 && sta2 > back1)
                                                    {
                                                        extra1 = extra1 + ahead1 - back1;
                                                    }

                                                    if (sta1 == back1) sta1 = ahead1;

                                                }
                                            }
                                        }

                                        double new_len = Math.Round(Convert.ToDouble(Convert.ToDecimal(existing_len) + Convert.ToDecimal(m2) - Convert.ToDecimal(sta1) - Convert.ToDecimal(extra1)), 1);
                                        dt_counts.Rows[j][mat1] = new_len;
                                        sta1 = m2;
                                    }
                                }
                            }
                        }
                    }
                    if (checkBox_mat_counts.Checked == true) Functions.Transfer_datatable_to_new_excel_spreadsheet(dt_counts, "MatCounts");
                }

            }

            try
            {
                if (W_mat != null)
                {
                    W_mat.Range["K2:N20000"].ClearContents();

                    W_mat.Range["K2"].Value2 = "Pipe Count [m]";
                    W_mat.Range["L2"].Value2 = "Elbow Count [m]";
                    W_mat.Range["M2"].Value2 = "By";
                    W_mat.Range["N2"].Value2 = "Date";
                    W_mat.Range["K:N"].ColumnWidth = 11.71;
                    W_mat.Range["K:L"].NumberFormat = "0.0";
                    W_mat.Range["M:N"].NumberFormat = "General";
                    W_mat.Range["K2:L2"].WrapText = true;

                    for (int i = 0; i < dt_materials.Rows.Count; ++i)
                    {
                        double pipe1 = 0;
                        double elbow1 = 0;
                        if (dt_materials.Rows[i]["pipe"] != DBNull.Value)
                        {
                            pipe1 = Convert.ToDouble(dt_materials.Rows[i]["pipe"]);
                        }
                        if (dt_materials.Rows[i]["elbow"] != DBNull.Value)
                        {
                            elbow1 = Convert.ToDouble(dt_materials.Rows[i]["elbow"]);
                        }

                        if (pipe1 > 0) W_mat.Range["K" + (3 + i).ToString()].Value2 = pipe1;
                        if (elbow1 > 0) W_mat.Range["L" + (3 + i).ToString()].Value2 = elbow1;
                        W_mat.Range["M" + (3 + i).ToString()].Value2 = Functions.get_initial();
                        W_mat.Range["N" + (3 + i).ToString()].Value2 = DateTime.Now.Month + "/" + DateTime.Now.Day + "/" + DateTime.Now.Year;

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            #endregion


            if (dt_buoy.Rows.Count > 0)
            {
                System.Data.DataTable dt_count_buoy = new System.Data.DataTable();

                for (int i = 0; i < dt_buoy.Rows.Count; ++i)
                {
                    if (dt_buoy.Rows[i][col_feature] != DBNull.Value)
                    {
                        string type1 = Convert.ToString(dt_buoy.Rows[i][col_feature]);

                        if (dt_buoy.Rows[i][col_count] != DBNull.Value)
                        {
                            if (dt_count_buoy.Columns.Contains(type1) == false)
                            {
                                dt_count_buoy.Columns.Add(type1, typeof(int));
                            }
                        }
                        else
                        {
                            if (dt_count_buoy.Columns.Contains(type1) == false)
                            {
                                dt_count_buoy.Columns.Add(type1, typeof(double));
                            }
                        }
                    }


                }

                dt_count_buoy.Rows.Add();
                for (int i = 0; i < dt_buoy.Rows.Count; ++i)
                {
                    if (dt_buoy.Rows[i][col_feature] != DBNull.Value)
                    {
                        string type1 = Convert.ToString(dt_buoy.Rows[i][col_feature]);

                        if (dt_buoy.Rows[i][col_count] != DBNull.Value)
                        {

                            int count1 = Convert.ToInt32(dt_buoy.Rows[i][col_count]);
                            int existing1 = 0;
                            if (dt_count_buoy.Rows[0][type1] != DBNull.Value)
                            {
                                existing1 = Convert.ToInt32(dt_count_buoy.Rows[0][type1]);
                            }
                            dt_count_buoy.Rows[0][type1] = existing1 + count1;
                        }
                        else
                        {

                            if (dt_buoy.Rows[i][col_start] != DBNull.Value && dt_buoy.Rows[i][col_end] != DBNull.Value)
                            {
                                double feat_start = Convert.ToDouble(dt_buoy.Rows[i][col_start]);
                                double feat_end = Convert.ToDouble(dt_buoy.Rows[i][col_end]);
                                double existing1 = 0;
                                if (dt_count_buoy.Rows[0][type1] != DBNull.Value)
                                {
                                    existing1 = Convert.ToDouble(dt_count_buoy.Rows[0][type1]);
                                }
                                dt_count_buoy.Rows[0][type1] = existing1 + feat_end - feat_start;
                            }
                        }
                    }


                }


                try
                {
                    if (W_mat != null)
                    {

                        int mat_count = dt_materials.Rows.Count + 4;

                        W_mat.Range["E" + mat_count.ToString() + ":J20000"].ClearContents();

                        for (int i = 0; i < dt_count_buoy.Columns.Count; ++i)
                        {
                            W_mat.Range["F" + (mat_count + i).ToString()].Value2 = dt_count_buoy.Columns[i].ColumnName;
                            W_mat.Range["K" + (mat_count + i).ToString()].Value2 = dt_count_buoy.Rows[0][i];
                        }

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }


            }

            System.Data.DataTable dt_transition = create_dt_transition(dt_compiled, dt_t);


            if (checkBox_transition_table.Checked == true && dt_cl != null && dt_cl.Rows.Count > 1 && dt_transition != null && dt_transition.Rows.Count > 0)
            {
                System.Data.DataTable dt9 = new System.Data.DataTable();
                dt9.Columns.Add("Transition Station", typeof(double));
                dt9.Columns.Add("Closest PI Station", typeof(double));
                dt9.Columns.Add("PI Angle", typeof(string));
                dt9.Columns.Add("PI Distance", typeof(double));
                dt9.Columns.Add("Elbow Start", typeof(double));
                dt9.Columns.Add("Elbow End", typeof(double));
                dt9.Columns.Add("Elbow Angle", typeof(double));


                for (int i = 0; i < dt_transition.Rows.Count; ++i)
                {
                    if (dt_transition.Rows[i][col_sta] != DBNull.Value)
                    {
                        double sta1 = Convert.ToDouble(dt_transition.Rows[i][col_sta]);
                        dt9.Rows.Add();
                        dt9.Rows[dt9.Rows.Count - 1][0] = sta1;

                        if (dt_cl.Rows[dt_cl.Rows.Count - 1]["3DSta"] != DBNull.Value)
                        {
                            double min_dist = Convert.ToDouble(dt_cl.Rows[dt_cl.Rows.Count - 1]["3DSta"]);
                            for (int j = 0; j < dt_cl.Rows.Count; ++j)
                            {
                                if (dt_cl.Rows[j]["DeflAng"] != DBNull.Value && dt_cl.Rows[j]["3DSta"] != DBNull.Value)
                                {
                                    double ang2 = Convert.ToDouble(dt_cl.Rows[j]["DeflAng"]);
                                    double sta2 = Convert.ToDouble(dt_cl.Rows[j]["3DSta"]);
                                    if (ang2 >= 1)
                                    {
                                        if (Math.Abs(sta1 - sta2) < min_dist)
                                        {
                                            bool is_elbow = false;
                                            if (dt_elbow != null && dt_elbow.Rows.Count > 0)
                                            {
                                                for (int k = 0; k < dt_elbow.Rows.Count; ++k)
                                                {
                                                    if (dt_elbow.Rows[k][col_elbow_pi] != DBNull.Value)
                                                    {
                                                        double sta3 = Convert.ToDouble(dt_elbow.Rows[k][col_elbow_pi]);
                                                        if (Math.Round(sta3, 1) == Math.Round(sta2, 1))
                                                        {
                                                            if (Math.Abs(sta1 - sta2) <= 15)
                                                            {
                                                                dt9.Rows[dt9.Rows.Count - 1][4] = dt_elbow.Rows[k][col_sta1];
                                                                dt9.Rows[dt9.Rows.Count - 1][5] = dt_elbow.Rows[k][col_sta2];
                                                                dt9.Rows[dt9.Rows.Count - 1][6] = dt_elbow.Rows[k][col_elbow_angle];
                                                            }

                                                            is_elbow = true;
                                                        }
                                                    }
                                                }
                                            }

                                            if (is_elbow == false)
                                            {
                                                min_dist = Math.Abs(sta1 - sta2);
                                                dt9.Rows[dt9.Rows.Count - 1][1] = sta2;
                                                dt9.Rows[dt9.Rows.Count - 1][2] = dt_cl.Rows[j]["DeflAngDMS"];
                                                dt9.Rows[dt9.Rows.Count - 1][3] = min_dist;
                                            }

                                        }
                                    }

                                }
                            }
                        }


                    }
                }



                if (dt_elbow != null && dt_elbow.Rows.Count > 0)
                {

                    System.Data.DataTable dt10 = new System.Data.DataTable();


                    dt10.Columns.Add("Elbow Start", typeof(double));
                    dt10.Columns.Add("Elbow PI", typeof(double));
                    dt10.Columns.Add("Elbow End", typeof(double));
                    dt10.Columns.Add("Elbow Angle", typeof(double));
                    dt10.Columns.Add("Closest Transition from Start", typeof(double));
                    dt10.Columns.Add("Closest Transition from End", typeof(double));
                    dt10.Columns.Add("Distance from Start", typeof(double));
                    dt10.Columns.Add("Distance from End", typeof(double));


                    for (int k = 0; k < dt_elbow.Rows.Count; ++k)
                    {
                        if (dt_elbow.Rows[k][col_elbow_pi] != DBNull.Value && dt_elbow.Rows[k][col_sta1] != DBNull.Value && dt_elbow.Rows[k][col_sta2] != DBNull.Value)
                        {
                            double sta = Convert.ToDouble(dt_elbow.Rows[k][col_elbow_pi]);
                            double sta1 = Convert.ToDouble(dt_elbow.Rows[k][col_sta1]);
                            double sta2 = Convert.ToDouble(dt_elbow.Rows[k][col_sta2]);

                            dt10.Rows.Add();
                            dt10.Rows[dt10.Rows.Count - 1][0] = sta1;
                            dt10.Rows[dt10.Rows.Count - 1][1] = sta;
                            dt10.Rows[dt10.Rows.Count - 1][2] = sta2;
                            dt10.Rows[dt10.Rows.Count - 1][3] = dt_elbow.Rows[k][col_elbow_angle];

                            double min1 = 50;
                            double min2 = 50;
                            for (int i = 0; i < dt_transition.Rows.Count; ++i)
                            {
                                if (dt_transition.Rows[i][col_sta] != DBNull.Value)
                                {
                                    double sta3 = Convert.ToDouble(dt_transition.Rows[i][col_sta]);
                                    if (Math.Abs(sta1 - sta3) < min1)
                                    {
                                        min1 = Math.Abs(sta1 - sta3);
                                        dt10.Rows[dt10.Rows.Count - 1][4] = sta3;
                                        dt10.Rows[dt10.Rows.Count - 1][6] = min1;
                                    }
                                    if (Math.Abs(sta2 - sta3) < min2)
                                    {
                                        min2 = Math.Abs(sta2 - sta3);
                                        dt10.Rows[dt10.Rows.Count - 1][5] = sta3;
                                        dt10.Rows[dt10.Rows.Count - 1][7] = min2;
                                    }
                                }
                            }



                        }
                    }
                    Functions.Transfer_datatable_to_new_excel_spreadsheet(dt10, "Transition_Elbow");
                }


                Functions.Transfer_datatable_to_new_excel_spreadsheet(dt9, "Transition_PI");

            }

            #region defining OD tables

            dt_compiled.Columns.Add("x1", typeof(double));
            dt_compiled.Columns.Add("y1", typeof(double));
            dt_compiled.Columns.Add("x2", typeof(double));
            dt_compiled.Columns.Add("y2", typeof(double));
            dt_compiled.Columns.Add("layer", typeof(string));
            dt_compiled.Columns.Add("ci", typeof(short));

            System.Data.DataTable dt_od_pipe = new System.Data.DataTable();
            dt_od_pipe.Columns.Add(col_od_pipe_type, typeof(string));
            dt_od_pipe.Columns.Add(col_od_wt, typeof(string));
            dt_od_pipe.Columns.Add(col_od_coat, typeof(string));
            dt_od_pipe.Columns.Add(col_od_class, typeof(string));
            dt_od_pipe.Columns.Add(col_od_mat_descr, typeof(string));
            dt_od_pipe.Columns.Add(col_od_descr, typeof(string));
            dt_od_pipe.Columns.Add(col_od_start, typeof(string));
            dt_od_pipe.Columns.Add(col_od_end, typeof(string));
            dt_od_pipe.Columns.Add(col_len, typeof(string));

            dt_od_pipe.Columns.Add(col_notes, typeof(string));
            dt_od_pipe.Columns.Add("id", typeof(ObjectId));


            System.Data.DataTable dt_od_buoy = new System.Data.DataTable();
            dt_od_buoy.Columns.Add(col_feature, typeof(string));
            dt_od_buoy.Columns.Add(col_start, typeof(string));
            dt_od_buoy.Columns.Add(col_end, typeof(string));
            dt_od_buoy.Columns.Add(col_len, typeof(string));
            dt_od_buoy.Columns.Add(col_spacing, typeof(string));
            dt_od_buoy.Columns.Add(col_count, typeof(string));
            dt_od_buoy.Columns.Add(col_just, typeof(string));
            dt_od_buoy.Columns.Add(col_notes, typeof(string));
            dt_od_buoy.Columns.Add("id", typeof(ObjectId));
            dt_od_buoy.Columns.Add("id1", typeof(ObjectId));
            dt_od_buoy.Columns.Add("id2", typeof(ObjectId));

            System.Data.DataTable dt_od_long_strap = new System.Data.DataTable();
            dt_od_long_strap = dt_od_buoy.Clone();

            System.Data.DataTable dt_od_cpac = new System.Data.DataTable();
            dt_od_cpac.Columns.Add(col_sta, typeof(string));
            dt_od_cpac.Columns.Add(col_descr, typeof(string));
            dt_od_cpac.Columns.Add(col_descr2, typeof(string));
            dt_od_cpac.Columns.Add(col_eq1, typeof(string));
            dt_od_cpac.Columns.Add(col_eq2, typeof(string));
            dt_od_cpac.Columns.Add(col_eq3, typeof(string));
            dt_od_cpac.Columns.Add(col_just, typeof(string));
            dt_od_cpac.Columns.Add(col_notes, typeof(string));
            dt_od_cpac.Columns.Add("id", typeof(ObjectId));

            System.Data.DataTable dt_od_es = new System.Data.DataTable();
            dt_od_es.Columns.Add(col_sta, typeof(string));
            dt_od_es.Columns.Add(col_ditchplug, typeof(string));
            dt_od_es.Columns.Add(col_es_spacing, typeof(string));
            dt_od_es.Columns.Add(col_notes, typeof(string));
            dt_od_es.Columns.Add("id", typeof(ObjectId));

            System.Data.DataTable dt_od_hydrotest_pt = new System.Data.DataTable();
            dt_od_hydrotest_pt.Columns.Add(col_sta, typeof(string));
            dt_od_hydrotest_pt.Columns.Add(col_descr, typeof(string));
            dt_od_hydrotest_pt.Columns.Add("id", typeof(ObjectId));

            System.Data.DataTable dt_od_hydrotest = new System.Data.DataTable();
            dt_od_hydrotest.Columns.Add(col_start, typeof(string));
            dt_od_hydrotest.Columns.Add(col_end, typeof(string));
            dt_od_hydrotest.Columns.Add(col_len, typeof(string));
            dt_od_hydrotest.Columns.Add(col_descr, typeof(string));
            dt_od_hydrotest.Columns.Add("id", typeof(ObjectId));


            System.Data.DataTable dt_od_xing = new System.Data.DataTable();
            dt_od_xing.Columns.Add(col_od_xingid, typeof(string));
            dt_od_xing.Columns.Add(col_od_xingtype, typeof(string));
            dt_od_xing.Columns.Add(col_station, typeof(string));
            dt_od_xing.Columns.Add(col_descr1, typeof(string));
            dt_od_xing.Columns.Add(col_descr2, typeof(string));
            dt_od_xing.Columns.Add(col_od_ref_dwg_id, typeof(string));
            dt_od_xing.Columns.Add(col_od_min_depth, typeof(string));
            dt_od_xing.Columns.Add(col_od_xing_method, typeof(string));
            dt_od_xing.Columns.Add(col_od_pipe_type, typeof(string));
            dt_od_xing.Columns.Add(col_od_pipe_class, typeof(string));
            dt_od_xing.Columns.Add(col_od_wt, typeof(string));
            dt_od_xing.Columns.Add(col_just, typeof(string));
            dt_od_xing.Columns.Add("id", typeof(ObjectId));

            System.Data.DataTable dt_od_trans = new System.Data.DataTable();
            dt_od_trans.Columns.Add(col_sta, typeof(string));
            dt_od_trans.Columns.Add("id", typeof(ObjectId));

            System.Data.DataTable dt_od_elbow = new System.Data.DataTable();
            dt_od_elbow.Columns.Add(col_elbow_id, typeof(string));
            dt_od_elbow.Columns.Add(col_elbow_angle, typeof(string));
            dt_od_elbow.Columns.Add(col_bend_type, typeof(string));
            dt_od_elbow.Columns.Add(col_descr, typeof(string));
            dt_od_elbow.Columns.Add(col_elbow_ref_dwg, typeof(string));
            dt_od_elbow.Columns.Add(col_elbow_pi, typeof(string));
            dt_od_elbow.Columns.Add(col_start, typeof(string));
            dt_od_elbow.Columns.Add(col_end, typeof(string));
            dt_od_elbow.Columns.Add(col_length, typeof(string));
            dt_od_elbow.Columns.Add(col_pipe_type, typeof(string));
            dt_od_elbow.Columns.Add(col_pipe_class, typeof(string));
            dt_od_elbow.Columns.Add(col_wt, typeof(string));
            dt_od_elbow.Columns.Add(col_elbow_notes, typeof(string));
            dt_od_elbow.Columns.Add("id", typeof(ObjectId));

            System.Data.DataTable dt_od_fab = new System.Data.DataTable();
            dt_od_fab.Columns.Add(col_fac_name, typeof(string));
            dt_od_fab.Columns.Add(col_sta1, typeof(string));
            dt_od_fab.Columns.Add(col_sta2, typeof(string));
            dt_od_fab.Columns.Add(col_len, typeof(string));
            dt_od_fab.Columns.Add(col_descr1, typeof(string));
            dt_od_fab.Columns.Add(col_descr2, typeof(string));
            dt_od_fab.Columns.Add(col_just, typeof(string));
            dt_od_fab.Columns.Add("id", typeof(ObjectId));

            System.Data.DataTable dt_od_pre_existing = new System.Data.DataTable();
            dt_od_pre_existing.Columns.Add(col_start, typeof(string));
            dt_od_pre_existing.Columns.Add(col_end, typeof(string));
            dt_od_pre_existing.Columns.Add(col_len, typeof(string));
            dt_od_pre_existing.Columns.Add(col_descr, typeof(string));
            dt_od_pre_existing.Columns.Add(col_just, typeof(string));
            dt_od_pre_existing.Columns.Add(col_notes, typeof(string));
            dt_od_pre_existing.Columns.Add("id", typeof(ObjectId));


            System.Data.DataTable dt_od_class = new System.Data.DataTable();
            dt_od_class.Columns.Add(col_od_pipe_type, typeof(string));
            dt_od_class.Columns.Add(col_sta1, typeof(string));
            dt_od_class.Columns.Add(col_descr1, typeof(string));
            dt_od_class.Columns.Add(col_sta2, typeof(string));
            dt_od_class.Columns.Add(col_descr2, typeof(string));
            dt_od_class.Columns.Add(col_wt, typeof(string));
            dt_od_class.Columns.Add(col_just, typeof(string));
            dt_od_class.Columns.Add("id", typeof(ObjectId));

            System.Data.DataTable dt_od_geotech = new System.Data.DataTable();
            dt_od_geotech.Columns.Add(col_geotech_od_sta1, typeof(string));
            dt_od_geotech.Columns.Add(col_geotech_od_descr1, typeof(string));
            dt_od_geotech.Columns.Add(col_geotech_od_sta2, typeof(string));
            dt_od_geotech.Columns.Add(col_geotech_od_descr2, typeof(string));
            dt_od_geotech.Columns.Add(col_len, typeof(string));
            dt_od_geotech.Columns.Add(col_geotech_od_class, typeof(string));
            dt_od_geotech.Columns.Add(col_geotech_od_type, typeof(string));
            dt_od_geotech.Columns.Add(col_geotech_od_label, typeof(string));
            dt_od_geotech.Columns.Add(col_notes, typeof(string));
            dt_od_geotech.Columns.Add("id", typeof(ObjectId));
            dt_od_geotech.Columns.Add("id1", typeof(ObjectId));
            dt_od_geotech.Columns.Add("id2", typeof(ObjectId));

            System.Data.DataTable dt_od_doc = new System.Data.DataTable();
            dt_od_doc.Columns.Add(col_doc_od_sta1, typeof(string));
            dt_od_doc.Columns.Add(col_doc_od_sta2, typeof(string));
            dt_od_doc.Columns.Add(col_doc_od_min_cvr, typeof(string));
            dt_od_doc.Columns.Add(col_len, typeof(string));
            dt_od_doc.Columns.Add(col_just, typeof(string));
            dt_od_doc.Columns.Add(col_notes, typeof(string));
            dt_od_doc.Columns.Add("id", typeof(ObjectId));

            System.Data.DataTable dt_od_muskeg = new System.Data.DataTable();
            dt_od_muskeg.Columns.Add(col_muskeg_od_sta, typeof(string));
            dt_od_muskeg.Columns.Add(col_muskeg_od_descr, typeof(string));
            dt_od_muskeg.Columns.Add(col_muskeg_od_label, typeof(string));
            dt_od_muskeg.Columns.Add("id", typeof(ObjectId));


            #endregion

            if (dt_cl != null && dt_cl.Rows.Count > 1 && dt_compiled != null && dt_compiled.Rows.Count > 1)
            {
                if (MessageBox.Show("Do you want to continue?", "Transfer to current drawing", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    try
                    {
                        Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                        Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                        ObjectId[] Empty_array = null;
                        Editor1.SetImpliedSelection(Empty_array);

                        using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                        {
                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                            {

                                BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite);
                                LayerTable layertable1 = Trans1.GetObject(ThisDrawing.Database.LayerTableId, OpenMode.ForRead) as LayerTable;
                                create_pipe_od_table();
                                if (dt_elbow.Rows.Count > 0) create_elbow_od_table();
                                if (dt_fab.Rows.Count > 0) create_facility_od_table();
                                if (dt_pre_existing.Rows.Count > 0) create_pre_existing_od_table();
                                if (dt_class.Rows.Count > 0) create_class_od_table();

                                if (dt_buoy.Rows.Count > 0)
                                {
                                    create_buoyancy_od_table();
                                    create_buoyancy_pt_od_table();
                                }
                                else
                                {
                                    if (dt_long_strap.Rows.Count > 0)
                                    {
                                        create_buoyancy_pt_od_table();
                                    }
                                }

                                if (dt_cpac.Rows.Count > 0) create_cpac_od_table();
                                if (dt_es.Rows.Count > 0) create_es_od_table();

                                if (dt_hydrotest.Rows.Count > 0)
                                {
                                    create_hydrotest_lines_od_table();
                                    create_hydrotest_point_od_table();
                                }


                                if (dt_xing.Rows.Count > 0) create_xing_od_table();
                                if (dt_transition.Rows.Count > 0) create_transition_od_table();


                                if (dt_geotech.Rows.Count > 0)
                                {
                                    create_geotech_od_table();
                                    create_geotech_pt_od_table();
                                }

                                if (dt_doc.Rows.Count > 0) create_doc_od_table();
                                if (dt_muskeg.Rows.Count > 0) create_muskeg_od_table();


                                Polyline Poly2D = Functions.Build_2d_poly_for_scanning(dt_cl);
                                Polyline3d Poly3D = Functions.Build_3d_poly_for_scanning(dt_cl);

                                double last_sta = 0;
                                if (dt_cl.Rows[dt_cl.Rows.Count - 1]["3DSta"] != DBNull.Value)
                                {
                                    last_sta = Convert.ToDouble(Convert.ToString(dt_cl.Rows[dt_cl.Rows.Count - 1]["3DSta"]).Replace("+", ""));
                                }



                                if (checkBox_draw_canadian_mat.Checked == true)
                                {
                                    #region pipes

                                    for (int i = 0; i < dt_compiled.Rows.Count; ++i)
                                    {
                                        if (dt_compiled.Rows[i][col_sta1] != DBNull.Value && dt_compiled.Rows[i][col_sta2] != DBNull.Value)
                                        {
                                            double staH1 = Convert.ToDouble(dt_compiled.Rows[i][col_sta1]);
                                            double staH2 = Convert.ToDouble(dt_compiled.Rows[i][col_sta2]);

                                            if (staH2 > last_sta) staH2 = last_sta - 0.0001;

                                            for (int j = 0; j < dt_cl.Rows.Count - 1; ++j)
                                            {
                                                if (dt_cl.Rows[j]["3DSta"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", "")) == true &&
                                                    dt_cl.Rows[j + 1]["3DSta"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", "")) == true &&
                                                    dt_cl.Rows[j]["X"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["X"])) == true &&
                                                    dt_cl.Rows[j]["Y"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Y"])) == true &&
                                                    dt_cl.Rows[j]["Z"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Z"])) == true &&
                                                     dt_cl.Rows[j + 1]["X"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["X"])) == true &&
                                                    dt_cl.Rows[j + 1]["Y"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Y"])) == true &&
                                                    dt_cl.Rows[j + 1]["Z"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Z"])) == true)
                                                {
                                                    double sta1 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", ""));
                                                    double sta2 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", ""));



                                                    if (staH1 >= sta1 && staH1 <= sta2)
                                                    {
                                                        double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                        double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);

                                                        double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                        double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);

                                                        double x = x1 + (x2 - x1) * (staH1 - sta1) / (sta2 - sta1);
                                                        double y = y1 + (y2 - y1) * (staH1 - sta1) / (sta2 - sta1);

                                                        dt_compiled.Rows[i]["x1"] = x;
                                                        dt_compiled.Rows[i]["y1"] = y;
                                                        dt_compiled.Rows[i]["layer"] = pipes_layer;
                                                    }




                                                    if (staH2 >= sta1 && staH2 <= sta2)
                                                    {
                                                        double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                        double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);
                                                        double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                        double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);
                                                        double x = x1 + (x2 - x1) * (staH2 - sta1) / (sta2 - sta1);
                                                        double y = y1 + (y2 - y1) * (staH2 - sta1) / (sta2 - sta1);
                                                        dt_compiled.Rows[i]["x2"] = x;
                                                        dt_compiled.Rows[i]["y2"] = y;
                                                        dt_compiled.Rows[i]["layer"] = pipes_layer;
                                                    }
                                                }
                                            }
                                        }

                                        string mat1 = Convert.ToString(dt_compiled.Rows[i][col_pipe_type]);

                                        switch (mat1)
                                        {
                                            case "1":
                                                dt_compiled.Rows[i]["ci"] = 9;
                                                break;
                                            case "2":
                                                dt_compiled.Rows[i]["ci"] = 7;
                                                break;
                                            case "3":
                                                dt_compiled.Rows[i]["ci"] = 1;
                                                break;
                                            case "4":
                                                dt_compiled.Rows[i]["ci"] = 2;
                                                break;
                                            case "5":
                                                dt_compiled.Rows[i]["ci"] = 3;
                                                break;
                                            case "6":
                                                dt_compiled.Rows[i]["ci"] = 4;
                                                break;
                                            case "7":
                                                dt_compiled.Rows[i]["ci"] = 7;
                                                break;
                                            default:
                                                dt_compiled.Rows[i]["ci"] = 6;
                                                break;
                                        }
                                    }


                                    //Functions.Transfer_datatable_to_new_excel_spreadsheet(dt_compiled,"Compiled Data");

                                    for (int i = 0; i < dt_compiled.Rows.Count; ++i)
                                    {
                                        bool is_elbow = false;

                                        string justif1 = "";
                                        if (dt_compiled.Rows[i][col_just] != DBNull.Value)
                                        {
                                            justif1 = Convert.ToString(dt_compiled.Rows[i][col_just]);
                                        }

                                        if (justif1.Contains("**elbow") == true)
                                        {
                                            is_elbow = true;
                                        }

                                        bool is_facility = false;
                                        if (dt_compiled.Rows[i][col_pipe_type] == DBNull.Value)
                                        {
                                            is_facility = true;
                                        }
                                        if (is_elbow == false && is_facility == false)
                                        {
                                            double x1 = Convert.ToDouble(dt_compiled.Rows[i]["x1"]);
                                            double y1 = Convert.ToDouble(dt_compiled.Rows[i]["y1"]);
                                            double x2 = Convert.ToDouble(dt_compiled.Rows[i]["x2"]);
                                            double y2 = Convert.ToDouble(dt_compiled.Rows[i]["y2"]);
                                            Point3d pt1 = Poly2D.GetClosestPointTo(new Point3d(x1, y1, 0), Vector3d.ZAxis, false);
                                            Point3d pt2 = Poly2D.GetClosestPointTo(new Point3d(x2, y2, 0), Vector3d.ZAxis, false);

                                            double param1 = Poly2D.GetParameterAtPoint(pt1);
                                            double param2 = Poly2D.GetParameterAtPoint(pt2);

                                            Polyline poly_heavy_wall = Functions.get_part_of_poly(Poly2D, param1, param2);

                                            if (dt_compiled.Rows[i]["ci"] != DBNull.Value && dt_compiled.Rows[i]["layer"] != DBNull.Value)
                                            {
                                                string layer1 = Convert.ToString(dt_compiled.Rows[i]["layer"]);
                                                Functions.Creaza_layer(layer1, 7, true);
                                                poly_heavy_wall.Layer = layer1;
                                                poly_heavy_wall.ColorIndex = Convert.ToInt16(dt_compiled.Rows[i]["ci"]);
                                            }
                                            BTrecord.AppendEntity(poly_heavy_wall);
                                            Trans1.AddNewlyCreatedDBObject(poly_heavy_wall, true);

                                            dt_od_pipe.Rows.Add();
                                            dt_od_pipe.Rows[dt_od_pipe.Rows.Count - 1]["id"] = poly_heavy_wall.ObjectId;
                                            dt_od_pipe.Rows[dt_od_pipe.Rows.Count - 1][col_od_pipe_type] = dt_compiled.Rows[i][col_pipe_type];



                                            dt_od_pipe.Rows[dt_od_pipe.Rows.Count - 1][col_od_start] = Functions.Get_chainage_from_double(Convert.ToDouble(dt_compiled.Rows[i][col_sta1]), "m", 1);
                                            dt_od_pipe.Rows[dt_od_pipe.Rows.Count - 1][col_od_end] = Functions.Get_chainage_from_double(Convert.ToDouble(dt_compiled.Rows[i][col_sta2]), "m", 1);


                                            double extra1 = 0;

                                            double s1 = Convert.ToDouble(dt_compiled.Rows[i][col_sta1]);
                                            double s2 = Convert.ToDouble(dt_compiled.Rows[i][col_sta2]);

                                            if (dt_eq != null && dt_eq.Rows.Count > 0)
                                            {
                                                for (int k = 0; k < dt_eq.Rows.Count; ++k)
                                                {
                                                    if (dt_eq.Rows[k][col_back] != DBNull.Value && dt_eq.Rows[k][col_ahead] != DBNull.Value)
                                                    {
                                                        double back1 = Convert.ToDouble(dt_eq.Rows[k][col_back]);
                                                        double ahead1 = Convert.ToDouble(dt_eq.Rows[k][col_ahead]);

                                                        if (s1 < back1 && ahead1 < s2)
                                                        {
                                                            extra1 = extra1 + ahead1 - back1;
                                                        }

                                                        if (s1 == back1) s1 = ahead1;

                                                    }
                                                }
                                            }

                                            dt_od_pipe.Rows[dt_od_pipe.Rows.Count - 1][col_len] = Math.Round(Convert.ToDecimal(s2) - Convert.ToDecimal(s1) - Convert.ToDecimal(extra1), 1);


                                            dt_od_pipe.Rows[dt_od_pipe.Rows.Count - 1][col_od_wt] = dt_compiled.Rows[i][col_wt];
                                            dt_od_pipe.Rows[dt_od_pipe.Rows.Count - 1][col_od_descr] = dt_compiled.Rows[i][col_just];
                                            dt_od_pipe.Rows[dt_od_pipe.Rows.Count - 1][col_od_coat] = dt_compiled.Rows[i][col_coating];
                                            dt_od_pipe.Rows[dt_od_pipe.Rows.Count - 1][col_od_mat_descr] = dt_compiled.Rows[i][col_descr];
                                            dt_od_pipe.Rows[dt_od_pipe.Rows.Count - 1][col_od_class] = dt_compiled.Rows[i][col_pipe_class];
                                            dt_od_pipe.Rows[dt_od_pipe.Rows.Count - 1][col_notes] = dt_compiled.Rows[i][col_notes];
                                        }


                                    }

                                    #endregion

                                    #region Buoyancy

                                    for (int i = 0; i < dt_buoy.Rows.Count; ++i)
                                    {
                                        if (dt_buoy.Rows[i][col_start] != DBNull.Value && dt_buoy.Rows[i][col_end] != DBNull.Value)
                                        {
                                            double staH1 = Convert.ToDouble(dt_buoy.Rows[i][col_start]);
                                            double staH2 = Convert.ToDouble(dt_buoy.Rows[i][col_end]);

                                            if (staH2 > last_sta) staH2 = last_sta - 0.0001;

                                            for (int j = 0; j < dt_cl.Rows.Count - 1; ++j)
                                            {
                                                if (dt_cl.Rows[j]["3DSta"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", "")) == true &&
                                                    dt_cl.Rows[j + 1]["3DSta"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", "")) == true &&
                                                    dt_cl.Rows[j]["X"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["X"])) == true &&
                                                    dt_cl.Rows[j]["Y"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Y"])) == true &&
                                                    dt_cl.Rows[j]["Z"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Z"])) == true &&
                                                     dt_cl.Rows[j + 1]["X"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["X"])) == true &&
                                                    dt_cl.Rows[j + 1]["Y"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Y"])) == true &&
                                                    dt_cl.Rows[j + 1]["Z"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Z"])) == true)
                                                {
                                                    double sta1 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", ""));
                                                    double sta2 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", ""));
                                                    if (staH1 >= sta1 && staH1 <= sta2)
                                                    {
                                                        double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                        double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);

                                                        double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                        double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);

                                                        double x = x1 + (x2 - x1) * (staH1 - sta1) / (sta2 - sta1);
                                                        double y = y1 + (y2 - y1) * (staH1 - sta1) / (sta2 - sta1);

                                                        dt_buoy.Rows[i]["x1"] = x;
                                                        dt_buoy.Rows[i]["y1"] = y;
                                                        dt_buoy.Rows[i]["layer"] = buoyancy_layer;
                                                    }

                                                    if (staH2 >= sta1 && staH2 <= sta2)
                                                    {
                                                        double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                        double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);
                                                        double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                        double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);
                                                        double x = x1 + (x2 - x1) * (staH2 - sta1) / (sta2 - sta1);
                                                        double y = y1 + (y2 - y1) * (staH2 - sta1) / (sta2 - sta1);
                                                        dt_buoy.Rows[i]["x2"] = x;
                                                        dt_buoy.Rows[i]["y2"] = y;
                                                        dt_buoy.Rows[i]["layer"] = buoyancy_layer;
                                                    }
                                                }
                                            }
                                        }

                                        string feature1 = Convert.ToString(dt_buoy.Rows[i][col_feature]);

                                        string feat1 = feature1.Replace(" ", "");

                                        switch (feat1)
                                        {
                                            case "screwanchors":
                                                dt_buoy.Rows[i]["ci"] = 1;
                                                break;
                                            case "concreteriverweights":
                                                dt_buoy.Rows[i]["ci"] = 5;
                                                break;
                                            case "continuousconcretecoating":
                                                dt_buoy.Rows[i]["ci"] = 2;
                                                break;
                                            case "screwanchor":
                                                dt_buoy.Rows[i]["ci"] = 1;
                                                break;
                                            case "concreteriverweight":
                                                dt_buoy.Rows[i]["ci"] = 5;
                                                break;
                                            case "concretecoating":
                                                dt_buoy.Rows[i]["ci"] = 2;
                                                break;
                                            case "sa":
                                                dt_buoy.Rows[i]["ci"] = 1;
                                                break;
                                            case "crw":
                                                dt_buoy.Rows[i]["ci"] = 5;
                                                break;
                                            case "ccc":
                                                dt_buoy.Rows[i]["ci"] = 2;
                                                break;
                                            case "cc":
                                                dt_buoy.Rows[i]["ci"] = 2;
                                                break;
                                            default:
                                                dt_buoy.Rows[i]["ci"] = 3;
                                                break;
                                        }
                                    }

                                    for (int i = 0; i < dt_long_strap.Rows.Count; ++i)
                                    {
                                        if (dt_long_strap.Rows[i][col_start] != DBNull.Value && dt_long_strap.Rows[i][col_end] != DBNull.Value)
                                        {
                                            double staH1 = Convert.ToDouble(dt_long_strap.Rows[i][col_start]);


                                            if (staH1 > last_sta) staH1 = last_sta - 0.0001;

                                            for (int j = 0; j < dt_cl.Rows.Count - 1; ++j)
                                            {
                                                if (dt_cl.Rows[j]["3DSta"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", "")) == true &&
                                                    dt_cl.Rows[j + 1]["3DSta"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", "")) == true &&
                                                    dt_cl.Rows[j]["X"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["X"])) == true &&
                                                    dt_cl.Rows[j]["Y"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Y"])) == true &&
                                                    dt_cl.Rows[j]["Z"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Z"])) == true &&
                                                     dt_cl.Rows[j + 1]["X"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["X"])) == true &&
                                                    dt_cl.Rows[j + 1]["Y"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Y"])) == true &&
                                                    dt_cl.Rows[j + 1]["Z"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Z"])) == true)
                                                {
                                                    double sta1 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", ""));
                                                    double sta2 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", ""));
                                                    if (staH1 >= sta1 && staH1 <= sta2)
                                                    {
                                                        double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                        double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);

                                                        double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                        double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);

                                                        double x = x1 + (x2 - x1) * (staH1 - sta1) / (sta2 - sta1);
                                                        double y = y1 + (y2 - y1) * (staH1 - sta1) / (sta2 - sta1);

                                                        dt_long_strap.Rows[i]["x1"] = x;
                                                        dt_long_strap.Rows[i]["y1"] = y;
                                                        dt_long_strap.Rows[i]["layer"] = buoyancy_pt_layer;
                                                    }
                                                }
                                            }
                                        }


                                    }


                                    Functions.Creaza_layer(buoyancy_pt_layer, 7, true);

                                    for (int i = 0; i < dt_buoy.Rows.Count; ++i)
                                    {
                                        double x1 = Convert.ToDouble(dt_buoy.Rows[i]["x1"]);
                                        double y1 = Convert.ToDouble(dt_buoy.Rows[i]["y1"]);
                                        double x2 = Convert.ToDouble(dt_buoy.Rows[i]["x2"]);
                                        double y2 = Convert.ToDouble(dt_buoy.Rows[i]["y2"]);
                                        Point3d pt1 = Poly2D.GetClosestPointTo(new Point3d(x1, y1, 0), Vector3d.ZAxis, false);
                                        Point3d pt2 = Poly2D.GetClosestPointTo(new Point3d(x2, y2, 0), Vector3d.ZAxis, false);



                                        double param1 = Poly2D.GetParameterAtPoint(pt1);
                                        double param2 = Poly2D.GetParameterAtPoint(pt2);

                                        DBPoint pt_buoy1 = new DBPoint(pt1);
                                        DBPoint pt_buoy2 = new DBPoint(pt2);


                                        Polyline buoy = Functions.get_part_of_poly(Poly2D, param1, param2);

                                        if (dt_buoy.Rows[i]["ci"] != DBNull.Value && dt_buoy.Rows[i]["layer"] != DBNull.Value)
                                        {
                                            string layer1 = Convert.ToString(dt_buoy.Rows[i]["layer"]);
                                            Functions.Creaza_layer(layer1, 7, true);
                                            buoy.Layer = layer1;
                                            buoy.ColorIndex = Convert.ToInt16(dt_buoy.Rows[i]["ci"]);

                                            pt_buoy1.Layer = buoyancy_pt_layer;
                                            pt_buoy1.ColorIndex = 256;
                                            pt_buoy2.Layer = buoyancy_pt_layer;
                                            pt_buoy2.ColorIndex = 256;

                                        }
                                        BTrecord.AppendEntity(buoy);
                                        Trans1.AddNewlyCreatedDBObject(buoy, true);
                                        BTrecord.AppendEntity(pt_buoy1);
                                        Trans1.AddNewlyCreatedDBObject(pt_buoy1, true);
                                        BTrecord.AppendEntity(pt_buoy2);
                                        Trans1.AddNewlyCreatedDBObject(pt_buoy2, true);






                                        dt_od_buoy.Rows.Add();
                                        dt_od_buoy.Rows[dt_od_buoy.Rows.Count - 1]["id"] = buoy.ObjectId;
                                        dt_od_buoy.Rows[dt_od_buoy.Rows.Count - 1]["id1"] = pt_buoy1.ObjectId;
                                        dt_od_buoy.Rows[dt_od_buoy.Rows.Count - 1]["id2"] = pt_buoy2.ObjectId;
                                        dt_od_buoy.Rows[dt_od_buoy.Rows.Count - 1][col_feature] = dt_buoy.Rows[i][col_feature];

                                        dt_od_buoy.Rows[dt_od_buoy.Rows.Count - 1][col_spacing] = dt_buoy.Rows[i][col_spacing];
                                        dt_od_buoy.Rows[dt_od_buoy.Rows.Count - 1][col_count] = dt_buoy.Rows[i][col_count];
                                        dt_od_buoy.Rows[dt_od_buoy.Rows.Count - 1][col_just] = dt_buoy.Rows[i][col_just];
                                        dt_od_buoy.Rows[dt_od_buoy.Rows.Count - 1][col_notes] = dt_buoy.Rows[i][col_notes];


                                        double s1 = -1.234;
                                        double s2 = -1.234;

                                        if (dt_buoy.Rows[i][col_start] != DBNull.Value)
                                        {
                                            s1 = Convert.ToDouble(dt_buoy.Rows[i][col_start]);
                                            dt_od_buoy.Rows[dt_od_buoy.Rows.Count - 1][col_start] = Functions.Get_chainage_from_double(s1, "m", 1);
                                        }
                                        else
                                        {
                                            dt_od_buoy.Rows[dt_od_buoy.Rows.Count - 1][col_start] = dt_buoy.Rows[i][col_start];
                                        }

                                        if (dt_buoy.Rows[i][col_end] != DBNull.Value)
                                        {
                                            s2 = Convert.ToDouble(dt_buoy.Rows[i][col_end]);
                                            dt_od_buoy.Rows[dt_od_buoy.Rows.Count - 1][col_end] = Functions.Get_chainage_from_double(s2, "m", 1);
                                        }
                                        else
                                        {
                                            dt_od_buoy.Rows[dt_od_buoy.Rows.Count - 1][col_end] = dt_buoy.Rows[i][col_end];
                                        }

                                        if (s1 != -1.234 && s2 != -1.234)
                                        {

                                            double extra1 = 0;

                                            if (dt_eq != null && dt_eq.Rows.Count > 0)
                                            {
                                                for (int k = 0; k < dt_eq.Rows.Count; ++k)
                                                {
                                                    if (dt_eq.Rows[k][col_back] != DBNull.Value && dt_eq.Rows[k][col_ahead] != DBNull.Value)
                                                    {
                                                        double back1 = Convert.ToDouble(dt_eq.Rows[k][col_back]);
                                                        double ahead1 = Convert.ToDouble(dt_eq.Rows[k][col_ahead]);

                                                        if (s1 < back1 && ahead1 < s2)
                                                        {
                                                            extra1 = extra1 + ahead1 - back1;
                                                        }

                                                        if (s1 == back1) s1 = ahead1;

                                                    }
                                                }
                                            }

                                            dt_od_buoy.Rows[dt_od_buoy.Rows.Count - 1][col_len] = Math.Round(Convert.ToDecimal(s2) - Convert.ToDecimal(s1) - Convert.ToDecimal(extra1), 1);

                                        }
                                    }


                                    for (int i = 0; i < dt_long_strap.Rows.Count; ++i)
                                    {
                                        double x1 = Convert.ToDouble(dt_long_strap.Rows[i]["x1"]);
                                        double y1 = Convert.ToDouble(dt_long_strap.Rows[i]["y1"]);
                                        Point3d pt1 = Poly2D.GetClosestPointTo(new Point3d(x1, y1, 0), Vector3d.ZAxis, false);




                                        double param1 = Poly2D.GetParameterAtPoint(pt1);


                                        DBPoint pt_long_strap1 = new DBPoint(pt1);
                                        pt_long_strap1.Layer = buoyancy_pt_layer;



                                        BTrecord.AppendEntity(pt_long_strap1);
                                        Trans1.AddNewlyCreatedDBObject(pt_long_strap1, true);







                                        dt_od_long_strap.Rows.Add();

                                        dt_od_long_strap.Rows[dt_od_long_strap.Rows.Count - 1]["id1"] = pt_long_strap1.ObjectId;

                                        dt_od_long_strap.Rows[dt_od_long_strap.Rows.Count - 1][col_feature] = dt_long_strap.Rows[i][col_feature];

                                        dt_od_long_strap.Rows[dt_od_long_strap.Rows.Count - 1][col_spacing] = dt_long_strap.Rows[i][col_spacing];
                                        dt_od_long_strap.Rows[dt_od_long_strap.Rows.Count - 1][col_count] = dt_long_strap.Rows[i][col_count];
                                        dt_od_long_strap.Rows[dt_od_long_strap.Rows.Count - 1][col_just] = dt_long_strap.Rows[i][col_just];
                                        dt_od_long_strap.Rows[dt_od_long_strap.Rows.Count - 1][col_notes] = dt_long_strap.Rows[i][col_notes];


                                        double s1 = -1.234;


                                        if (dt_long_strap.Rows[i][col_start] != DBNull.Value)
                                        {
                                            s1 = Convert.ToDouble(dt_long_strap.Rows[i][col_start]);
                                            dt_od_long_strap.Rows[dt_od_long_strap.Rows.Count - 1][col_start] = Functions.Get_chainage_from_double(s1, "m", 1);
                                        }
                                        else
                                        {
                                            dt_od_long_strap.Rows[dt_od_long_strap.Rows.Count - 1][col_start] = dt_long_strap.Rows[i][col_start];
                                        }
                                        dt_od_long_strap.Rows[dt_od_long_strap.Rows.Count - 1][col_len] = 1;


                                    }
                                    #endregion

                                    #region CPAC

                                    for (int i = 0; i < dt_cpac.Rows.Count; ++i)
                                    {
                                        if (dt_cpac.Rows[i][col_sta] != DBNull.Value)
                                        {
                                            double sta_cpac = Convert.ToDouble(dt_cpac.Rows[i][col_sta]);
                                            if (sta_cpac > last_sta) sta_cpac = last_sta - 0.0001;
                                            for (int j = 0; j < dt_cl.Rows.Count - 1; ++j)
                                            {
                                                if (dt_cl.Rows[j]["3DSta"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", "")) == true &&
                                                    dt_cl.Rows[j + 1]["3DSta"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", "")) == true &&
                                                    dt_cl.Rows[j]["X"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["X"])) == true &&
                                                    dt_cl.Rows[j]["Y"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Y"])) == true &&
                                                    dt_cl.Rows[j]["Z"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Z"])) == true &&
                                                     dt_cl.Rows[j + 1]["X"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["X"])) == true &&
                                                    dt_cl.Rows[j + 1]["Y"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Y"])) == true &&
                                                    dt_cl.Rows[j + 1]["Z"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Z"])) == true)
                                                {
                                                    double sta1 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", ""));
                                                    double sta2 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", ""));
                                                    if (sta_cpac >= sta1 && sta_cpac <= sta2)
                                                    {
                                                        double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                        double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);

                                                        double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                        double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);

                                                        double x = x1 + (x2 - x1) * (sta_cpac - sta1) / (sta2 - sta1);
                                                        double y = y1 + (y2 - y1) * (sta_cpac - sta1) / (sta2 - sta1);

                                                        dt_cpac.Rows[i]["x1"] = x;
                                                        dt_cpac.Rows[i]["y1"] = y;
                                                        dt_cpac.Rows[i]["layer"] = cpac_layer;
                                                    }
                                                }
                                            }
                                        }

                                    }

                                    for (int i = 0; i < dt_cpac.Rows.Count; ++i)
                                    {
                                        if (dt_cpac.Rows[i]["x1"] != DBNull.Value && dt_cpac.Rows[i]["y1"] != DBNull.Value)
                                        {
                                            double x1 = Convert.ToDouble(dt_cpac.Rows[i]["x1"]);
                                            double y1 = Convert.ToDouble(dt_cpac.Rows[i]["y1"]);

                                            Point3d pt1 = Poly2D.GetClosestPointTo(new Point3d(x1, y1, 0), Vector3d.ZAxis, false);

                                            DBPoint pt_cpac = new DBPoint(pt1);

                                            if (dt_cpac.Rows[i]["layer"] != DBNull.Value)
                                            {
                                                string layer1 = Convert.ToString(dt_cpac.Rows[i]["layer"]);
                                                Functions.Creaza_layer(layer1, 1, true);
                                                pt_cpac.Layer = layer1;
                                                pt_cpac.ColorIndex = 256;
                                            }
                                            BTrecord.AppendEntity(pt_cpac);
                                            Trans1.AddNewlyCreatedDBObject(pt_cpac, true);


                                            dt_od_cpac.Rows.Add();
                                            dt_od_cpac.Rows[dt_od_cpac.Rows.Count - 1]["id"] = pt_cpac.ObjectId;
                                            dt_od_cpac.Rows[dt_od_cpac.Rows.Count - 1][col_sta] = Functions.Get_chainage_from_double(Convert.ToDouble(dt_cpac.Rows[i][col_sta]), "m", 1);
                                            dt_od_cpac.Rows[dt_od_cpac.Rows.Count - 1][col_descr] = dt_cpac.Rows[i][col_descr];
                                            dt_od_cpac.Rows[dt_od_cpac.Rows.Count - 1][col_eq1] = dt_cpac.Rows[i][col_eq1];
                                            dt_od_cpac.Rows[dt_od_cpac.Rows.Count - 1][col_eq2] = dt_cpac.Rows[i][col_eq2];
                                            dt_od_cpac.Rows[dt_od_cpac.Rows.Count - 1][col_eq3] = dt_cpac.Rows[i][col_eq3];
                                            dt_od_cpac.Rows[dt_od_cpac.Rows.Count - 1][col_just] = dt_cpac.Rows[i][col_just];
                                            dt_od_cpac.Rows[dt_od_cpac.Rows.Count - 1][col_descr2] = dt_cpac.Rows[i][col_descr2];
                                            dt_od_cpac.Rows[dt_od_cpac.Rows.Count - 1][col_notes] = dt_cpac.Rows[i][col_notes];
                                        }
                                    }
                                    #endregion

                                    #region e&s

                                    for (int i = 0; i < dt_es.Rows.Count; ++i)
                                    {
                                        if (dt_es.Rows[i][col_sta] != DBNull.Value)
                                        {
                                            double sta_es = Convert.ToDouble(dt_es.Rows[i][col_sta]);
                                            if (sta_es > last_sta) sta_es = last_sta - 0.0001;

                                            for (int j = 0; j < dt_cl.Rows.Count - 1; ++j)
                                            {
                                                if (dt_cl.Rows[j]["3DSta"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", "")) == true &&
                                                    dt_cl.Rows[j + 1]["3DSta"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", "")) == true &&
                                                    dt_cl.Rows[j]["X"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["X"])) == true &&
                                                    dt_cl.Rows[j]["Y"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Y"])) == true &&
                                                    dt_cl.Rows[j]["Z"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Z"])) == true &&
                                                     dt_cl.Rows[j + 1]["X"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["X"])) == true &&
                                                    dt_cl.Rows[j + 1]["Y"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Y"])) == true &&
                                                    dt_cl.Rows[j + 1]["Z"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Z"])) == true)
                                                {
                                                    double sta1 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", ""));
                                                    double sta2 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", ""));
                                                    if (sta_es >= sta1 && sta_es <= sta2)
                                                    {
                                                        double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                        double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);

                                                        double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                        double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);

                                                        double x = x1 + (x2 - x1) * (sta_es - sta1) / (sta2 - sta1);
                                                        double y = y1 + (y2 - y1) * (sta_es - sta1) / (sta2 - sta1);

                                                        dt_es.Rows[i]["x1"] = x;
                                                        dt_es.Rows[i]["y1"] = y;
                                                        dt_es.Rows[i]["layer"] = es_layer;
                                                        if (dt_es.Rows[i][col_ditchplug] != DBNull.Value)
                                                        {
                                                            string dp = Convert.ToString(dt_es.Rows[i][col_ditchplug]);
                                                            if (dp.ToLower() == "y" || dp.ToLower() == "yes" || dp.ToLower() == "true")
                                                            {
                                                                dt_es.Rows[i]["ci"] = 2;

                                                            }
                                                            else
                                                            {
                                                                dt_es.Rows[i]["ci"] = 7;

                                                            }
                                                        }
                                                        else
                                                        {
                                                            dt_es.Rows[i]["ci"] = 7;
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                    }

                                    for (int i = 0; i < dt_es.Rows.Count; ++i)
                                    {
                                        double x1 = Convert.ToDouble(dt_es.Rows[i]["x1"]);
                                        double y1 = Convert.ToDouble(dt_es.Rows[i]["y1"]);

                                        Point3d pt1 = Poly2D.GetClosestPointTo(new Point3d(x1, y1, 0), Vector3d.ZAxis, false);

                                        DBPoint pt_es = new DBPoint(pt1);

                                        if (dt_es.Rows[i]["layer"] != DBNull.Value && dt_es.Rows[i]["ci"] != DBNull.Value)
                                        {
                                            string layer1 = Convert.ToString(dt_es.Rows[i]["layer"]);
                                            short ci = Convert.ToInt16(dt_es.Rows[i]["ci"]);
                                            Functions.Creaza_layer(layer1, 1, true);
                                            pt_es.Layer = layer1;
                                            pt_es.ColorIndex = ci;
                                        }
                                        BTrecord.AppendEntity(pt_es);
                                        Trans1.AddNewlyCreatedDBObject(pt_es, true);


                                        dt_od_es.Rows.Add();
                                        dt_od_es.Rows[dt_od_es.Rows.Count - 1]["id"] = pt_es.ObjectId;
                                        dt_od_es.Rows[dt_od_es.Rows.Count - 1][col_sta] = Functions.Get_chainage_from_double(Convert.ToDouble(dt_es.Rows[i][col_sta]), "m", 1);
                                        dt_od_es.Rows[dt_od_es.Rows.Count - 1][col_ditchplug] = dt_es.Rows[i][col_ditchplug];
                                        dt_od_es.Rows[dt_od_es.Rows.Count - 1][col_notes] = dt_es.Rows[i][col_notes];
                                        dt_od_es.Rows[dt_od_es.Rows.Count - 1][col_es_spacing] = dt_es.Rows[i][col_es_spacing];

                                    }
                                    #endregion

                                    #region hydrotest
                                    for (int i = 0; i < dt_hydrotest.Rows.Count; ++i)
                                    {
                                        if (dt_hydrotest.Rows[i][col_sta1] != DBNull.Value && dt_hydrotest.Rows[i][col_sta2] != DBNull.Value)
                                        {
                                            double sta_hydrotest1 = Convert.ToDouble(dt_hydrotest.Rows[i][col_sta1]);
                                            double sta_hydrotest2 = Convert.ToDouble(dt_hydrotest.Rows[i][col_sta2]);

                                            if (sta_hydrotest2 > last_sta) sta_hydrotest2 = last_sta - 0.0001;

                                            for (int j = 0; j < dt_cl.Rows.Count - 1; ++j)
                                            {
                                                if (dt_cl.Rows[j]["3DSta"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", "")) == true &&
                                                    dt_cl.Rows[j + 1]["3DSta"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", "")) == true &&
                                                    dt_cl.Rows[j]["X"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["X"])) == true &&
                                                    dt_cl.Rows[j]["Y"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Y"])) == true &&
                                                    dt_cl.Rows[j]["Z"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Z"])) == true &&
                                                     dt_cl.Rows[j + 1]["X"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["X"])) == true &&
                                                    dt_cl.Rows[j + 1]["Y"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Y"])) == true &&
                                                    dt_cl.Rows[j + 1]["Z"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Z"])) == true)
                                                {
                                                    double sta1 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", ""));
                                                    double sta2 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", ""));
                                                    if (sta_hydrotest1 >= sta1 && sta_hydrotest1 <= sta2)
                                                    {
                                                        double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                        double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);

                                                        double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                        double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);

                                                        double x = x1 + (x2 - x1) * (sta_hydrotest1 - sta1) / (sta2 - sta1);
                                                        double y = y1 + (y2 - y1) * (sta_hydrotest1 - sta1) / (sta2 - sta1);

                                                        dt_hydrotest.Rows[i]["x1"] = x;
                                                        dt_hydrotest.Rows[i]["y1"] = y;
                                                        dt_hydrotest.Rows[i]["layer"] = hydrotest_layerPT;

                                                    }
                                                    if (sta_hydrotest2 >= sta1 && sta_hydrotest2 <= sta2)
                                                    {
                                                        double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                        double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);

                                                        double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                        double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);

                                                        double x = x1 + (x2 - x1) * (sta_hydrotest2 - sta1) / (sta2 - sta1);
                                                        double y = y1 + (y2 - y1) * (sta_hydrotest2 - sta1) / (sta2 - sta1);

                                                        dt_hydrotest.Rows[i]["x2"] = x;
                                                        dt_hydrotest.Rows[i]["y2"] = y;
                                                        dt_hydrotest.Rows[i]["layer"] = hydrotest_layerPT;

                                                    }

                                                }
                                            }
                                        }

                                    }
                                    string descr = "";

                                    Functions.Creaza_layer(hydrotest_layer, 7, true);

                                    for (int i = 0; i < dt_hydrotest.Rows.Count; ++i)
                                    {
                                        double x1 = Convert.ToDouble(dt_hydrotest.Rows[i]["x1"]);
                                        double y1 = Convert.ToDouble(dt_hydrotest.Rows[i]["y1"]);

                                        double x2 = Convert.ToDouble(dt_hydrotest.Rows[i]["x2"]);
                                        double y2 = Convert.ToDouble(dt_hydrotest.Rows[i]["y2"]);

                                        Point3d pt1 = Poly2D.GetClosestPointTo(new Point3d(x1, y1, 0), Vector3d.ZAxis, false);
                                        Point3d pt2 = Poly2D.GetClosestPointTo(new Point3d(x2, y2, 0), Vector3d.ZAxis, false);

                                        #region db points
                                        DBPoint pt_hydrotest1 = new DBPoint(pt1);
                                        DBPoint pt_hydrotest2 = new DBPoint(pt2);

                                        string descr1 = "";


                                        if (dt_hydrotest.Rows[i][col_descr1] != DBNull.Value)
                                        {
                                            descr1 = Convert.ToString(dt_hydrotest.Rows[i][col_descr1]);
                                        }


                                        if (dt_hydrotest.Rows[i]["layer"] != DBNull.Value)
                                        {
                                            string layer1 = Convert.ToString(dt_hydrotest.Rows[i]["layer"]);
                                            Functions.Creaza_layer(layer1, 4, true);
                                            pt_hydrotest1.Layer = layer1;
                                            pt_hydrotest1.ColorIndex = 256;
                                            pt_hydrotest2.Layer = layer1;
                                            pt_hydrotest2.ColorIndex = 256;
                                        }
                                        BTrecord.AppendEntity(pt_hydrotest1);
                                        Trans1.AddNewlyCreatedDBObject(pt_hydrotest1, true);


                                        dt_od_hydrotest_pt.Rows.Add();
                                        dt_od_hydrotest_pt.Rows[dt_od_hydrotest_pt.Rows.Count - 1]["id"] = pt_hydrotest1.ObjectId;
                                        dt_od_hydrotest_pt.Rows[dt_od_hydrotest_pt.Rows.Count - 1][col_sta] = Functions.Get_chainage_from_double(Convert.ToDouble(dt_hydrotest.Rows[i][col_sta1]), "m", 1);

                                        if (i == 0)
                                        {
                                            dt_od_hydrotest_pt.Rows[dt_od_hydrotest_pt.Rows.Count - 1][col_descr] = descr1;
                                        }
                                        else
                                        {
                                            dt_od_hydrotest_pt.Rows[dt_od_hydrotest_pt.Rows.Count - 1][col_descr] = descr + "/" + descr1;
                                        }

                                        if (i == dt_hydrotest.Rows.Count - 1)
                                        {
                                            BTrecord.AppendEntity(pt_hydrotest2);
                                            Trans1.AddNewlyCreatedDBObject(pt_hydrotest2, true);
                                            dt_od_hydrotest_pt.Rows.Add();
                                            dt_od_hydrotest_pt.Rows[dt_od_hydrotest_pt.Rows.Count - 1]["id"] = pt_hydrotest2.ObjectId;
                                            dt_od_hydrotest_pt.Rows[dt_od_hydrotest_pt.Rows.Count - 1][col_sta] = Functions.Get_chainage_from_double(Convert.ToDouble(dt_hydrotest.Rows[i][col_sta2]), "m", 1);
                                            dt_od_hydrotest_pt.Rows[dt_od_hydrotest_pt.Rows.Count - 1][col_descr] = dt_hydrotest.Rows[i][col_descr2];

                                        }



                                        if (dt_hydrotest.Rows[i][col_descr2] != DBNull.Value)
                                        {
                                            descr = Convert.ToString(dt_hydrotest.Rows[i][col_descr2]);
                                        }
                                        else
                                        {
                                            descr = "";
                                        }
                                        #endregion

                                        #region linework

                                        double param1 = Poly2D.GetParameterAtPoint(pt1);
                                        double param2 = Poly2D.GetParameterAtPoint(pt2);

                                        Polyline poly_tst = Functions.get_part_of_poly(Poly2D, param1, param2);
                                        poly_tst.Layer = hydrotest_layer;
                                        BTrecord.AppendEntity(poly_tst);
                                        Trans1.AddNewlyCreatedDBObject(poly_tst, true);

                                        dt_od_hydrotest.Rows.Add();
                                        dt_od_hydrotest.Rows[dt_od_hydrotest.Rows.Count - 1]["id"] = poly_tst.ObjectId;



                                        dt_od_hydrotest.Rows[dt_od_hydrotest.Rows.Count - 1][col_start] = Functions.Get_chainage_from_double(Convert.ToDouble(dt_hydrotest.Rows[i][col_sta1]), "m", 1);
                                        dt_od_hydrotest.Rows[dt_od_hydrotest.Rows.Count - 1][col_end] = Functions.Get_chainage_from_double(Convert.ToDouble(dt_hydrotest.Rows[i][col_sta2]), "m", 1);


                                        double extra1 = 0;

                                        double s1 = Convert.ToDouble(dt_hydrotest.Rows[i][col_sta1]);
                                        double s2 = Convert.ToDouble(dt_hydrotest.Rows[i][col_sta2]);

                                        if (dt_eq != null && dt_eq.Rows.Count > 0)
                                        {
                                            for (int k = 0; k < dt_eq.Rows.Count; ++k)
                                            {
                                                if (dt_eq.Rows[k][col_back] != DBNull.Value && dt_eq.Rows[k][col_ahead] != DBNull.Value)
                                                {
                                                    double back1 = Convert.ToDouble(dt_eq.Rows[k][col_back]);
                                                    double ahead1 = Convert.ToDouble(dt_eq.Rows[k][col_ahead]);

                                                    if (s1 < back1 && ahead1 < s2)
                                                    {
                                                        extra1 = extra1 + ahead1 - back1;
                                                    }

                                                    if (s1 == back1) s1 = ahead1;

                                                }
                                            }
                                        }

                                        dt_od_hydrotest.Rows[dt_od_hydrotest.Rows.Count - 1][col_len] = Math.Round(Convert.ToDecimal(s2) - Convert.ToDecimal(s1) - Convert.ToDecimal(extra1), 1);

                                        string secno = "";

                                        if (dt_hydrotest.Rows[i][col_tst_sec] != DBNull.Value)
                                        {
                                            secno = Convert.ToString(dt_hydrotest.Rows[i][col_tst_sec]);
                                        }

                                        dt_od_hydrotest.Rows[dt_od_hydrotest.Rows.Count - 1][col_descr] = "Hydotest Section " + secno;


                                        #endregion

                                    }
                                    #endregion

                                    #region Xing

                                    for (int i = 0; i < dt_xing.Rows.Count; ++i)
                                    {
                                        if (dt_xing.Rows[i][col_sta] != DBNull.Value)
                                        {
                                            double sta_xing = Convert.ToDouble(dt_xing.Rows[i][col_sta]);

                                            if (sta_xing > last_sta) sta_xing = last_sta - 0.0001;

                                            for (int j = 0; j < dt_cl.Rows.Count - 1; ++j)
                                            {
                                                if (dt_cl.Rows[j]["3DSta"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", "")) == true &&
                                                    dt_cl.Rows[j + 1]["3DSta"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", "")) == true &&
                                                    dt_cl.Rows[j]["X"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["X"])) == true &&
                                                    dt_cl.Rows[j]["Y"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Y"])) == true &&
                                                    dt_cl.Rows[j]["Z"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Z"])) == true &&
                                                     dt_cl.Rows[j + 1]["X"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["X"])) == true &&
                                                    dt_cl.Rows[j + 1]["Y"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Y"])) == true &&
                                                    dt_cl.Rows[j + 1]["Z"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Z"])) == true)
                                                {
                                                    double sta1 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", ""));
                                                    double sta2 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", ""));
                                                    if (sta_xing >= sta1 && sta_xing <= sta2)
                                                    {
                                                        double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                        double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);

                                                        double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                        double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);

                                                        double x = x1 + (x2 - x1) * (sta_xing - sta1) / (sta2 - sta1);
                                                        double y = y1 + (y2 - y1) * (sta_xing - sta1) / (sta2 - sta1);

                                                        dt_xing.Rows[i]["x1"] = x;
                                                        dt_xing.Rows[i]["y1"] = y;
                                                        dt_xing.Rows[i]["layer"] = xing_layer;
                                                    }
                                                }
                                            }
                                        }

                                    }

                                    for (int i = 0; i < dt_xing.Rows.Count; ++i)
                                    {
                                        double x1 = Convert.ToDouble(dt_xing.Rows[i]["x1"]);
                                        double y1 = Convert.ToDouble(dt_xing.Rows[i]["y1"]);

                                        Point3d pt1 = Poly2D.GetClosestPointTo(new Point3d(x1, y1, 0), Vector3d.ZAxis, false);

                                        DBPoint pt_xing = new DBPoint(pt1);

                                        if (dt_xing.Rows[i]["layer"] != DBNull.Value)
                                        {
                                            string layer1 = Convert.ToString(dt_xing.Rows[i]["layer"]);
                                            Functions.Creaza_layer(layer1, 5, true);
                                            pt_xing.Layer = layer1;
                                            pt_xing.ColorIndex = 256;
                                        }
                                        BTrecord.AppendEntity(pt_xing);
                                        Trans1.AddNewlyCreatedDBObject(pt_xing, true);


                                        dt_od_xing.Rows.Add();
                                        dt_od_xing.Rows[dt_od_xing.Rows.Count - 1]["id"] = pt_xing.ObjectId;
                                        dt_od_xing.Rows[dt_od_xing.Rows.Count - 1][col_station] = Functions.Get_chainage_from_double(Convert.ToDouble(dt_xing.Rows[i][col_sta]), "m", 1);
                                        dt_od_xing.Rows[dt_od_xing.Rows.Count - 1][col_descr1] = dt_xing.Rows[i][col_descr1];
                                        dt_od_xing.Rows[dt_od_xing.Rows.Count - 1][col_descr2] = dt_xing.Rows[i][col_descr2];
                                        dt_od_xing.Rows[dt_od_xing.Rows.Count - 1][col_od_xingid] = dt_xing.Rows[i][col_xingid];
                                        dt_od_xing.Rows[dt_od_xing.Rows.Count - 1][col_od_xingtype] = dt_xing.Rows[i][col_xingtype];
                                        dt_od_xing.Rows[dt_od_xing.Rows.Count - 1][col_od_ref_dwg_id] = dt_xing.Rows[i][col_ref_dwg_id];
                                        dt_od_xing.Rows[dt_od_xing.Rows.Count - 1][col_od_min_depth] = dt_xing.Rows[i][col_min_depth];
                                        dt_od_xing.Rows[dt_od_xing.Rows.Count - 1][col_od_xing_method] = dt_xing.Rows[i][col_xing_method];
                                        dt_od_xing.Rows[dt_od_xing.Rows.Count - 1][col_od_pipe_type] = dt_xing.Rows[i][col_pipe_type];
                                        dt_od_xing.Rows[dt_od_xing.Rows.Count - 1][col_od_pipe_class] = dt_xing.Rows[i][col_pipe_class];
                                        dt_od_xing.Rows[dt_od_xing.Rows.Count - 1][col_od_wt] = dt_xing.Rows[i][col_wt];
                                        dt_od_xing.Rows[dt_od_xing.Rows.Count - 1][col_just] = dt_xing.Rows[i][col_just];
                                    }
                                    //Functions.Transfer_datatable_to_new_excel_spreadsheet(dt_od_xing, "dt_od_xing");

                                    #endregion

                                    #region transition

                                    for (int i = 0; i < dt_transition.Rows.Count; ++i)
                                    {
                                        if (dt_transition.Rows[i][col_sta] != DBNull.Value)
                                        {
                                            double sta_tra = Convert.ToDouble(dt_transition.Rows[i][col_sta]);
                                            if (sta_tra > last_sta) sta_tra = last_sta - 0.0001;

                                            for (int j = 0; j < dt_cl.Rows.Count - 1; ++j)
                                            {
                                                if (dt_cl.Rows[j]["3DSta"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", "")) == true &&
                                                    dt_cl.Rows[j + 1]["3DSta"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", "")) == true &&
                                                    dt_cl.Rows[j]["X"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["X"])) == true &&
                                                    dt_cl.Rows[j]["Y"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Y"])) == true &&
                                                    dt_cl.Rows[j]["Z"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Z"])) == true &&
                                                     dt_cl.Rows[j + 1]["X"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["X"])) == true &&
                                                    dt_cl.Rows[j + 1]["Y"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Y"])) == true &&
                                                    dt_cl.Rows[j + 1]["Z"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Z"])) == true)
                                                {
                                                    double sta1 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", ""));
                                                    double sta2 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", ""));
                                                    if (sta_tra >= sta1 && sta_tra <= sta2)
                                                    {
                                                        double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                        double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);

                                                        double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                        double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);

                                                        double x = x1 + (x2 - x1) * (sta_tra - sta1) / (sta2 - sta1);
                                                        double y = y1 + (y2 - y1) * (sta_tra - sta1) / (sta2 - sta1);

                                                        dt_transition.Rows[i]["x1"] = x;
                                                        dt_transition.Rows[i]["y1"] = y;
                                                        dt_transition.Rows[i]["layer"] = transition_layer;
                                                    }
                                                }
                                            }
                                        }

                                    }

                                    for (int i = 0; i < dt_transition.Rows.Count; ++i)
                                    {
                                        double x1 = Convert.ToDouble(dt_transition.Rows[i]["x1"]);
                                        double y1 = Convert.ToDouble(dt_transition.Rows[i]["y1"]);

                                        Point3d pt1 = Poly2D.GetClosestPointTo(new Point3d(x1, y1, 0), Vector3d.ZAxis, false);

                                        DBPoint pt_trans = new DBPoint(pt1);

                                        if (dt_transition.Rows[i]["layer"] != DBNull.Value)
                                        {
                                            string layer1 = Convert.ToString(dt_transition.Rows[i]["layer"]);
                                            Functions.Creaza_layer(layer1, 9, true);
                                            pt_trans.Layer = layer1;
                                            pt_trans.ColorIndex = 256;
                                        }
                                        BTrecord.AppendEntity(pt_trans);
                                        Trans1.AddNewlyCreatedDBObject(pt_trans, true);


                                        dt_od_trans.Rows.Add();
                                        dt_od_trans.Rows[dt_od_trans.Rows.Count - 1]["id"] = pt_trans.ObjectId;
                                        dt_od_trans.Rows[dt_od_trans.Rows.Count - 1][col_sta] = Functions.Get_chainage_from_double(Convert.ToDouble(dt_transition.Rows[i][col_sta]), "m", 1);

                                    }
                                    #endregion

                                    #region elbows
                                    if (dt_elbow != null && dt_elbow.Rows.Count > 0)
                                    {
                                        dt_elbow.Columns.Add("x1", typeof(double));
                                        dt_elbow.Columns.Add("y1", typeof(double));
                                        dt_elbow.Columns.Add("x2", typeof(double));
                                        dt_elbow.Columns.Add("y2", typeof(double));
                                        dt_elbow.Columns.Add("layer", typeof(string));
                                        dt_elbow.Columns.Add("ci", typeof(short));

                                        for (int i = 0; i < dt_elbow.Rows.Count; ++i)
                                        {
                                            if (dt_elbow.Rows[i][col_sta1] != DBNull.Value && dt_elbow.Rows[i][col_sta2] != DBNull.Value)
                                            {
                                                double staH1 = Convert.ToDouble(dt_elbow.Rows[i][col_sta1]);
                                                double staH2 = Convert.ToDouble(dt_elbow.Rows[i][col_sta2]);

                                                if (staH2 > last_sta) staH2 = last_sta - 0.0001;

                                                for (int j = 0; j < dt_cl.Rows.Count - 1; ++j)
                                                {
                                                    if (dt_cl.Rows[j]["3DSta"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", "")) == true &&
                                                        dt_cl.Rows[j + 1]["3DSta"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", "")) == true &&
                                                        dt_cl.Rows[j]["X"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["X"])) == true &&
                                                        dt_cl.Rows[j]["Y"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Y"])) == true &&
                                                        dt_cl.Rows[j]["Z"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Z"])) == true &&
                                                         dt_cl.Rows[j + 1]["X"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["X"])) == true &&
                                                        dt_cl.Rows[j + 1]["Y"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Y"])) == true &&
                                                        dt_cl.Rows[j + 1]["Z"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Z"])) == true)
                                                    {
                                                        double sta1 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", ""));
                                                        double sta2 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", ""));

                                                        if (staH1 >= sta1 && staH1 <= sta2)
                                                        {
                                                            double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                            double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);
                                                            double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                            double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);
                                                            double x = x1 + (x2 - x1) * (staH1 - sta1) / (sta2 - sta1);
                                                            double y = y1 + (y2 - y1) * (staH1 - sta1) / (sta2 - sta1);
                                                            dt_elbow.Rows[i]["x1"] = x;
                                                            dt_elbow.Rows[i]["y1"] = y;
                                                            dt_elbow.Rows[i]["layer"] = elbows_layer;
                                                        }


                                                        if (staH2 >= sta1 && staH2 <= sta2)
                                                        {
                                                            double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                            double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);
                                                            double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                            double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);
                                                            double x = x1 + (x2 - x1) * (staH2 - sta1) / (sta2 - sta1);
                                                            double y = y1 + (y2 - y1) * (staH2 - sta1) / (sta2 - sta1);
                                                            dt_elbow.Rows[i]["x2"] = x;
                                                            dt_elbow.Rows[i]["y2"] = y;
                                                            dt_elbow.Rows[i]["layer"] = elbows_layer;
                                                        }
                                                    }
                                                }
                                            }

                                            string mat1 = Convert.ToString(dt_elbow.Rows[i][col_pipe_type]);

                                            switch (mat1)
                                            {
                                                case "1":
                                                    dt_elbow.Rows[i]["ci"] = 9;
                                                    break;
                                                case "2":
                                                    dt_elbow.Rows[i]["ci"] = 7;
                                                    break;
                                                case "3":
                                                    dt_elbow.Rows[i]["ci"] = 1;
                                                    break;
                                                case "4":
                                                    dt_elbow.Rows[i]["ci"] = 2;
                                                    break;
                                                case "5":
                                                    dt_elbow.Rows[i]["ci"] = 3;
                                                    break;
                                                case "6":
                                                    dt_elbow.Rows[i]["ci"] = 4;
                                                    break;
                                                case "7":
                                                    dt_elbow.Rows[i]["ci"] = 7;
                                                    break;
                                                default:
                                                    dt_elbow.Rows[i]["ci"] = 6;
                                                    break;
                                            }

                                        }


                                        for (int i = 0; i < dt_elbow.Rows.Count; ++i)
                                        {
                                            double x1 = Convert.ToDouble(dt_elbow.Rows[i]["x1"]);
                                            double y1 = Convert.ToDouble(dt_elbow.Rows[i]["y1"]);
                                            double x2 = Convert.ToDouble(dt_elbow.Rows[i]["x2"]);
                                            double y2 = Convert.ToDouble(dt_elbow.Rows[i]["y2"]);
                                            Point3d pt1 = Poly2D.GetClosestPointTo(new Point3d(x1, y1, 0), Vector3d.ZAxis, false);
                                            Point3d pt2 = Poly2D.GetClosestPointTo(new Point3d(x2, y2, 0), Vector3d.ZAxis, false);

                                            double param1 = Poly2D.GetParameterAtPoint(pt1);
                                            double param2 = Poly2D.GetParameterAtPoint(pt2);

                                            Polyline poly_elbows = Functions.get_part_of_poly(Poly2D, param1, param2);

                                            if (dt_elbow.Rows[i]["ci"] != DBNull.Value && dt_elbow.Rows[i]["layer"] != DBNull.Value)
                                            {
                                                string layer1 = Convert.ToString(dt_elbow.Rows[i]["layer"]);
                                                Functions.Creaza_layer(layer1, 7, true);
                                                poly_elbows.Layer = layer1;
                                                poly_elbows.ColorIndex = Convert.ToInt16(dt_elbow.Rows[i]["ci"]);
                                            }
                                            BTrecord.AppendEntity(poly_elbows);
                                            Trans1.AddNewlyCreatedDBObject(poly_elbows, true);

                                            dt_od_elbow.Rows.Add();
                                            dt_od_elbow.Rows[dt_od_elbow.Rows.Count - 1]["id"] = poly_elbows.ObjectId;
                                            dt_od_elbow.Rows[dt_od_elbow.Rows.Count - 1][col_elbow_id] = dt_elbow.Rows[i][col_elbow_id];
                                            dt_od_elbow.Rows[dt_od_elbow.Rows.Count - 1][col_elbow_angle] = dt_elbow.Rows[i][col_elbow_angle];
                                            dt_od_elbow.Rows[dt_od_elbow.Rows.Count - 1][col_bend_type] = dt_elbow.Rows[i][col_bend_type];
                                            dt_od_elbow.Rows[dt_od_elbow.Rows.Count - 1][col_descr] = dt_elbow.Rows[i][col_descr];
                                            dt_od_elbow.Rows[dt_od_elbow.Rows.Count - 1][col_elbow_ref_dwg] = dt_elbow.Rows[i][col_elbow_ref_dwg];
                                            dt_od_elbow.Rows[dt_od_elbow.Rows.Count - 1][col_pipe_type] = dt_elbow.Rows[i][col_pipe_type];
                                            dt_od_elbow.Rows[dt_od_elbow.Rows.Count - 1][col_pipe_class] = dt_elbow.Rows[i][col_pipe_class];
                                            dt_od_elbow.Rows[dt_od_elbow.Rows.Count - 1][col_wt] = dt_elbow.Rows[i][col_wt];
                                            dt_od_elbow.Rows[dt_od_elbow.Rows.Count - 1][col_elbow_pi] = Functions.Get_chainage_from_double(Convert.ToDouble(dt_elbow.Rows[i][col_elbow_pi]), "m", 1);

                                            double s1 = -1.234;
                                            double s2 = -1.234;

                                            if (dt_elbow.Rows[i][col_sta1] != DBNull.Value)
                                            {
                                                s1 = Convert.ToDouble(dt_elbow.Rows[i][col_sta1]);
                                                dt_od_elbow.Rows[dt_od_elbow.Rows.Count - 1][col_start] = Functions.Get_chainage_from_double(s1, "m", 1);
                                            }
                                            else
                                            {
                                                dt_od_elbow.Rows[dt_od_elbow.Rows.Count - 1][col_start] = dt_elbow.Rows[i][col_sta1];
                                            }

                                            if (dt_elbow.Rows[i][col_sta2] != DBNull.Value)
                                            {
                                                s2 = Convert.ToDouble(dt_elbow.Rows[i][col_sta2]);
                                                dt_od_elbow.Rows[dt_od_elbow.Rows.Count - 1][col_end] = Functions.Get_chainage_from_double(s2, "m", 1);
                                            }
                                            else
                                            {
                                                dt_od_elbow.Rows[dt_od_elbow.Rows.Count - 1][col_end] = dt_elbow.Rows[i][col_sta2];
                                            }

                                            if (s1 != -1.234 && s2 != -1.234)
                                            {

                                                double extra1 = 0;

                                                if (dt_eq != null && dt_eq.Rows.Count > 0)
                                                {
                                                    for (int k = 0; k < dt_eq.Rows.Count; ++k)
                                                    {
                                                        if (dt_eq.Rows[k][col_back] != DBNull.Value && dt_eq.Rows[k][col_ahead] != DBNull.Value)
                                                        {
                                                            double back1 = Convert.ToDouble(dt_eq.Rows[k][col_back]);
                                                            double ahead1 = Convert.ToDouble(dt_eq.Rows[k][col_ahead]);

                                                            if (s1 < back1 && ahead1 < s2)
                                                            {
                                                                extra1 = extra1 + ahead1 - back1;
                                                            }

                                                            if (s1 == back1) s1 = ahead1;

                                                        }
                                                    }
                                                }

                                                dt_od_elbow.Rows[dt_od_elbow.Rows.Count - 1][col_len] = Math.Round(Convert.ToDecimal(s2) - Convert.ToDecimal(s1) - Convert.ToDecimal(extra1), 1);

                                            }

                                            dt_od_elbow.Rows[dt_od_elbow.Rows.Count - 1][col_elbow_notes] = dt_elbow.Rows[i][col_elbow_notes];
                                        }
                                    }
                                    #endregion

                                    #region facility
                                    if (dt_fab != null && dt_fab.Rows.Count > 0)
                                    {
                                        dt_fab.Columns.Add("x1", typeof(double));
                                        dt_fab.Columns.Add("y1", typeof(double));
                                        dt_fab.Columns.Add("x2", typeof(double));
                                        dt_fab.Columns.Add("y2", typeof(double));


                                        for (int i = 0; i < dt_fab.Rows.Count; ++i)
                                        {
                                            if (dt_fab.Rows[i][col_sta1] != DBNull.Value && dt_fab.Rows[i][col_sta2] != DBNull.Value)
                                            {
                                                double staH1 = Convert.ToDouble(dt_fab.Rows[i][col_sta1]);
                                                double staH2 = Convert.ToDouble(dt_fab.Rows[i][col_sta2]);

                                                if (staH2 > last_sta) staH2 = last_sta - 0.0001;

                                                for (int j = 0; j < dt_cl.Rows.Count - 1; ++j)
                                                {
                                                    if (dt_cl.Rows[j]["3DSta"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", "")) == true &&
                                                        dt_cl.Rows[j + 1]["3DSta"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", "")) == true &&
                                                        dt_cl.Rows[j]["X"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["X"])) == true &&
                                                        dt_cl.Rows[j]["Y"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Y"])) == true &&
                                                        dt_cl.Rows[j]["Z"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Z"])) == true &&
                                                         dt_cl.Rows[j + 1]["X"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["X"])) == true &&
                                                        dt_cl.Rows[j + 1]["Y"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Y"])) == true &&
                                                        dt_cl.Rows[j + 1]["Z"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Z"])) == true)
                                                    {
                                                        double sta1 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", ""));
                                                        double sta2 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", ""));

                                                        if (staH1 >= sta1 && staH1 <= sta2)
                                                        {
                                                            double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                            double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);
                                                            double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                            double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);
                                                            double x = x1 + (x2 - x1) * (staH1 - sta1) / (sta2 - sta1);
                                                            double y = y1 + (y2 - y1) * (staH1 - sta1) / (sta2 - sta1);
                                                            dt_fab.Rows[i]["x1"] = x;
                                                            dt_fab.Rows[i]["y1"] = y;

                                                        }


                                                        if (staH2 >= sta1 && staH2 <= sta2)
                                                        {
                                                            double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                            double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);
                                                            double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                            double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);
                                                            double x = x1 + (x2 - x1) * (staH2 - sta1) / (sta2 - sta1);
                                                            double y = y1 + (y2 - y1) * (staH2 - sta1) / (sta2 - sta1);
                                                            dt_fab.Rows[i]["x2"] = x;
                                                            dt_fab.Rows[i]["y2"] = y;

                                                        }
                                                    }
                                                }
                                            }


                                        }

                                        Functions.Creaza_layer(fab_layer, 7, true);
                                        for (int i = 0; i < dt_fab.Rows.Count; ++i)
                                        {
                                            double x1 = Convert.ToDouble(dt_fab.Rows[i]["x1"]);
                                            double y1 = Convert.ToDouble(dt_fab.Rows[i]["y1"]);
                                            double x2 = Convert.ToDouble(dt_fab.Rows[i]["x2"]);
                                            double y2 = Convert.ToDouble(dt_fab.Rows[i]["y2"]);
                                            Point3d pt1 = Poly2D.GetClosestPointTo(new Point3d(x1, y1, 0), Vector3d.ZAxis, false);
                                            Point3d pt2 = Poly2D.GetClosestPointTo(new Point3d(x2, y2, 0), Vector3d.ZAxis, false);

                                            double param1 = Poly2D.GetParameterAtPoint(pt1);
                                            double param2 = Poly2D.GetParameterAtPoint(pt2);

                                            Polyline poly_fab = Functions.get_part_of_poly(Poly2D, param1, param2);
                                            poly_fab.Layer = fab_layer;
                                            BTrecord.AppendEntity(poly_fab);
                                            Trans1.AddNewlyCreatedDBObject(poly_fab, true);

                                            dt_od_fab.Rows.Add();
                                            dt_od_fab.Rows[dt_od_fab.Rows.Count - 1]["id"] = poly_fab.ObjectId;
                                            dt_od_fab.Rows[dt_od_fab.Rows.Count - 1][col_sta1] = Functions.Get_chainage_from_double(Convert.ToDouble(dt_fab.Rows[i][col_sta1]), "m", 1);
                                            dt_od_fab.Rows[dt_od_fab.Rows.Count - 1][col_sta2] = Functions.Get_chainage_from_double(Convert.ToDouble(dt_fab.Rows[i][col_sta2]), "m", 1);
                                            dt_od_fab.Rows[dt_od_fab.Rows.Count - 1][col_just] = dt_fab.Rows[i][col_just];
                                            dt_od_fab.Rows[dt_od_fab.Rows.Count - 1][col_fac_name] = dt_fab.Rows[i][col_fac_name];
                                            dt_od_fab.Rows[dt_od_fab.Rows.Count - 1][col_descr1] = dt_fab.Rows[i][col_descr1];
                                            dt_od_fab.Rows[dt_od_fab.Rows.Count - 1][col_descr2] = dt_fab.Rows[i][col_descr2];


                                            double extra1 = 0;

                                            double s1 = Convert.ToDouble(dt_fab.Rows[i][col_sta1]);
                                            double s2 = Convert.ToDouble(dt_fab.Rows[i][col_sta2]);

                                            if (dt_eq != null && dt_eq.Rows.Count > 0)
                                            {
                                                for (int k = 0; k < dt_eq.Rows.Count; ++k)
                                                {
                                                    if (dt_eq.Rows[k][col_back] != DBNull.Value && dt_eq.Rows[k][col_ahead] != DBNull.Value)
                                                    {
                                                        double back1 = Convert.ToDouble(dt_eq.Rows[k][col_back]);
                                                        double ahead1 = Convert.ToDouble(dt_eq.Rows[k][col_ahead]);

                                                        if (s1 < back1 && ahead1 < s2)
                                                        {
                                                            extra1 = extra1 + ahead1 - back1;
                                                        }

                                                        if (s1 == back1) s1 = ahead1;

                                                    }
                                                }
                                            }

                                            dt_od_fab.Rows[dt_od_fab.Rows.Count - 1][col_len] = Math.Round(Convert.ToDecimal(s2) - Convert.ToDecimal(s1) - Convert.ToDecimal(extra1), 1);


                                        }
                                    }
                                    #endregion

                                    #region class
                                    if (dt_class != null && dt_class.Rows.Count > 0)
                                    {
                                        dt_class.Columns.Add("x1", typeof(double));
                                        dt_class.Columns.Add("y1", typeof(double));
                                        dt_class.Columns.Add("x2", typeof(double));
                                        dt_class.Columns.Add("y2", typeof(double));
                                        dt_class.Columns.Add("layer", typeof(string));
                                        dt_class.Columns.Add("ci", typeof(short));

                                        for (int i = 0; i < dt_class.Rows.Count; ++i)
                                        {
                                            if (dt_class.Rows[i][col_sta1] != DBNull.Value && dt_class.Rows[i][col_sta2] != DBNull.Value)
                                            {
                                                double staH1 = Convert.ToDouble(dt_class.Rows[i][col_sta1]);
                                                double staH2 = Convert.ToDouble(dt_class.Rows[i][col_sta2]);

                                                if (staH2 > last_sta) staH2 = last_sta - 0.0001;

                                                for (int j = 0; j < dt_cl.Rows.Count - 1; ++j)
                                                {
                                                    if (dt_cl.Rows[j]["3DSta"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", "")) == true &&
                                                        dt_cl.Rows[j + 1]["3DSta"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", "")) == true &&
                                                        dt_cl.Rows[j]["X"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["X"])) == true &&
                                                        dt_cl.Rows[j]["Y"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Y"])) == true &&
                                                        dt_cl.Rows[j]["Z"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Z"])) == true &&
                                                         dt_cl.Rows[j + 1]["X"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["X"])) == true &&
                                                        dt_cl.Rows[j + 1]["Y"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Y"])) == true &&
                                                        dt_cl.Rows[j + 1]["Z"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Z"])) == true)
                                                    {
                                                        double sta1 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", ""));
                                                        double sta2 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", ""));

                                                        if (staH1 >= sta1 && staH1 <= sta2)
                                                        {
                                                            double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                            double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);
                                                            double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                            double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);
                                                            double x = x1 + (x2 - x1) * (staH1 - sta1) / (sta2 - sta1);
                                                            double y = y1 + (y2 - y1) * (staH1 - sta1) / (sta2 - sta1);
                                                            dt_class.Rows[i]["x1"] = x;
                                                            dt_class.Rows[i]["y1"] = y;
                                                            dt_class.Rows[i]["layer"] = class_layer;
                                                        }


                                                        if (staH2 >= sta1 && staH2 <= sta2)
                                                        {
                                                            double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                            double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);
                                                            double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                            double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);
                                                            double x = x1 + (x2 - x1) * (staH2 - sta1) / (sta2 - sta1);
                                                            double y = y1 + (y2 - y1) * (staH2 - sta1) / (sta2 - sta1);
                                                            dt_class.Rows[i]["x2"] = x;
                                                            dt_class.Rows[i]["y2"] = y;
                                                            dt_class.Rows[i]["layer"] = class_layer;
                                                        }
                                                    }
                                                }
                                            }

                                            string mat1 = Convert.ToString(dt_class.Rows[i][col_pipe_type]);

                                            switch (mat1)
                                            {
                                                case "1":
                                                    dt_class.Rows[i]["ci"] = 9;
                                                    break;
                                                case "2":
                                                    dt_class.Rows[i]["ci"] = 7;
                                                    break;
                                                case "3":
                                                    dt_class.Rows[i]["ci"] = 1;
                                                    break;
                                                case "4":
                                                    dt_class.Rows[i]["ci"] = 2;
                                                    break;
                                                case "5":
                                                    dt_class.Rows[i]["ci"] = 3;
                                                    break;
                                                case "6":
                                                    dt_class.Rows[i]["ci"] = 4;
                                                    break;
                                                case "7":
                                                    dt_class.Rows[i]["ci"] = 7;
                                                    break;
                                                default:
                                                    dt_class.Rows[i]["ci"] = 6;
                                                    break;
                                            }

                                        }


                                        for (int i = 0; i < dt_class.Rows.Count; ++i)
                                        {
                                            double x1 = Convert.ToDouble(dt_class.Rows[i]["x1"]);
                                            double y1 = Convert.ToDouble(dt_class.Rows[i]["y1"]);
                                            double x2 = Convert.ToDouble(dt_class.Rows[i]["x2"]);
                                            double y2 = Convert.ToDouble(dt_class.Rows[i]["y2"]);
                                            Point3d pt1 = Poly2D.GetClosestPointTo(new Point3d(x1, y1, 0), Vector3d.ZAxis, false);
                                            Point3d pt2 = Poly2D.GetClosestPointTo(new Point3d(x2, y2, 0), Vector3d.ZAxis, false);

                                            double param1 = Poly2D.GetParameterAtPoint(pt1);
                                            double param2 = Poly2D.GetParameterAtPoint(pt2);

                                            Polyline poly_class = Functions.get_part_of_poly(Poly2D, param1, param2);

                                            if (dt_class.Rows[i]["ci"] != DBNull.Value && dt_class.Rows[i]["layer"] != DBNull.Value)
                                            {
                                                string layer1 = Convert.ToString(dt_class.Rows[i]["layer"]);
                                                Functions.Creaza_layer(layer1, 7, true);
                                                poly_class.Layer = layer1;
                                                poly_class.ColorIndex = Convert.ToInt16(dt_class.Rows[i]["ci"]);
                                            }
                                            BTrecord.AppendEntity(poly_class);
                                            Trans1.AddNewlyCreatedDBObject(poly_class, true);

                                            dt_od_class.Rows.Add();
                                            dt_od_class.Rows[dt_od_class.Rows.Count - 1]["id"] = poly_class.ObjectId;
                                            dt_od_class.Rows[dt_od_class.Rows.Count - 1][col_sta1] = Functions.Get_chainage_from_double(Convert.ToDouble(dt_class.Rows[i][col_sta1]), "m", 1);
                                            dt_od_class.Rows[dt_od_class.Rows.Count - 1][col_sta2] = Functions.Get_chainage_from_double(Convert.ToDouble(dt_class.Rows[i][col_sta2]), "m", 1);
                                            dt_od_class.Rows[dt_od_class.Rows.Count - 1][col_od_pipe_type] = dt_class.Rows[i][col_pipe_type];
                                            dt_od_class.Rows[dt_od_class.Rows.Count - 1][col_wt] = dt_class.Rows[i][col_wt];
                                            dt_od_class.Rows[dt_od_class.Rows.Count - 1][col_just] = dt_class.Rows[i][col_descr];
                                            dt_od_class.Rows[dt_od_class.Rows.Count - 1][col_descr1] = dt_class.Rows[i][col_descr1];
                                            dt_od_class.Rows[dt_od_class.Rows.Count - 1][col_descr2] = dt_class.Rows[i][col_descr2];

                                        }
                                    }
                                    #endregion

                                    #region Geotech

                                    for (int i = 0; i < dt_geotech.Rows.Count; ++i)
                                    {
                                        if (dt_geotech.Rows[i][col_geotech_sta1] != DBNull.Value && dt_geotech.Rows[i][col_geotech_sta2] != DBNull.Value)
                                        {
                                            double staH1 = Convert.ToDouble(dt_geotech.Rows[i][col_geotech_sta1]);
                                            double staH2 = Convert.ToDouble(dt_geotech.Rows[i][col_geotech_sta2]);

                                            if (staH2 > last_sta) staH2 = last_sta - 0.0001;

                                            for (int j = 0; j < dt_cl.Rows.Count - 1; ++j)
                                            {
                                                if (dt_cl.Rows[j]["3DSta"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", "")) == true &&
                                                    dt_cl.Rows[j + 1]["3DSta"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", "")) == true &&
                                                    dt_cl.Rows[j]["X"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["X"])) == true &&
                                                    dt_cl.Rows[j]["Y"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Y"])) == true &&
                                                    dt_cl.Rows[j]["Z"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Z"])) == true &&
                                                     dt_cl.Rows[j + 1]["X"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["X"])) == true &&
                                                    dt_cl.Rows[j + 1]["Y"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Y"])) == true &&
                                                    dt_cl.Rows[j + 1]["Z"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Z"])) == true)
                                                {
                                                    double sta1 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", ""));
                                                    double sta2 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", ""));
                                                    if (staH1 >= sta1 && staH1 <= sta2)
                                                    {
                                                        double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                        double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);

                                                        double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                        double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);

                                                        double x = x1 + (x2 - x1) * (staH1 - sta1) / (sta2 - sta1);
                                                        double y = y1 + (y2 - y1) * (staH1 - sta1) / (sta2 - sta1);

                                                        dt_geotech.Rows[i]["x1"] = x;
                                                        dt_geotech.Rows[i]["y1"] = y;
                                                        dt_geotech.Rows[i]["layer"] = geotech_layer;
                                                    }

                                                    if (staH2 >= sta1 && staH2 <= sta2)
                                                    {
                                                        double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                        double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);
                                                        double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                        double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);
                                                        double x = x1 + (x2 - x1) * (staH2 - sta1) / (sta2 - sta1);
                                                        double y = y1 + (y2 - y1) * (staH2 - sta1) / (sta2 - sta1);
                                                        dt_geotech.Rows[i]["x2"] = x;
                                                        dt_geotech.Rows[i]["y2"] = y;
                                                        dt_geotech.Rows[i]["layer"] = geotech_layer;
                                                    }
                                                }
                                            }
                                        }

                                    }


                                    Functions.Creaza_layer(geotech_layer_pt, 7, true);

                                    for (int i = 0; i < dt_geotech.Rows.Count; ++i)
                                    {
                                        double x1 = Convert.ToDouble(dt_geotech.Rows[i]["x1"]);
                                        double y1 = Convert.ToDouble(dt_geotech.Rows[i]["y1"]);
                                        double x2 = Convert.ToDouble(dt_geotech.Rows[i]["x2"]);
                                        double y2 = Convert.ToDouble(dt_geotech.Rows[i]["y2"]);
                                        Point3d pt1 = Poly2D.GetClosestPointTo(new Point3d(x1, y1, 0), Vector3d.ZAxis, false);
                                        Point3d pt2 = Poly2D.GetClosestPointTo(new Point3d(x2, y2, 0), Vector3d.ZAxis, false);



                                        double param1 = Poly2D.GetParameterAtPoint(pt1);
                                        double param2 = Poly2D.GetParameterAtPoint(pt2);

                                        DBPoint pt_geotech1 = new DBPoint(pt1);
                                        DBPoint pt_geotech2 = new DBPoint(pt2);


                                        Polyline geotech = Functions.get_part_of_poly(Poly2D, param1, param2);

                                        if (dt_geotech.Rows[i]["layer"] != DBNull.Value)
                                        {
                                            string layer1 = Convert.ToString(dt_geotech.Rows[i]["layer"]);
                                            Functions.Creaza_layer(layer1, 7, true);
                                            geotech.Layer = layer1;
                                            geotech.ColorIndex = 256;

                                            pt_geotech1.Layer = geotech_layer_pt;
                                            pt_geotech1.ColorIndex = 256;
                                            pt_geotech2.Layer = geotech_layer_pt;
                                            pt_geotech2.ColorIndex = 256;

                                        }
                                        BTrecord.AppendEntity(geotech);
                                        Trans1.AddNewlyCreatedDBObject(geotech, true);
                                        BTrecord.AppendEntity(pt_geotech1);
                                        Trans1.AddNewlyCreatedDBObject(pt_geotech1, true);
                                        BTrecord.AppendEntity(pt_geotech2);
                                        Trans1.AddNewlyCreatedDBObject(pt_geotech2, true);






                                        dt_od_geotech.Rows.Add();
                                        dt_od_geotech.Rows[dt_od_geotech.Rows.Count - 1]["id"] = geotech.ObjectId;
                                        dt_od_geotech.Rows[dt_od_geotech.Rows.Count - 1]["id1"] = pt_geotech1.ObjectId;
                                        dt_od_geotech.Rows[dt_od_geotech.Rows.Count - 1]["id2"] = pt_geotech2.ObjectId;
                                        dt_od_geotech.Rows[dt_od_geotech.Rows.Count - 1][col_geotech_od_class] = dt_geotech.Rows[i][col_geotech_class];
                                        dt_od_geotech.Rows[dt_od_geotech.Rows.Count - 1][col_geotech_od_descr1] = dt_geotech.Rows[i][col_geotech_descr1];
                                        dt_od_geotech.Rows[dt_od_geotech.Rows.Count - 1][col_geotech_od_descr2] = dt_geotech.Rows[i][col_geotech_descr2];
                                        dt_od_geotech.Rows[dt_od_geotech.Rows.Count - 1][col_geotech_od_label] = dt_geotech.Rows[i][col_geotech_label];
                                        dt_od_geotech.Rows[dt_od_geotech.Rows.Count - 1][col_geotech_od_type] = dt_geotech.Rows[i][col_geotech_type];
                                        dt_od_geotech.Rows[dt_od_geotech.Rows.Count - 1][col_notes] = dt_geotech.Rows[i][col_notes];




                                        double s1 = -1.234;
                                        double s2 = -1.234;

                                        if (dt_geotech.Rows[i][col_geotech_sta1] != DBNull.Value)
                                        {
                                            s1 = Convert.ToDouble(dt_geotech.Rows[i][col_geotech_sta1]);
                                            dt_od_geotech.Rows[dt_od_geotech.Rows.Count - 1][col_geotech_od_sta1] = Functions.Get_chainage_from_double(s1, "m", 1);
                                        }
                                        else
                                        {
                                            dt_od_geotech.Rows[dt_od_geotech.Rows.Count - 1][col_geotech_od_sta1] = dt_geotech.Rows[i][col_geotech_sta1];
                                        }

                                        if (dt_geotech.Rows[i][col_geotech_sta2] != DBNull.Value)
                                        {
                                            s2 = Convert.ToDouble(dt_geotech.Rows[i][col_geotech_sta2]);
                                            dt_od_geotech.Rows[dt_od_geotech.Rows.Count - 1][col_geotech_od_sta2] = Functions.Get_chainage_from_double(s2, "m", 1);
                                        }
                                        else
                                        {
                                            dt_od_geotech.Rows[dt_od_geotech.Rows.Count - 1][col_geotech_od_sta2] = dt_geotech.Rows[i][col_geotech_sta2];
                                        }

                                        if (s1 != -1.234 && s2 != -1.234)
                                        {

                                            double extra1 = 0;

                                            if (dt_eq != null && dt_eq.Rows.Count > 0)
                                            {
                                                for (int k = 0; k < dt_eq.Rows.Count; ++k)
                                                {
                                                    if (dt_eq.Rows[k][col_back] != DBNull.Value && dt_eq.Rows[k][col_ahead] != DBNull.Value)
                                                    {
                                                        double back1 = Convert.ToDouble(dt_eq.Rows[k][col_back]);
                                                        double ahead1 = Convert.ToDouble(dt_eq.Rows[k][col_ahead]);

                                                        if (s1 < back1 && ahead1 < s2)
                                                        {
                                                            extra1 = extra1 + ahead1 - back1;
                                                        }

                                                        if (s1 == back1) s1 = ahead1;

                                                    }
                                                }
                                            }

                                            dt_od_geotech.Rows[dt_od_geotech.Rows.Count - 1][col_len] = Math.Round(Convert.ToDecimal(s2) - Convert.ToDecimal(s1) - Convert.ToDecimal(extra1), 1);

                                        }
                                    }
                                    #endregion

                                    #region depth of cover
                                    if (dt_doc != null && dt_doc.Rows.Count > 0)
                                    {
                                        for (int i = 0; i < dt_doc.Rows.Count; ++i)
                                        {
                                            if (dt_doc.Rows[i][col_doc_sta1] != DBNull.Value && dt_doc.Rows[i][col_doc_sta2] != DBNull.Value)
                                            {
                                                double staH1 = Convert.ToDouble(dt_doc.Rows[i][col_doc_sta1]);
                                                double staH2 = Convert.ToDouble(dt_doc.Rows[i][col_doc_sta2]);

                                                if (staH2 > last_sta) staH2 = last_sta - 0.0001;

                                                for (int j = 0; j < dt_cl.Rows.Count - 1; ++j)
                                                {
                                                    if (dt_cl.Rows[j]["3DSta"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", "")) == true &&
                                                        dt_cl.Rows[j + 1]["3DSta"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", "")) == true &&
                                                        dt_cl.Rows[j]["X"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["X"])) == true &&
                                                        dt_cl.Rows[j]["Y"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Y"])) == true &&
                                                        dt_cl.Rows[j]["Z"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Z"])) == true &&
                                                         dt_cl.Rows[j + 1]["X"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["X"])) == true &&
                                                        dt_cl.Rows[j + 1]["Y"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Y"])) == true &&
                                                        dt_cl.Rows[j + 1]["Z"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Z"])) == true)
                                                    {
                                                        double sta1 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", ""));
                                                        double sta2 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", ""));

                                                        if (staH1 >= sta1 && staH1 <= sta2)
                                                        {
                                                            double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                            double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);
                                                            double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                            double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);
                                                            double x = x1 + (x2 - x1) * (staH1 - sta1) / (sta2 - sta1);
                                                            double y = y1 + (y2 - y1) * (staH1 - sta1) / (sta2 - sta1);
                                                            dt_doc.Rows[i]["x1"] = x;
                                                            dt_doc.Rows[i]["y1"] = y;
                                                            dt_doc.Rows[i]["layer"] = doc_layer;
                                                        }


                                                        if (staH2 >= sta1 && staH2 <= sta2)
                                                        {
                                                            double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                            double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);
                                                            double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                            double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);
                                                            double x = x1 + (x2 - x1) * (staH2 - sta1) / (sta2 - sta1);
                                                            double y = y1 + (y2 - y1) * (staH2 - sta1) / (sta2 - sta1);
                                                            dt_doc.Rows[i]["x2"] = x;
                                                            dt_doc.Rows[i]["y2"] = y;
                                                            dt_doc.Rows[i]["layer"] = doc_layer;
                                                        }
                                                    }
                                                }
                                            }



                                        }


                                        for (int i = 0; i < dt_doc.Rows.Count; ++i)
                                        {
                                            double x1 = Convert.ToDouble(dt_doc.Rows[i]["x1"]);
                                            double y1 = Convert.ToDouble(dt_doc.Rows[i]["y1"]);
                                            double x2 = Convert.ToDouble(dt_doc.Rows[i]["x2"]);
                                            double y2 = Convert.ToDouble(dt_doc.Rows[i]["y2"]);
                                            Point3d pt1 = Poly2D.GetClosestPointTo(new Point3d(x1, y1, 0), Vector3d.ZAxis, false);
                                            Point3d pt2 = Poly2D.GetClosestPointTo(new Point3d(x2, y2, 0), Vector3d.ZAxis, false);

                                            double param1 = Poly2D.GetParameterAtPoint(pt1);
                                            double param2 = Poly2D.GetParameterAtPoint(pt2);

                                            Polyline poly_doc = Functions.get_part_of_poly(Poly2D, param1, param2);

                                            if (dt_doc.Rows[i]["layer"] != DBNull.Value)
                                            {
                                                string layer1 = Convert.ToString(dt_doc.Rows[i]["layer"]);
                                                Functions.Creaza_layer(layer1, 7, true);
                                                poly_doc.Layer = layer1;
                                                poly_doc.ColorIndex = 256;
                                            }
                                            BTrecord.AppendEntity(poly_doc);
                                            Trans1.AddNewlyCreatedDBObject(poly_doc, true);

                                            dt_od_doc.Rows.Add();
                                            dt_od_doc.Rows[dt_od_doc.Rows.Count - 1]["id"] = poly_doc.ObjectId;
                                            dt_od_doc.Rows[dt_od_doc.Rows.Count - 1][col_doc_od_min_cvr] = dt_doc.Rows[i][col_doc_min_cvr];
                                            dt_od_doc.Rows[dt_od_doc.Rows.Count - 1][col_just] = dt_doc.Rows[i][col_just];
                                            dt_od_doc.Rows[dt_od_doc.Rows.Count - 1][col_notes] = dt_doc.Rows[i][col_notes];

                                            double s1 = -1.234;
                                            double s2 = -1.234;

                                            if (dt_doc.Rows[i][col_doc_sta1] != DBNull.Value)
                                            {
                                                s1 = Convert.ToDouble(dt_doc.Rows[i][col_doc_sta1]);
                                                dt_od_doc.Rows[dt_od_doc.Rows.Count - 1][col_doc_od_sta1] = Functions.Get_chainage_from_double(s1, "m", 1);
                                            }
                                            else
                                            {
                                                dt_od_doc.Rows[dt_od_doc.Rows.Count - 1][col_doc_od_sta1] = dt_doc.Rows[i][col_doc_sta1];
                                            }

                                            if (dt_doc.Rows[i][col_doc_sta2] != DBNull.Value)
                                            {
                                                s2 = Convert.ToDouble(dt_doc.Rows[i][col_doc_sta2]);
                                                dt_od_doc.Rows[dt_od_doc.Rows.Count - 1][col_doc_od_sta2] = Functions.Get_chainage_from_double(s2, "m", 1);
                                            }
                                            else
                                            {
                                                dt_od_doc.Rows[dt_od_doc.Rows.Count - 1][col_doc_od_sta2] = dt_doc.Rows[i][col_doc_sta2];
                                            }

                                            if (s1 != -1.234 && s2 != -1.234)
                                            {

                                                double extra1 = 0;

                                                if (dt_eq != null && dt_eq.Rows.Count > 0)
                                                {
                                                    for (int k = 0; k < dt_eq.Rows.Count; ++k)
                                                    {
                                                        if (dt_eq.Rows[k][col_back] != DBNull.Value && dt_eq.Rows[k][col_ahead] != DBNull.Value)
                                                        {
                                                            double back1 = Convert.ToDouble(dt_eq.Rows[k][col_back]);
                                                            double ahead1 = Convert.ToDouble(dt_eq.Rows[k][col_ahead]);

                                                            if (s1 < back1 && ahead1 < s2)
                                                            {
                                                                extra1 = extra1 + ahead1 - back1;
                                                            }

                                                            if (s1 == back1) s1 = ahead1;

                                                        }
                                                    }
                                                }

                                                dt_od_doc.Rows[dt_od_doc.Rows.Count - 1][col_len] = Math.Round(Convert.ToDecimal(s2) - Convert.ToDecimal(s1) - Convert.ToDecimal(extra1), 1);

                                            }



                                        }
                                    }
                                    #endregion

                                    #region muskeg
                                    for (int i = 0; i < dt_muskeg.Rows.Count; ++i)
                                    {
                                        if (dt_muskeg.Rows[i][col_muskeg_sta1] != DBNull.Value && dt_muskeg.Rows[i][col_muskeg_sta2] != DBNull.Value)
                                        {
                                            double sta_muskeg1 = Convert.ToDouble(dt_muskeg.Rows[i][col_muskeg_sta1]);
                                            double sta_muskeg2 = Convert.ToDouble(dt_muskeg.Rows[i][col_muskeg_sta2]);

                                            if (sta_muskeg2 > last_sta) sta_muskeg2 = last_sta - 0.0001;

                                            for (int j = 0; j < dt_cl.Rows.Count - 1; ++j)
                                            {
                                                if (dt_cl.Rows[j]["3DSta"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", "")) == true &&
                                                    dt_cl.Rows[j + 1]["3DSta"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", "")) == true &&
                                                    dt_cl.Rows[j]["X"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["X"])) == true &&
                                                    dt_cl.Rows[j]["Y"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Y"])) == true &&
                                                    dt_cl.Rows[j]["Z"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Z"])) == true &&
                                                     dt_cl.Rows[j + 1]["X"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["X"])) == true &&
                                                    dt_cl.Rows[j + 1]["Y"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Y"])) == true &&
                                                    dt_cl.Rows[j + 1]["Z"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Z"])) == true)
                                                {
                                                    double sta1 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", ""));
                                                    double sta2 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", ""));
                                                    if (sta_muskeg1 >= sta1 && sta_muskeg1 <= sta2)
                                                    {
                                                        double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                        double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);

                                                        double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                        double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);

                                                        double x = x1 + (x2 - x1) * (sta_muskeg1 - sta1) / (sta2 - sta1);
                                                        double y = y1 + (y2 - y1) * (sta_muskeg1 - sta1) / (sta2 - sta1);

                                                        dt_muskeg.Rows[i]["x1"] = x;
                                                        dt_muskeg.Rows[i]["y1"] = y;
                                                        dt_muskeg.Rows[i]["layer"] = muskeg_layer;

                                                    }
                                                    if (sta_muskeg2 >= sta1 && sta_muskeg2 <= sta2)
                                                    {
                                                        double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                        double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);

                                                        double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                        double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);

                                                        double x = x1 + (x2 - x1) * (sta_muskeg2 - sta1) / (sta2 - sta1);
                                                        double y = y1 + (y2 - y1) * (sta_muskeg2 - sta1) / (sta2 - sta1);

                                                        dt_muskeg.Rows[i]["x2"] = x;
                                                        dt_muskeg.Rows[i]["y2"] = y;
                                                        dt_muskeg.Rows[i]["layer"] = muskeg_layer;

                                                    }

                                                }
                                            }
                                        }

                                    }

                                    for (int i = 0; i < dt_muskeg.Rows.Count; ++i)
                                    {
                                        double x1 = Convert.ToDouble(dt_muskeg.Rows[i]["x1"]);
                                        double y1 = Convert.ToDouble(dt_muskeg.Rows[i]["y1"]);

                                        double x2 = Convert.ToDouble(dt_muskeg.Rows[i]["x2"]);
                                        double y2 = Convert.ToDouble(dt_muskeg.Rows[i]["y2"]);

                                        Point3d pt1 = Poly2D.GetClosestPointTo(new Point3d(x1, y1, 0), Vector3d.ZAxis, false);
                                        Point3d pt2 = Poly2D.GetClosestPointTo(new Point3d(x2, y2, 0), Vector3d.ZAxis, false);

                                        DBPoint pt_muskeg1 = new DBPoint(pt1);
                                        DBPoint pt_muskeg2 = new DBPoint(pt2);

                                        if (dt_muskeg.Rows[i]["layer"] != DBNull.Value)
                                        {
                                            string layer1 = Convert.ToString(dt_muskeg.Rows[i]["layer"]);
                                            Functions.Creaza_layer(layer1, 4, true);
                                            pt_muskeg1.Layer = layer1;
                                            pt_muskeg1.ColorIndex = 256;
                                            pt_muskeg2.Layer = layer1;
                                            pt_muskeg2.ColorIndex = 256;
                                        }
                                        BTrecord.AppendEntity(pt_muskeg1);
                                        Trans1.AddNewlyCreatedDBObject(pt_muskeg1, true);
                                        BTrecord.AppendEntity(pt_muskeg2);
                                        Trans1.AddNewlyCreatedDBObject(pt_muskeg2, true);

                                        dt_od_muskeg.Rows.Add();
                                        dt_od_muskeg.Rows[dt_od_muskeg.Rows.Count - 1]["id"] = pt_muskeg1.ObjectId;
                                        dt_od_muskeg.Rows[dt_od_muskeg.Rows.Count - 1][col_muskeg_od_sta] = Functions.Get_chainage_from_double(Convert.ToDouble(dt_muskeg.Rows[i][col_muskeg_sta1]), "m", 1);
                                        dt_od_muskeg.Rows[dt_od_muskeg.Rows.Count - 1][col_muskeg_od_descr] = dt_muskeg.Rows[i][col_muskeg_descr1];
                                        dt_od_muskeg.Rows[dt_od_muskeg.Rows.Count - 1][col_muskeg_od_label] = dt_muskeg.Rows[i][col_muskeg_label];

                                        dt_od_muskeg.Rows.Add();
                                        dt_od_muskeg.Rows[dt_od_muskeg.Rows.Count - 1]["id"] = pt_muskeg2.ObjectId;
                                        dt_od_muskeg.Rows[dt_od_muskeg.Rows.Count - 1][col_muskeg_od_sta] = Functions.Get_chainage_from_double(Convert.ToDouble(dt_muskeg.Rows[i][col_muskeg_sta2]), "m", 1);
                                        dt_od_muskeg.Rows[dt_od_muskeg.Rows.Count - 1][col_muskeg_od_descr] = dt_muskeg.Rows[i][col_muskeg_descr2];
                                        dt_od_muskeg.Rows[dt_od_muskeg.Rows.Count - 1][col_muskeg_od_label] = dt_muskeg.Rows[i][col_muskeg_label];


                                    }
                                    #endregion

                                    #region pre_existing pipes (treated as facilities)
                                    if (dt_pre_existing != null && dt_pre_existing.Rows.Count > 0)
                                    {
                                        dt_pre_existing.Columns.Add("x1", typeof(double));
                                        dt_pre_existing.Columns.Add("y1", typeof(double));
                                        dt_pre_existing.Columns.Add("x2", typeof(double));
                                        dt_pre_existing.Columns.Add("y2", typeof(double));

                                        for (int i = 0; i < dt_pre_existing.Rows.Count; ++i)
                                        {
                                            if (dt_pre_existing.Rows[i][col_start] != DBNull.Value && dt_pre_existing.Rows[i][col_end] != DBNull.Value)
                                            {
                                                double staH1 = Convert.ToDouble(dt_pre_existing.Rows[i][col_start]);
                                                double staH2 = Convert.ToDouble(dt_pre_existing.Rows[i][col_end]);

                                                if (staH2 > last_sta) staH2 = last_sta - 0.0001;

                                                for (int j = 0; j < dt_cl.Rows.Count - 1; ++j)
                                                {
                                                    if (dt_cl.Rows[j]["3DSta"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", "")) == true &&
                                                        dt_cl.Rows[j + 1]["3DSta"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", "")) == true &&
                                                        dt_cl.Rows[j]["X"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["X"])) == true &&
                                                        dt_cl.Rows[j]["Y"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Y"])) == true &&
                                                        dt_cl.Rows[j]["Z"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Z"])) == true &&
                                                         dt_cl.Rows[j + 1]["X"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["X"])) == true &&
                                                        dt_cl.Rows[j + 1]["Y"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Y"])) == true &&
                                                        dt_cl.Rows[j + 1]["Z"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Z"])) == true)
                                                    {
                                                        double sta1 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", ""));
                                                        double sta2 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", ""));

                                                        if (staH1 >= sta1 && staH1 <= sta2)
                                                        {
                                                            double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                            double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);
                                                            double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                            double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);
                                                            double x = x1 + (x2 - x1) * (staH1 - sta1) / (sta2 - sta1);
                                                            double y = y1 + (y2 - y1) * (staH1 - sta1) / (sta2 - sta1);
                                                            dt_pre_existing.Rows[i]["x1"] = x;
                                                            dt_pre_existing.Rows[i]["y1"] = y;
                                                        }


                                                        if (staH2 >= sta1 && staH2 <= sta2)
                                                        {
                                                            double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                            double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);
                                                            double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                            double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);
                                                            double x = x1 + (x2 - x1) * (staH2 - sta1) / (sta2 - sta1);
                                                            double y = y1 + (y2 - y1) * (staH2 - sta1) / (sta2 - sta1);
                                                            dt_pre_existing.Rows[i]["x2"] = x;
                                                            dt_pre_existing.Rows[i]["y2"] = y;
                                                        }
                                                    }
                                                }
                                            }



                                        }
                                        Functions.Creaza_layer(preexisting_layer, 7, true);

                                        for (int i = 0; i < dt_pre_existing.Rows.Count; ++i)
                                        {
                                            double x1 = Convert.ToDouble(dt_pre_existing.Rows[i]["x1"]);
                                            double y1 = Convert.ToDouble(dt_pre_existing.Rows[i]["y1"]);
                                            double x2 = Convert.ToDouble(dt_pre_existing.Rows[i]["x2"]);
                                            double y2 = Convert.ToDouble(dt_pre_existing.Rows[i]["y2"]);
                                            Point3d pt1 = Poly2D.GetClosestPointTo(new Point3d(x1, y1, 0), Vector3d.ZAxis, false);
                                            Point3d pt2 = Poly2D.GetClosestPointTo(new Point3d(x2, y2, 0), Vector3d.ZAxis, false);

                                            double param1 = Poly2D.GetParameterAtPoint(pt1);
                                            double param2 = Poly2D.GetParameterAtPoint(pt2);

                                            Polyline poly_fab = Functions.get_part_of_poly(Poly2D, param1, param2);
                                            poly_fab.Layer = preexisting_layer;
                                            BTrecord.AppendEntity(poly_fab);
                                            Trans1.AddNewlyCreatedDBObject(poly_fab, true);

                                            dt_od_pre_existing.Rows.Add();
                                            dt_od_pre_existing.Rows[dt_od_pre_existing.Rows.Count - 1]["id"] = poly_fab.ObjectId;
                                            dt_od_pre_existing.Rows[dt_od_pre_existing.Rows.Count - 1][col_start] = Functions.Get_chainage_from_double(Convert.ToDouble(dt_pre_existing.Rows[i][col_start]), "m", 1);
                                            dt_od_pre_existing.Rows[dt_od_pre_existing.Rows.Count - 1][col_end] = Functions.Get_chainage_from_double(Convert.ToDouble(dt_pre_existing.Rows[i][col_end]), "m", 1);
                                            dt_od_pre_existing.Rows[dt_od_pre_existing.Rows.Count - 1][col_descr] = dt_pre_existing.Rows[i][col_descr];
                                            dt_od_pre_existing.Rows[dt_od_pre_existing.Rows.Count - 1][col_notes] = dt_pre_existing.Rows[i][col_notes];
                                            dt_od_pre_existing.Rows[dt_od_pre_existing.Rows.Count - 1][col_just] = dt_pre_existing.Rows[i][col_just];


                                            double extra1 = 0;

                                            double s1 = Convert.ToDouble(dt_pre_existing.Rows[i][col_start]);
                                            double s2 = Convert.ToDouble(dt_pre_existing.Rows[i][col_end]);

                                            if (dt_eq != null && dt_eq.Rows.Count > 0)
                                            {
                                                for (int k = 0; k < dt_eq.Rows.Count; ++k)
                                                {
                                                    if (dt_eq.Rows[k][col_back] != DBNull.Value && dt_eq.Rows[k][col_ahead] != DBNull.Value)
                                                    {
                                                        double back1 = Convert.ToDouble(dt_eq.Rows[k][col_back]);
                                                        double ahead1 = Convert.ToDouble(dt_eq.Rows[k][col_ahead]);

                                                        if (s1 < back1 && ahead1 < s2)
                                                        {
                                                            extra1 = extra1 + ahead1 - back1;
                                                        }

                                                        if (s1 == back1) s1 = ahead1;

                                                    }
                                                }
                                            }

                                            dt_od_pre_existing.Rows[dt_od_pre_existing.Rows.Count - 1][col_len] = Math.Round(Convert.ToDecimal(s2) - Convert.ToDecimal(s1) - Convert.ToDecimal(extra1), 1);


                                        }
                                    }
                                    #endregion
                                }


                                if (checkBox_output_agen_files.Checked == true)
                                {
                                    #region MATERIAL_LINEAR excel generation for alignment sheet
                                    for (int i = 0; i < dt_compiled.Rows.Count; ++i)
                                    {
                                        if (dt_compiled.Rows[i][col_sta1] != DBNull.Value && dt_compiled.Rows[i][col_sta2] != DBNull.Value)
                                        {
                                            double stabeg = Convert.ToDouble(dt_compiled.Rows[i][col_sta1]);
                                            double staend = Convert.ToDouble(dt_compiled.Rows[i][col_sta2]);

                                            dt_mat_lin.Rows.Add();
                                            dt_mat_lin.Rows[dt_mat_lin.Rows.Count - 1][col_3dbeg] = stabeg;
                                            dt_mat_lin.Rows[dt_mat_lin.Rows.Count - 1][col_3dend] = staend;
                                            dt_mat_lin.Rows[dt_mat_lin.Rows.Count - 1][col_3dlen] = Convert.ToDecimal(staend) - Convert.ToDecimal(stabeg);

                                            dt_mat_lin.Rows[dt_mat_lin.Rows.Count - 1][col_mat] = dt_compiled.Rows[i][col_pipe_type];
                                            dt_mat_lin.Rows[dt_mat_lin.Rows.Count - 1][col_Item_No] = dt_compiled.Rows[i][col_pipe_type];
                                            dt_mat_lin.Rows[dt_mat_lin.Rows.Count - 1][col_id] = dt_compiled.Rows[i][col_elbow_ref_id];

                                            if (dt_compiled.Rows[i][col_descr] != DBNull.Value)
                                            {
                                                string elbow_descr = Convert.ToString(dt_compiled.Rows[i][col_descr]);
                                                for (int k = 0; k < dt_materials.Rows.Count; ++k)
                                                {
                                                    if (dt_materials.Rows[k][col_descr] != DBNull.Value)
                                                    {
                                                        if (Convert.ToString(dt_materials.Rows[k][col_descr]) == elbow_descr)
                                                        {
                                                            elbow_descr = "pipe";
                                                        }
                                                    }
                                                }

                                                string mat1 = "";

                                                if (dt_compiled.Rows[i][col_pipe_type] != DBNull.Value)
                                                {
                                                    mat1 = Convert.ToString(dt_compiled.Rows[i][col_pipe_type]);
                                                }

                                                double pi_elbow = -1;

                                                if (dt_compiled.Rows[i][col_elbow_pi] != DBNull.Value)
                                                {
                                                    pi_elbow = Convert.ToDouble(dt_compiled.Rows[i][col_elbow_pi]);
                                                    dt_mat_lin.Rows[dt_mat_lin.Rows.Count - 1][col_sta] = Functions.Get_chainage_from_double(pi_elbow, "m", 1);
                                                }


                                                if (elbow_descr != "pipe")
                                                {
                                                    dt_mat_lin.Rows[dt_mat_lin.Rows.Count - 1][col_blockdescr] = elbow_descr;
                                                    dt_mat_lin.Rows[dt_mat_lin.Rows.Count - 1][col_block] = "ELBOW";
                                                    if (mat1 == "1")
                                                    {
                                                        dt_mat_lin.Rows[dt_mat_lin.Rows.Count - 1][col_visibility] = "Line Pipe";
                                                    }
                                                    else
                                                    {
                                                        dt_mat_lin.Rows[dt_mat_lin.Rows.Count - 1][col_visibility] = "Heavy Wall";

                                                    }
                                                }
                                                else
                                                {
                                                    dt_mat_lin.Rows[dt_mat_lin.Rows.Count - 1][col_block] = "MAT";

                                                    if (mat1 == "1")
                                                    {
                                                        dt_mat_lin.Rows[dt_mat_lin.Rows.Count - 1][col_visibility] = "Mat1";
                                                    }
                                                    else if (mat1 == "2")
                                                    {
                                                        dt_mat_lin.Rows[dt_mat_lin.Rows.Count - 1][col_visibility] = "Line Pipe";
                                                    }
                                                    else
                                                    {
                                                        dt_mat_lin.Rows[dt_mat_lin.Rows.Count - 1][col_visibility] = "Heavy Wall";

                                                    }
                                                }
                                            }


                                            for (int j = 0; j < dt_cl.Rows.Count - 1; ++j)
                                            {
                                                if (dt_cl.Rows[j]["3DSta"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", "")) == true &&
                                                   dt_cl.Rows[j + 1]["3DSta"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", "")) == true &&
                                                   dt_cl.Rows[j]["X"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["X"])) == true &&
                                                   dt_cl.Rows[j]["Y"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Y"])) == true &&
                                                    dt_cl.Rows[j + 1]["X"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["X"])) == true &&
                                                   dt_cl.Rows[j + 1]["Y"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Y"])) == true)
                                                {
                                                    double sta1 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", ""));
                                                    double sta2 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", ""));
                                                    if (stabeg >= sta1 && stabeg <= sta2)
                                                    {


                                                        double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                        double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);
                                                        double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                        double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);


                                                        double x = x1 + (x2 - x1) * (stabeg - sta1) / (sta2 - sta1);
                                                        double y = y1 + (y2 - y1) * (stabeg - sta1) / (sta2 - sta1);
                                                        dt_mat_lin.Rows[dt_mat_lin.Rows.Count - 1][col_xbeg] = x;
                                                        dt_mat_lin.Rows[dt_mat_lin.Rows.Count - 1][col_ybeg] = y;

                                                        double parameter1 = Poly2D.GetParameterAtPoint(Poly2D.GetClosestPointTo(new Point3d(x, y, Poly2D.Elevation), Vector3d.ZAxis, false));
                                                        if (parameter1 > Poly3D.EndParam) parameter1 = Poly3D.EndParam;

                                                        dt_mat_lin.Rows[dt_mat_lin.Rows.Count - 1][col_mstartcanada] = Poly3D.GetDistanceAtParameter(parameter1);


                                                    }

                                                    if (staend >= sta1 && staend <= sta2)
                                                    {


                                                        double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                        double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);
                                                        double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                        double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);


                                                        double x = x1 + (x2 - x1) * (staend - sta1) / (sta2 - sta1);
                                                        double y = y1 + (y2 - y1) * (staend - sta1) / (sta2 - sta1);
                                                        dt_mat_lin.Rows[dt_mat_lin.Rows.Count - 1][col_xend] = x;
                                                        dt_mat_lin.Rows[dt_mat_lin.Rows.Count - 1][col_yend] = y;


                                                        double parameter1 = Poly2D.GetParameterAtPoint(Poly2D.GetClosestPointTo(new Point3d(x, y, Poly2D.Elevation), Vector3d.ZAxis, false));
                                                        if (parameter1 > Poly3D.EndParam) parameter1 = Poly3D.EndParam;
                                                        dt_mat_lin.Rows[dt_mat_lin.Rows.Count - 1][col_mendcanada] = Poly3D.GetDistanceAtParameter(parameter1);
                                                    }

                                                }

                                            }

                                        }

                                    }
                                    #endregion

                                    #region material points
                                    for (int i = 0; i < dt_xing.Rows.Count; i++)
                                    {
                                        if (dt_xing.Rows[i][col_sta] != DBNull.Value)
                                        {
                                            double sta_xing = Convert.ToDouble(dt_xing.Rows[i][col_sta]);
                                            dt_mat_pt.Rows.Add();
                                            dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_3dsta] = sta_xing;
                                            dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_block] = "MAT_XING";
                                            dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_id] = dt_xing.Rows[i][col_ref_dwg_id];



                                            string desc_xing = "";

                                            if (dt_xing.Rows[i][col_descr1] != DBNull.Value)
                                            {
                                                desc_xing = Convert.ToString(dt_xing.Rows[i][col_descr1]);
                                            }

                                            if (desc_xing != "")
                                            {
                                                if (dt_xing.Rows[i][col_descr2] != DBNull.Value)
                                                {
                                                    desc_xing = desc_xing + "\r\n" + Convert.ToString(dt_xing.Rows[i][col_descr2]);
                                                }
                                            }
                                            else
                                            {
                                                if (dt_xing.Rows[i][col_descr2] != DBNull.Value)
                                                {
                                                    desc_xing = Convert.ToString(dt_xing.Rows[i][col_descr2]);
                                                }
                                            }

                                            string xingid = "";
                                            if (dt_xing.Rows[i][col_xingid] != DBNull.Value)
                                            {
                                                xingid = "(" + Convert.ToString(dt_xing.Rows[i][col_xingid]) + ")";
                                            }

                                            if (desc_xing != "")
                                            {
                                                desc_xing = desc_xing + "\r\n" + xingid;
                                            }
                                            else
                                            {
                                                desc_xing = xingid;
                                            }

                                            if (desc_xing != "")
                                            {
                                                dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_blockdescr] = desc_xing;
                                            }

                                            string cvr1 = "";
                                            if (dt_xing.Rows[i][col_agen_cvr] != DBNull.Value)
                                            {
                                                cvr1 = Convert.ToString(dt_xing.Rows[i][col_agen_cvr]);
                                                if (cvr1.Length > 7)
                                                {
                                                    cvr1 = "{\\W0.78;" + cvr1 + "}";
                                                }

                                            }

                                            if (cvr1 != "")
                                            {
                                                dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_cvr] = cvr1;
                                            }


                                            string type1 = "";
                                            string vis1 = "GEN";

                                            if (dt_xing.Rows[i][col_xingtype] != DBNull.Value)
                                            {
                                                type1 = Convert.ToString(dt_xing.Rows[i][col_xingtype]);

                                                if (type1.ToUpper().Contains("PIPE") == true) vis1 = "1PIPE";
                                                if (type1.ToUpper().Contains("COMMUN") == true) vis1 = "OH";

                                            }

                                            dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_Item_No] = dt_xing.Rows[i][col_xingtype];
                                            dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_visibility] = vis1;


                                            for (int j = 0; j < dt_cl.Rows.Count - 1; ++j)
                                            {
                                                if (dt_cl.Rows[j]["3DSta"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", "")) == true &&
                                                   dt_cl.Rows[j + 1]["3DSta"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", "")) == true &&
                                                   dt_cl.Rows[j]["X"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["X"])) == true &&
                                                   dt_cl.Rows[j]["Y"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Y"])) == true &&
                                                    dt_cl.Rows[j + 1]["X"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["X"])) == true &&
                                                   dt_cl.Rows[j + 1]["Y"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Y"])) == true)
                                                {
                                                    double sta1 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", ""));
                                                    double sta2 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", ""));
                                                    if (sta_xing >= sta1 && sta_xing <= sta2)
                                                    {


                                                        double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                        double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);
                                                        double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                        double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);


                                                        double x = x1 + (x2 - x1) * (sta_xing - sta1) / (sta2 - sta1);
                                                        double y = y1 + (y2 - y1) * (sta_xing - sta1) / (sta2 - sta1);
                                                        dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_x] = x;
                                                        dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_y] = y;

                                                        double parameter1 = Poly2D.GetParameterAtPoint(Poly2D.GetClosestPointTo(new Point3d(x, y, Poly2D.Elevation), Vector3d.ZAxis, false));
                                                        if (parameter1 > Poly3D.EndParam) parameter1 = Poly3D.EndParam;

                                                        dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_mcanada] = Poly3D.GetDistanceAtParameter(parameter1);

                                                        j = dt_cl.Rows.Count;
                                                    }



                                                }

                                            }

                                        }
                                    }


                                    for (int i = 0; i < dt_class.Rows.Count; i++)
                                    {
                                        if (dt_class.Rows[i][col_sta1] != DBNull.Value)
                                        {
                                            double sta_class1 = Convert.ToDouble(dt_class.Rows[i][col_sta1]);
                                            dt_mat_pt.Rows.Add();
                                            dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_3dsta] = sta_class1;
                                            dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_block] = "MAT_XING";
                                            dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_blockdescr] = dt_class.Rows[i][col_descr1];

                                            string type1 = "CLASS";
                                            string vis1 = "NOCL";

                                            dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_Item_No] = type1;
                                            dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_visibility] = vis1;


                                            for (int j = 0; j < dt_cl.Rows.Count - 1; ++j)
                                            {
                                                if (dt_cl.Rows[j]["3DSta"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", "")) == true &&
                                                   dt_cl.Rows[j + 1]["3DSta"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", "")) == true &&
                                                   dt_cl.Rows[j]["X"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["X"])) == true &&
                                                   dt_cl.Rows[j]["Y"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Y"])) == true &&
                                                    dt_cl.Rows[j + 1]["X"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["X"])) == true &&
                                                   dt_cl.Rows[j + 1]["Y"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Y"])) == true)
                                                {
                                                    double sta1 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", ""));
                                                    double sta2 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", ""));
                                                    if (sta_class1 >= sta1 && sta_class1 <= sta2)
                                                    {
                                                        double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                        double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);
                                                        double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                        double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);
                                                        double x = x1 + (x2 - x1) * (sta_class1 - sta1) / (sta2 - sta1);
                                                        double y = y1 + (y2 - y1) * (sta_class1 - sta1) / (sta2 - sta1);
                                                        dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_x] = x;
                                                        dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_y] = y;

                                                        double parameter1 = Poly2D.GetParameterAtPoint(Poly2D.GetClosestPointTo(new Point3d(x, y, Poly2D.Elevation), Vector3d.ZAxis, false));
                                                        if (parameter1 > Poly3D.EndParam) parameter1 = Poly3D.EndParam;

                                                        dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_mcanada] = Poly3D.GetDistanceAtParameter(parameter1);

                                                        j = dt_cl.Rows.Count;
                                                    }
                                                }
                                            }
                                        }
                                    }

                                    for (int i = 0; i < dt_class.Rows.Count; i++)
                                    {
                                        if (dt_class.Rows[i][col_sta2] != DBNull.Value)
                                        {
                                            double sta_class2 = Convert.ToDouble(dt_class.Rows[i][col_sta2]);
                                            dt_mat_pt.Rows.Add();
                                            dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_3dsta] = sta_class2;
                                            dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_block] = "MAT_XING";
                                            dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_blockdescr] = dt_class.Rows[i][col_descr2];
                                            string type1 = "CLASS";
                                            string vis1 = "NOCL";
                                            dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_Item_No] = type1;
                                            dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_visibility] = vis1;

                                            for (int j = 0; j < dt_cl.Rows.Count - 1; ++j)
                                            {
                                                if (dt_cl.Rows[j]["3DSta"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", "")) == true &&
                                                   dt_cl.Rows[j + 1]["3DSta"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", "")) == true &&
                                                   dt_cl.Rows[j]["X"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["X"])) == true &&
                                                   dt_cl.Rows[j]["Y"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Y"])) == true &&
                                                    dt_cl.Rows[j + 1]["X"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["X"])) == true &&
                                                   dt_cl.Rows[j + 1]["Y"] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Y"])) == true)
                                                {
                                                    double sta1 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", ""));
                                                    double sta2 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", ""));
                                                    if (sta_class2 >= sta1 && sta_class2 <= sta2)
                                                    {
                                                        double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                        double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);
                                                        double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                        double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);
                                                        double x = x1 + (x2 - x1) * (sta_class2 - sta1) / (sta2 - sta1);
                                                        double y = y1 + (y2 - y1) * (sta_class2 - sta1) / (sta2 - sta1);
                                                        dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_x] = x;
                                                        dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_y] = y;

                                                        double parameter1 = Poly2D.GetParameterAtPoint(Poly2D.GetClosestPointTo(new Point3d(x, y, Poly2D.Elevation), Vector3d.ZAxis, false));
                                                        if (parameter1 > Poly3D.EndParam) parameter1 = Poly3D.EndParam;

                                                        dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_mcanada] = Poly3D.GetDistanceAtParameter(parameter1);

                                                        j = dt_cl.Rows.Count;
                                                    }
                                                }
                                            }
                                        }
                                    }




                                    if (dt_hydrotest != null && dt_hydrotest.Rows.Count > 0)
                                    {
                                        string type1 = "HYDROTEST";
                                        string vis1 = "NOCL";

                                        for (int i = 0; i < dt_hydrotest.Rows.Count; i++)
                                        {
                                            if (dt_hydrotest.Rows[i][col_sta1] != DBNull.Value)
                                            {
                                                double sta_hydrotest1 = Convert.ToDouble(dt_hydrotest.Rows[i][col_sta1]);
                                                dt_mat_pt.Rows.Add();
                                                dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_3dsta] = sta_hydrotest1;
                                                dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_block] = "MAT_XING";
                                                dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_blockdescr] = dt_hydrotest.Rows[i][col_descr1];
                                                dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_Item_No] = type1;
                                                dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_visibility] = vis1;

                                                for (int j = 0; j < dt_cl.Rows.Count - 1; ++j)
                                                {
                                                    if (dt_cl.Rows[j]["3DSta"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", "")) == true &&
                                                       dt_cl.Rows[j + 1]["3DSta"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", "")) == true &&
                                                       dt_cl.Rows[j]["X"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["X"])) == true &&
                                                       dt_cl.Rows[j]["Y"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Y"])) == true &&
                                                        dt_cl.Rows[j + 1]["X"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["X"])) == true &&
                                                       dt_cl.Rows[j + 1]["Y"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Y"])) == true)
                                                    {
                                                        double sta1 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", ""));
                                                        double sta2 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", ""));
                                                        if (sta_hydrotest1 >= sta1 && sta_hydrotest1 <= sta2)
                                                        {
                                                            double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                            double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);
                                                            double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                            double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);

                                                            double x = x1 + (x2 - x1) * (sta_hydrotest1 - sta1) / (sta2 - sta1);
                                                            double y = y1 + (y2 - y1) * (sta_hydrotest1 - sta1) / (sta2 - sta1);
                                                            dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_x] = x;
                                                            dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_y] = y;

                                                            double parameter1 = Poly2D.GetParameterAtPoint(Poly2D.GetClosestPointTo(new Point3d(x, y, Poly2D.Elevation), Vector3d.ZAxis, false));
                                                            if (parameter1 > Poly3D.EndParam) parameter1 = Poly3D.EndParam;

                                                            dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_mcanada] = Poly3D.GetDistanceAtParameter(parameter1);

                                                            j = dt_cl.Rows.Count;
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                        for (int i = 0; i < dt_hydrotest.Rows.Count; i++)
                                        {
                                            if (dt_hydrotest.Rows[i][col_sta2] != DBNull.Value)
                                            {
                                                double sta_hydrotest2 = Convert.ToDouble(dt_hydrotest.Rows[i][col_sta2]);
                                                dt_mat_pt.Rows.Add();
                                                dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_3dsta] = sta_hydrotest2;
                                                dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_block] = "MAT_XING";
                                                dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_blockdescr] = dt_hydrotest.Rows[i][col_descr2];
                                                dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_Item_No] = type1;
                                                dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_visibility] = vis1;

                                                for (int j = 0; j < dt_cl.Rows.Count - 1; ++j)
                                                {
                                                    if (dt_cl.Rows[j]["3DSta"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", "")) == true &&
                                                       dt_cl.Rows[j + 1]["3DSta"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", "")) == true &&
                                                       dt_cl.Rows[j]["X"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["X"])) == true &&
                                                       dt_cl.Rows[j]["Y"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Y"])) == true &&
                                                        dt_cl.Rows[j + 1]["X"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["X"])) == true &&
                                                       dt_cl.Rows[j + 1]["Y"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Y"])) == true)
                                                    {
                                                        double sta1 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", ""));
                                                        double sta2 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", ""));
                                                        if (sta_hydrotest2 >= sta1 && sta_hydrotest2 <= sta2)
                                                        {
                                                            double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                            double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);
                                                            double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                            double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);
                                                            double x = x1 + (x2 - x1) * (sta_hydrotest2 - sta1) / (sta2 - sta1);
                                                            double y = y1 + (y2 - y1) * (sta_hydrotest2 - sta1) / (sta2 - sta1);
                                                            dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_x] = x;
                                                            dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_y] = y;

                                                            double parameter1 = Poly2D.GetParameterAtPoint(Poly2D.GetClosestPointTo(new Point3d(x, y, Poly2D.Elevation), Vector3d.ZAxis, false));
                                                            if (parameter1 > Poly3D.EndParam) parameter1 = Poly3D.EndParam;

                                                            dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_mcanada] = Poly3D.GetDistanceAtParameter(parameter1);

                                                            j = dt_cl.Rows.Count;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }


                                    if (dt_cpac != null && dt_cpac.Rows.Count > 0)
                                    {
                                        for (int i = 0; i < dt_cpac.Rows.Count; i++)
                                        {
                                            if (dt_cpac.Rows[i][col_sta] != DBNull.Value)
                                            {
                                                double sta_cpac = Convert.ToDouble(dt_cpac.Rows[i][col_sta]);
                                                dt_mat_pt.Rows.Add();
                                                dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_3dsta] = sta_cpac;
                                                dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_block] = "MAT_XING";
                                                dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_blockdescr] = dt_cpac.Rows[i][col_descr];
                                                dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_id] = dt_cpac.Rows[i][col_descr2];

                                                string type1 = "CPAC";
                                                string vis1 = "CP";

                                                dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_Item_No] = type1;
                                                dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_visibility] = vis1;

                                                for (int j = 0; j < dt_cl.Rows.Count - 1; ++j)
                                                {
                                                    if (dt_cl.Rows[j]["3DSta"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", "")) == true &&
                                                       dt_cl.Rows[j + 1]["3DSta"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", "")) == true &&
                                                       dt_cl.Rows[j]["X"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["X"])) == true &&
                                                       dt_cl.Rows[j]["Y"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Y"])) == true &&
                                                        dt_cl.Rows[j + 1]["X"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["X"])) == true &&
                                                       dt_cl.Rows[j + 1]["Y"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Y"])) == true)
                                                    {
                                                        double sta1 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", ""));
                                                        double sta2 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", ""));
                                                        if (sta_cpac >= sta1 && sta_cpac <= sta2)
                                                        {


                                                            double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                            double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);
                                                            double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                            double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);


                                                            double x = x1 + (x2 - x1) * (sta_cpac - sta1) / (sta2 - sta1);
                                                            double y = y1 + (y2 - y1) * (sta_cpac - sta1) / (sta2 - sta1);
                                                            dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_x] = x;
                                                            dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_y] = y;

                                                            double parameter1 = Poly2D.GetParameterAtPoint(Poly2D.GetClosestPointTo(new Point3d(x, y, Poly2D.Elevation), Vector3d.ZAxis, false));
                                                            if (parameter1 > Poly3D.EndParam) parameter1 = Poly3D.EndParam;

                                                            dt_mat_pt.Rows[dt_mat_pt.Rows.Count - 1][col_mcanada] = Poly3D.GetDistanceAtParameter(parameter1);

                                                            j = dt_cl.Rows.Count;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }


                                    #endregion

                                    #region material linear extra
                                    if (dt_buoy != null && dt_buoy.Rows.Count > 0)
                                    {

                                        for (int i = 0; i < dt_buoy.Rows.Count; ++i)
                                        {
                                            if (dt_buoy.Rows[i][col_start] != DBNull.Value && dt_buoy.Rows[i][col_end] != DBNull.Value)
                                            {
                                                double stabeg = Convert.ToDouble(dt_buoy.Rows[i][col_start]);
                                                double staend = Convert.ToDouble(dt_buoy.Rows[i][col_end]);

                                                dt_mat_extra.Rows.Add();
                                                dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_3dbeg] = stabeg;
                                                dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_3dend] = staend;
                                                dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_3dlen] = Convert.ToDecimal(staend) - Convert.ToDecimal(stabeg);
                                                dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_blockdescr] = dt_buoy.Rows[i][col_descr1];
                                                dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_note1] = dt_buoy.Rows[i][col_descr2];
                                                dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_block] = "SA";

                                                if (dt_buoy.Rows[i][col_descr1] != DBNull.Value)
                                                {
                                                    string descr = Convert.ToString(dt_buoy.Rows[i][col_descr1]);

                                                    if (descr.ToUpper().Contains("SA") == true)
                                                    {
                                                        dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_Item_No] = "SA";
                                                        dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_mat] = "SA";
                                                        dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_qty] = dt_buoy.Rows[i][col_count];
                                                        dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_visibility] = "General";
                                                    }
                                                    else if (descr.ToUpper().Contains("CRW") == true)
                                                    {
                                                        dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_Item_No] = "CRW";
                                                        dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_mat] = "CRW";
                                                        dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_qty] = dt_buoy.Rows[i][col_count];
                                                        dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_visibility] = "RiverWeight";

                                                    }
                                                    else if (descr.ToUpper().Contains("CCC") == true)
                                                    {
                                                        dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_Item_No] = "CCC";
                                                        dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_mat] = "CCC";
                                                        dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_qty] = Convert.ToDecimal(staend) - Convert.ToDecimal(stabeg);
                                                        dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_visibility] = "CC";

                                                    }

                                                }


                                                for (int j = 0; j < dt_cl.Rows.Count - 1; ++j)
                                                {
                                                    if (dt_cl.Rows[j]["3DSta"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", "")) == true &&
                                                       dt_cl.Rows[j + 1]["3DSta"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", "")) == true &&
                                                       dt_cl.Rows[j]["X"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["X"])) == true &&
                                                       dt_cl.Rows[j]["Y"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Y"])) == true &&
                                                        dt_cl.Rows[j + 1]["X"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["X"])) == true &&
                                                       dt_cl.Rows[j + 1]["Y"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Y"])) == true)
                                                    {
                                                        double sta1 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", ""));
                                                        double sta2 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", ""));
                                                        double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                        double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);
                                                        double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                        double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);


                                                        if (stabeg >= sta1 && stabeg <= sta2)
                                                        {
                                                            double x = x1 + (x2 - x1) * (stabeg - sta1) / (sta2 - sta1);
                                                            double y = y1 + (y2 - y1) * (stabeg - sta1) / (sta2 - sta1);
                                                            dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_xbeg] = x;
                                                            dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_ybeg] = y;

                                                            double parameter1 = Poly2D.GetParameterAtPoint(Poly2D.GetClosestPointTo(new Point3d(x, y, Poly2D.Elevation), Vector3d.ZAxis, false));
                                                            if (parameter1 > Poly3D.EndParam) parameter1 = Poly3D.EndParam;

                                                            dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_mstartcanada] = Poly3D.GetDistanceAtParameter(parameter1);


                                                        }

                                                        if (staend >= sta1 && staend <= sta2)
                                                        {
                                                            double x = x1 + (x2 - x1) * (staend - sta1) / (sta2 - sta1);
                                                            double y = y1 + (y2 - y1) * (staend - sta1) / (sta2 - sta1);
                                                            dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_xend] = x;
                                                            dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_yend] = y;


                                                            double parameter1 = Poly2D.GetParameterAtPoint(Poly2D.GetClosestPointTo(new Point3d(x, y, Poly2D.Elevation), Vector3d.ZAxis, false));
                                                            if (parameter1 > Poly3D.EndParam) parameter1 = Poly3D.EndParam;
                                                            dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_mendcanada] = Poly3D.GetDistanceAtParameter(parameter1);
                                                        }

                                                    }

                                                }

                                            }

                                        }

                                    }

                                    if (dt_doc != null && dt_doc.Rows.Count > 0)
                                    {

                                        for (int i = 0; i < dt_doc.Rows.Count; ++i)
                                        {
                                            if (dt_doc.Rows[i][col_doc_sta1] != DBNull.Value && dt_doc.Rows[i][col_doc_sta2] != DBNull.Value)
                                            {
                                                double stabeg = Convert.ToDouble(dt_doc.Rows[i][col_doc_sta1]);
                                                double staend = Convert.ToDouble(dt_doc.Rows[i][col_doc_sta2]);

                                                dt_mat_extra.Rows.Add();
                                                dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_3dbeg] = stabeg;
                                                dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_3dend] = staend;
                                                dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_3dlen] = Convert.ToDecimal(staend) - Convert.ToDecimal(stabeg);
                                                dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_blockdescr] = dt_doc.Rows[i][col_descr1];
                                                dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_note1] = dt_doc.Rows[i][col_descr2];
                                                dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_block] = "SA";
                                                dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_Item_No] = "DOC";
                                                dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_mat] = "DOC";
                                                dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_qty] = 1;
                                                dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_visibility] = "DOC";





                                                for (int j = 0; j < dt_cl.Rows.Count - 1; ++j)
                                                {
                                                    if (dt_cl.Rows[j]["3DSta"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", "")) == true &&
                                                       dt_cl.Rows[j + 1]["3DSta"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", "")) == true &&
                                                       dt_cl.Rows[j]["X"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["X"])) == true &&
                                                       dt_cl.Rows[j]["Y"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Y"])) == true &&
                                                        dt_cl.Rows[j + 1]["X"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["X"])) == true &&
                                                       dt_cl.Rows[j + 1]["Y"] != DBNull.Value &&
                                                        Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Y"])) == true)
                                                    {
                                                        double sta1 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", ""));
                                                        double sta2 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", ""));
                                                        double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                                        double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);
                                                        double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                                        double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);


                                                        if (stabeg >= sta1 && stabeg <= sta2)
                                                        {
                                                            double x = x1 + (x2 - x1) * (stabeg - sta1) / (sta2 - sta1);
                                                            double y = y1 + (y2 - y1) * (stabeg - sta1) / (sta2 - sta1);
                                                            dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_xbeg] = x;
                                                            dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_ybeg] = y;

                                                            double parameter1 = Poly2D.GetParameterAtPoint(Poly2D.GetClosestPointTo(new Point3d(x, y, Poly2D.Elevation), Vector3d.ZAxis, false));
                                                            if (parameter1 > Poly3D.EndParam) parameter1 = Poly3D.EndParam;

                                                            dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_mstartcanada] = Poly3D.GetDistanceAtParameter(parameter1);


                                                        }

                                                        if (staend >= sta1 && staend <= sta2)
                                                        {
                                                            double x = x1 + (x2 - x1) * (staend - sta1) / (sta2 - sta1);
                                                            double y = y1 + (y2 - y1) * (staend - sta1) / (sta2 - sta1);
                                                            dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_xend] = x;
                                                            dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_yend] = y;


                                                            double parameter1 = Poly2D.GetParameterAtPoint(Poly2D.GetClosestPointTo(new Point3d(x, y, Poly2D.Elevation), Vector3d.ZAxis, false));
                                                            if (parameter1 > Poly3D.EndParam) parameter1 = Poly3D.EndParam;
                                                            dt_mat_extra.Rows[dt_mat_extra.Rows.Count - 1][col_mendcanada] = Poly3D.GetDistanceAtParameter(parameter1);
                                                        }

                                                    }

                                                }

                                            }

                                        }

                                    }

                                    #endregion
                                }

                                Poly3D.Erase();
                                Trans1.Commit();
                            }

                            if (checkBox_draw_canadian_mat.Checked == true)
                            {
                                attach_od_to_pipes(dt_od_pipe);
                                attach_od_to_buoyancy(dt_od_buoy);
                                attach_od_to_buoyancy_start(dt_od_buoy);
                                attach_od_to_buoyancy_start(dt_od_long_strap);

                                attach_od_to_buoyancy_end(dt_od_buoy);
                                attach_od_to_cpac(dt_od_cpac);
                                attach_od_to_es(dt_od_es);
                                attach_od_to_hydrotest(dt_od_hydrotest);
                                attach_od_to_hydrotestPT(dt_od_hydrotest_pt);
                                attach_od_to_transition(dt_od_trans);

                                attach_od_to_objects(dt_od_xing, xing_od);
                                attach_od_to_objects(dt_od_elbow, elbows_od);
                                attach_od_to_objects(dt_od_fab, fab_od);
                                attach_od_to_objects(dt_od_class, class_od);
                                attach_od_to_objects(dt_od_muskeg, muskeg_od);
                                attach_od_to_objects(dt_od_doc, doc_od);
                                attach_od_to_objects(dt_od_pre_existing, preexisting_od);

                                attach_od_to_geotech(dt_od_geotech, geotech_od);
                                attach_od_to_geotech_start(dt_od_geotech);
                                attach_od_to_geotech_end(dt_od_geotech);

                            }
                        }

                        Functions.Transfer_datatable_to_new_excel_spreadsheet(dt_mat_lin, "MatLinear");
                        Functions.Transfer_datatable_to_new_excel_spreadsheet(dt_mat_pt, "MatPts");
                        Functions.Transfer_datatable_to_new_excel_spreadsheet(dt_mat_extra, "MatExtra");

                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                }
            }



            set_enable_true();
        }




        private void insert_elbows(ref System.Data.DataTable dt_ref, System.Data.DataTable dt_in)
        {


            string col_sta1 = "Sta1";
            string col_sta2 = "Sta2";

            string col_defl = "Deflection";


            if (dt_ref != null && dt_ref.Rows.Count > 0 && dt_in != null && dt_in.Rows.Count > 0)
            {
                for (int i = 0; i < dt_ref.Rows.Count; ++i)
                {
                    double sta1 = Convert.ToDouble(dt_ref.Rows[i][col_sta1]);
                    double sta2 = Convert.ToDouble(dt_ref.Rows[i][col_sta2]);

                    for (int j = dt_in.Rows.Count - 1; j >= 0; --j)
                    {
                        if (dt_in.Rows[j][col_sta1] != DBNull.Value && dt_in.Rows[j][col_sta2] != DBNull.Value)
                        {
                            double sta_elbow1 = Convert.ToDouble(dt_in.Rows[j][col_sta1]);
                            double sta_elbow2 = Convert.ToDouble(dt_in.Rows[j][col_sta2]);
                            string just_elbow1 = "";

                            if (dt_in.Rows[j][col_just] != DBNull.Value)
                            {
                                just_elbow1 = Convert.ToString(dt_in.Rows[j][col_just]);
                            }

                            if (just_elbow1.ToLower().Contains("**elbow") == false)
                            {
                                just_elbow1 = just_elbow1 + "**elbow";
                            }

                            if (sta1 == sta_elbow1 && sta2 == sta_elbow2)
                            {
                                dt_ref.Rows[i][col_wt] = dt_in.Rows[j][col_wt];
                                dt_ref.Rows[i][col_pipe_type] = dt_in.Rows[j][col_pipe_type];
                                dt_ref.Rows[i][col_elbow_pi] = dt_in.Rows[j][col_elbow_pi];
                                dt_ref.Rows[i][col_defl] = dt_in.Rows[j][col_defl];
                                dt_ref.Rows[i][col_descr] = dt_in.Rows[j][col_descr];
                                dt_ref.Rows[i][col_elbow_ref_id] = dt_in.Rows[j][col_elbow_ref_dwg];
                                dt_ref.Rows[i][col_just] = just_elbow1;
                                dt_in.Rows[j].Delete();
                            }

                            else
                            {
                                #region sta_elbow1 > sta1 && sta_elbow2 < sta2
                                if (sta_elbow1 > sta1 && sta_elbow2 < sta2)
                                {
                                    System.Data.DataRow row1 = dt_ref.NewRow();
                                    row1.ItemArray = dt_ref.Rows[i].ItemArray;
                                    row1[col_sta1] = sta_elbow2;

                                    System.Data.DataRow row2 = dt_ref.NewRow();
                                    row2[col_sta1] = sta_elbow1;
                                    row2[col_sta2] = sta_elbow2;
                                    row2[col_wt] = dt_in.Rows[j][col_wt];
                                    row2[col_pipe_type] = dt_in.Rows[j][col_pipe_type];
                                    row2[col_elbow_pi] = dt_in.Rows[j][col_elbow_pi];
                                    row2[col_defl] = dt_in.Rows[j][col_defl];
                                    row2[col_descr] = dt_in.Rows[j][col_descr];
                                    row2[col_elbow_ref_id] = dt_in.Rows[j][col_elbow_ref_dwg];

                                    row2[col_just] = just_elbow1;

                                    dt_ref.Rows.Add(row1);
                                    dt_ref.Rows.Add(row2);

                                    dt_ref.Rows[i][col_sta2] = sta_elbow1;
                                    dt_in.Rows[j].Delete();
                                }
                                #endregion
                                #region sta_elbow1 == sta1 && sta_elbow2 < sta2
                                else if (sta_elbow1 == sta1 && sta_elbow2 < sta2)
                                {
                                    System.Data.DataRow row2 = dt_ref.NewRow();
                                    row2[col_sta1] = sta_elbow1;
                                    row2[col_sta2] = sta_elbow2;
                                    row2[col_wt] = dt_in.Rows[j][col_wt];
                                    row2[col_pipe_type] = dt_in.Rows[j][col_pipe_type];
                                    row2[col_elbow_pi] = dt_in.Rows[j][col_elbow_pi];
                                    row2[col_defl] = dt_in.Rows[j][col_defl];

                                    row2[col_descr] = dt_in.Rows[j][col_descr];
                                    row2[col_elbow_ref_id] = dt_in.Rows[j][col_elbow_ref_dwg];
                                    row2[col_just] = just_elbow1;

                                    dt_ref.Rows.Add(row2);

                                    dt_ref.Rows[i][col_sta1] = sta_elbow2;
                                    dt_in.Rows[j].Delete();

                                }
                                #endregion
                                #region sta_elbow1 > sta1 && sta_elbow2 == sta2
                                else if (sta_elbow1 > sta1 && sta_elbow2 == sta2)
                                {
                                    System.Data.DataRow row2 = dt_ref.NewRow();
                                    row2[col_sta1] = sta_elbow1;
                                    row2[col_sta2] = sta_elbow2;
                                    row2[col_wt] = dt_in.Rows[j][col_wt];
                                    row2[col_pipe_type] = dt_in.Rows[j][col_pipe_type];
                                    row2[col_elbow_pi] = dt_in.Rows[j][col_elbow_pi];
                                    row2[col_defl] = dt_in.Rows[j][col_defl];

                                    row2[col_descr] = dt_in.Rows[j][col_descr];
                                    row2[col_elbow_ref_id] = dt_in.Rows[j][col_elbow_ref_dwg];
                                    row2[col_just] = just_elbow1;
                                    dt_ref.Rows.Add(row2);
                                    dt_ref.Rows[i][col_sta2] = sta_elbow1;
                                    dt_in.Rows[j].Delete();
                                }
                                #endregion

                                #region sta_elbow1 > sta1 && sta_elbow2 > sta2 && sta_elbow1 < sta2
                                else if (sta_elbow1 > sta1 && sta_elbow2 > sta2 && sta_elbow1 < sta2)
                                {

                                    System.Data.DataRow row2 = dt_ref.NewRow();
                                    row2[col_sta1] = sta_elbow1;
                                    row2[col_sta2] = sta_elbow2;
                                    row2[col_wt] = dt_in.Rows[j][col_wt];
                                    row2[col_pipe_type] = dt_in.Rows[j][col_pipe_type];
                                    row2[col_elbow_pi] = dt_in.Rows[j][col_elbow_pi];
                                    row2[col_defl] = dt_in.Rows[j][col_defl];

                                    row2[col_descr] = dt_in.Rows[j][col_descr];
                                    row2[col_elbow_ref_id] = dt_in.Rows[j][col_elbow_ref_dwg];
                                    row2[col_just] = just_elbow1;

                                    dt_ref.Rows.Add(row2);
                                    dt_ref.Rows[i][col_sta2] = sta_elbow1;
                                    if (i < dt_ref.Rows.Count - 1)
                                    {
                                        dt_ref.Rows[i + 1][col_sta1] = sta_elbow2;
                                    }

                                    dt_in.Rows[j].Delete();
                                }
                                #endregion

                            }
                        }
                    }
                }

                dt_ref = Functions.Sort_data_table(dt_ref, col_sta1);
            }

        }

        private void insert_dt_into_dt_compiledV2(ref System.Data.DataTable dt_comp, System.Data.DataTable dt_in, System.Data.DataTable dt_mat)
        {


            Polyline poly1 = new Polyline();
            int k = 0;
            for (int j = 0; j < dt_comp.Rows.Count; ++j)
            {
                if (dt_comp.Rows[j][col_sta1] != DBNull.Value)
                {
                    double sta1 = Convert.ToDouble(dt_comp.Rows[j][col_sta1]);
                    poly1.AddVertexAt(k, new Point2d(sta1, 0), 0, 0, 0);
                    ++k;

                }
            }
            double sta_end = Convert.ToDouble(dt_comp.Rows[dt_comp.Rows.Count - 1][col_sta2]);
            poly1.AddVertexAt(k, new Point2d(sta_end, 0), 0, 0, 0);

            if (dt_in != null && dt_in.Rows.Count > 0)
            {
                for (int i = 0; i < dt_in.Rows.Count; ++i)
                {
                    if (dt_in.Rows[i][col_sta1] != DBNull.Value && dt_in.Rows[i][col_sta2] != DBNull.Value)
                    {
                        double sta1 = Convert.ToDouble(dt_in.Rows[i][col_sta1]);
                        double sta2 = Convert.ToDouble(dt_in.Rows[i][col_sta2]);
                        double wt1 = Convert.ToDouble(dt_in.Rows[i][col_wt]);
                        string mat1 = Convert.ToString(dt_in.Rows[i][col_pipe_type]);
                        string coating1 = get_coating(dt_mat, mat1).ToUpper();

                        if (poly1.Length + poly1.StartPoint.X < sta1) sta1 = poly1.Length + poly1.StartPoint.X;
                        if (poly1.Length + poly1.StartPoint.X < sta2) sta2 = poly1.Length + poly1.StartPoint.X;

                        double par1 = poly1.GetParameterAtPoint(poly1.GetClosestPointTo(new Point3d(sta1, 0, 0), Vector3d.ZAxis, false));
                        int p1 = Convert.ToInt32(Math.Ceiling(par1));
                        double d1 = Math.Abs(poly1.GetDistanceAtParameter(p1));

                        double par2 = poly1.GetParameterAtPoint(poly1.GetClosestPointTo(new Point3d(sta2, 0, 0), Vector3d.ZAxis, false));
                        int p2 = Convert.ToInt32(Math.Ceiling(par2));
                        double d2 = Math.Abs(poly1.GetDistanceAtParameter(p2));

                        int idx1 = p1;
                        int idx2 = p2;
                        double d_sta1 = poly1.GetDistanceAtParameter(par1);
                        double d_sta2 = poly1.GetDistanceAtParameter(par2);

                        // if (p1 > poly1.EndParam) p1 = poly1.EndParam;
                        //if (p2 > poly1.EndParam) p2 = poly1.EndParam;

                        if (Math.Round(d1, 1) == Math.Round(d_sta1, 1))
                        {

                        }
                        else
                        {
                            System.Data.DataRow row_extra1 = dt_comp.NewRow();


                            row_extra1.ItemArray = dt_comp.Rows[idx1 - 1].ItemArray;

                            dt_comp.Rows[idx1 - 1][col_sta2] = sta1;
                            row_extra1[col_sta1] = sta1;

                            dt_comp.Rows.InsertAt(row_extra1, idx1);
                            poly1.AddVertexAt(idx1, new Point2d(sta1, 0), 0, 0, 0);

                            idx2 = p2 + 1;
                        }






                        if (Math.Round(d2, 1) == Math.Round(d_sta2, 1))
                        {

                        }
                        else
                        {
                            System.Data.DataRow row_extra2 = dt_comp.NewRow();
                            row_extra2.ItemArray = dt_comp.Rows[idx2 - 1].ItemArray;
                            dt_comp.Rows[idx2 - 1][col_sta2] = sta2;
                            row_extra2[col_sta1] = sta2;
                            dt_comp.Rows.InsertAt(row_extra2, idx2);
                            poly1.AddVertexAt(idx2, new Point2d(sta2, 0), 0, 0, 0);
                        }

                        if (idx1 == idx2) idx2 = idx2 + 1;
                        for (int j = idx1; j < idx2; ++j)
                        {
                            double wt2 = Convert.ToDouble(dt_comp.Rows[j][col_wt]);
                            string mat2 = Convert.ToString(dt_comp.Rows[j][col_pipe_type]);
                            string coating2 = get_coating(dt_mat, mat2).ToUpper();

                            if (wt2 < wt1 || (wt1 == wt2 && coating2 == "FBE" && coating1 != "FBE"))
                            {
                                dt_comp.Rows[j][col_pipe_type] = mat1;
                                dt_comp.Rows[j][col_wt] = wt1;
                                dt_comp.Rows[j][col_just] = dt_in.Rows[i][col_just];
                                if (dt_in.Columns.Contains(col_notes) == true) dt_comp.Rows[j][col_notes] = dt_in.Rows[i][col_notes];
                            }


                        }


                    }

                }


            }


            for (int i = dt_comp.Rows.Count - 1; i > 0; --i)
            {

                string mat2 = Convert.ToString(dt_comp.Rows[i][col_pipe_type]);
                string just2 = Convert.ToString(dt_comp.Rows[i][col_just]);
                string note2 = Convert.ToString(dt_comp.Rows[i][col_notes]);


                double sta21 = Convert.ToDouble(dt_comp.Rows[i][col_sta1]);
                double sta22 = Convert.ToDouble(dt_comp.Rows[i][col_sta2]);

                string mat1 = Convert.ToString(dt_comp.Rows[i - 1][col_pipe_type]);
                string just1 = Convert.ToString(dt_comp.Rows[i - 1][col_just]);
                string note1 = Convert.ToString(dt_comp.Rows[i - 1][col_notes]);




                if (sta21 < sta22)
                {
                    if (mat2 == mat1)
                    {
                        dt_comp.Rows[i - 1][col_sta2] = sta22;
                        if (just1 != just2)
                        {
                            dt_comp.Rows[i - 1][col_just] = just1 + ", " + just2;
                        }

                        if (note1 != note2)
                        {
                            dt_comp.Rows[i - 1][col_notes] = note1 + ", " + note2;
                        }

                        dt_comp.Rows[i].Delete();
                    }
                }
                else
                {
                    dt_comp.Rows[i].Delete();
                }


            }

        }


        private void insert_fab(ref System.Data.DataTable dt1, System.Data.DataTable dt_fab, System.Data.DataTable dt_pre_existing)
        {
            string col_pi = "PI";
            string col_angle = "Angle";
            string col_defl = "Deflection";


            if (dt1 != null && dt1.Rows.Count > 0 && dt_fab != null && dt_fab.Rows.Count > 0)
            {

                for (int i = 0; i < dt_fab.Rows.Count; ++i)
                {
                    if (dt_fab.Rows[i][col_sta1] != DBNull.Value && dt_fab.Rows[i][col_sta2] != DBNull.Value)
                    {
                        double sta_fab1 = Convert.ToDouble(dt_fab.Rows[i][col_sta1]);
                        double sta_fab2 = Convert.ToDouble(dt_fab.Rows[i][col_sta2]);
                        string just_fab = "Assembly";
                        if (dt_fab.Rows[i][col_just] != DBNull.Value)
                        {
                            just_fab = Convert.ToString(dt_fab.Rows[i][col_just]);
                        }

                        for (int j = dt1.Rows.Count - 1; j >= 0; --j)
                        {
                            double sta1 = Convert.ToDouble(dt1.Rows[j][col_sta1]);
                            double sta2 = Convert.ToDouble(dt1.Rows[j][col_sta2]);

                            if (sta1 >= sta_fab1 && sta2 <= sta_fab2)
                            {
                                dt1.Rows[j].Delete();
                            }

                            else if (sta2 > sta_fab2 && sta1 >= sta_fab1 && sta1 < sta_fab2)
                            {
                                dt1.Rows[j][col_sta1] = sta_fab2;
                            }

                            else if (sta2 <= sta_fab2 && sta1 < sta_fab1 && sta2 > sta_fab1)
                            {
                                dt1.Rows[j][col_sta2] = sta_fab1;
                            }
                            else if (sta2 > sta_fab2 && sta1 < sta_fab1 && sta2 > sta_fab1)
                            {

                                System.Data.DataRow row1 = dt1.NewRow();
                                row1.ItemArray = dt1.Rows[j].ItemArray;
                                row1[col_sta1] = sta_fab2;
                                dt1.Rows[j][col_sta2] = sta_fab1;
                                dt1.Rows.Add(row1);
                            }

                        }

                        dt1.Rows.Add();
                        dt1.Rows[dt1.Rows.Count - 1][col_sta1] = sta_fab1;
                        dt1.Rows[dt1.Rows.Count - 1][col_sta2] = sta_fab2;
                        dt1.Rows[dt1.Rows.Count - 1][col_wt] = DBNull.Value;
                        dt1.Rows[dt1.Rows.Count - 1][col_pipe_type] = DBNull.Value;
                        dt1.Rows[dt1.Rows.Count - 1][col_pi] = DBNull.Value;
                        dt1.Rows[dt1.Rows.Count - 1][col_defl] = DBNull.Value;

                        dt1.Rows[dt1.Rows.Count - 1][col_descr] = DBNull.Value;
                        dt1.Rows[dt1.Rows.Count - 1][col_just] = just_fab;
                    }
                }


                dt1 = Functions.Sort_data_table(dt1, col_sta1);

            }

            if (dt1 != null && dt1.Rows.Count > 0 && dt_pre_existing != null && dt_pre_existing.Rows.Count > 0)
            {

                for (int i = 0; i < dt_pre_existing.Rows.Count; ++i)
                {
                    if (dt_pre_existing.Rows[i][col_start] != DBNull.Value && dt_pre_existing.Rows[i][col_end] != DBNull.Value)
                    {
                        double sta_fab1 = Convert.ToDouble(dt_pre_existing.Rows[i][col_start]);
                        double sta_fab2 = Convert.ToDouble(dt_pre_existing.Rows[i][col_end]);
                        string just_fab = "Existing Pipe";
                        if (dt_pre_existing.Rows[i][col_just] != DBNull.Value)
                        {
                            just_fab = Convert.ToString(dt_pre_existing.Rows[i][col_just]);
                        }

                        for (int j = dt1.Rows.Count - 1; j >= 0; --j)
                        {
                            double sta1 = Convert.ToDouble(dt1.Rows[j][col_sta1]);
                            double sta2 = Convert.ToDouble(dt1.Rows[j][col_sta2]);

                            if (sta1 >= sta_fab1 && sta2 <= sta_fab2)
                            {
                                dt1.Rows[j].Delete();
                            }

                            else if (sta2 > sta_fab2 && sta1 >= sta_fab1 && sta1 < sta_fab2)
                            {
                                dt1.Rows[j][col_sta1] = sta_fab2;
                            }

                            else if (sta2 <= sta_fab2 && sta1 < sta_fab1 && sta2 > sta_fab1)
                            {
                                dt1.Rows[j][col_sta2] = sta_fab1;
                            }
                            else if (sta2 > sta_fab2 && sta1 < sta_fab1 && sta2 > sta_fab1)
                            {

                                System.Data.DataRow row1 = dt1.NewRow();
                                row1.ItemArray = dt1.Rows[j].ItemArray;
                                row1[col_sta1] = sta_fab2;
                                dt1.Rows[j][col_sta2] = sta_fab1;
                                dt1.Rows.Add(row1);
                            }

                        }

                        dt1.Rows.Add();
                        dt1.Rows[dt1.Rows.Count - 1][col_sta1] = sta_fab1;
                        dt1.Rows[dt1.Rows.Count - 1][col_sta2] = sta_fab2;
                        dt1.Rows[dt1.Rows.Count - 1][col_wt] = DBNull.Value;
                        dt1.Rows[dt1.Rows.Count - 1][col_pipe_type] = DBNull.Value;
                        dt1.Rows[dt1.Rows.Count - 1][col_pi] = DBNull.Value;
                        dt1.Rows[dt1.Rows.Count - 1][col_defl] = DBNull.Value;

                        dt1.Rows[dt1.Rows.Count - 1][col_descr] = DBNull.Value;
                        dt1.Rows[dt1.Rows.Count - 1][col_just] = just_fab;
                    }
                }


                dt1 = Functions.Sort_data_table(dt1, col_sta1);

            }

        }

        private void insert_adjacent(ref System.Data.DataTable dt1, System.Data.DataTable dt_el_copy1, System.Data.DataTable dt_mat)
        {



            if (dt1 != null && dt1.Rows.Count > 0 && dt_el_copy1 != null && dt_el_copy1.Rows.Count > 0)
            {
                System.Data.DataTable dt_el_copy2 = new System.Data.DataTable();
                dt_el_copy2 = dt_el_copy1.Copy();
                // here i did for left first then for right
                // i did not want to have 4 pipesta to check it against sta1 and sta2

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    double sta1 = Convert.ToDouble(dt1.Rows[i][col_sta1]);
                    double sta2 = Convert.ToDouble(dt1.Rows[i][col_sta2]);
                    double wt1 = Convert.ToDouble(dt1.Rows[i][col_wt]);
                    string mat1 = Convert.ToString(dt1.Rows[i][col_pipe_type]);
                    string coat1 = get_coating(dt_mat, mat1);

                    bool delete_row_j = false;


                    for (int j = dt_el_copy1.Rows.Count - 1; j >= 0; --j)
                    {



                        string just1 = "Adjacent Pipe";
                        double wt2 = Convert.ToDouble(dt_el_copy1.Rows[j][col_wt]);
                        string mat2 = Convert.ToString(dt_el_copy1.Rows[j][col_pipe_type]);
                        string coat2 = get_coating(dt_mat, mat2);





                        if (dt_el_copy1.Rows[j][col_start_elbow_adjacent] != DBNull.Value && dt_el_copy1.Rows[j][col_sta1] != DBNull.Value)
                        {
                            double sta_pipe1 = Convert.ToDouble(dt_el_copy1.Rows[j][col_start_elbow_adjacent]);
                            double sta_pipe2 = Convert.ToDouble(dt_el_copy1.Rows[j][col_sta1]);
                            if (wt2 > wt1 || (wt2 == wt1 && coat1 == "FBE" && coat1 != coat2))
                            {
                                if (sta1 == sta_pipe1 && sta2 == sta_pipe2)
                                {
                                    dt1.Rows[i][col_wt] = dt_el_copy1.Rows[j][col_wt];
                                    dt1.Rows[i][col_pipe_type] = dt_el_copy1.Rows[j][col_pipe_type];
                                    dt1.Rows[i][col_just] = just1;
                                    delete_row_j = true;

                                }

                                else
                                {
                                    #region sta_pipe1 > sta1 && sta_pipe2 < sta2
                                    if (sta_pipe1 > sta1 && sta_pipe2 < sta2)
                                    {
                                        System.Data.DataRow row1 = dt1.NewRow();
                                        row1.ItemArray = dt1.Rows[i].ItemArray;
                                        row1[col_sta1] = sta_pipe2;

                                        System.Data.DataRow row2 = dt1.NewRow();
                                        row2[col_sta1] = sta_pipe1;
                                        row2[col_sta2] = sta_pipe2;
                                        row2[col_wt] = dt_el_copy1.Rows[j][col_wt];
                                        row2[col_pipe_type] = dt_el_copy1.Rows[j][col_pipe_type];
                                        row2[col_just] = just1;

                                        dt1.Rows.Add(row1);
                                        dt1.Rows.Add(row2);

                                        dt1.Rows[i][col_sta2] = sta_pipe1;
                                        delete_row_j = true;

                                    }
                                    #endregion
                                    #region sta_pipe1 == sta1 && sta_pipe2 < sta2
                                    else if (sta_pipe1 == sta1 && sta_pipe2 < sta2)
                                    {
                                        System.Data.DataRow row2 = dt1.NewRow();
                                        row2[col_sta1] = sta_pipe1;
                                        row2[col_sta2] = sta_pipe2;
                                        row2[col_wt] = dt_el_copy1.Rows[j][col_wt];
                                        row2[col_pipe_type] = dt_el_copy1.Rows[j][col_pipe_type];
                                        row2[col_just] = just1;

                                        dt1.Rows.Add(row2);

                                        dt1.Rows[i][col_sta1] = sta_pipe2;
                                        delete_row_j = true;


                                    }
                                    #endregion
                                    #region sta_pipe1 > sta1 && sta_pipe2 == sta2
                                    else if (sta_pipe1 > sta1 && sta_pipe2 == sta2)
                                    {
                                        System.Data.DataRow row2 = dt1.NewRow();
                                        row2[col_sta1] = sta_pipe1;
                                        row2[col_sta2] = sta_pipe2;
                                        row2[col_wt] = dt_el_copy1.Rows[j][col_wt];
                                        row2[col_pipe_type] = dt_el_copy1.Rows[j][col_pipe_type];
                                        row2[col_just] = just1;
                                        dt1.Rows.Add(row2);
                                        dt1.Rows[i][col_sta2] = sta_pipe1;
                                        delete_row_j = true;

                                    }
                                    #endregion

                                    #region sta_pipe1 > sta1 && sta_pipe2 > sta2 && sta_pipe1 < sta2
                                    else if (sta_pipe1 > sta1 && sta_pipe2 > sta2 && sta_pipe1 < sta2)
                                    {

                                        System.Data.DataRow row2 = dt1.NewRow();
                                        row2[col_sta1] = sta_pipe1;
                                        row2[col_sta2] = sta_pipe2;
                                        row2[col_wt] = dt_el_copy1.Rows[j][col_wt];
                                        row2[col_pipe_type] = dt_el_copy1.Rows[j][col_pipe_type];
                                        row2[col_just] = just1;

                                        dt1.Rows.Add(row2);
                                        dt1.Rows[i][col_sta2] = sta_pipe1;
                                        if (i < dt1.Rows.Count - 1)
                                        {
                                            dt1.Rows[i + 1][col_sta1] = sta_pipe2;
                                        }

                                        delete_row_j = true;
                                    }
                                    #endregion

                                }
                            }
                        }






                        if (delete_row_j == true)
                        {
                            dt_el_copy1.Rows[j].Delete();

                        }

                    }

                }

                dt1 = Functions.Sort_data_table(dt1, col_sta1);


                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    double sta1 = Convert.ToDouble(dt1.Rows[i][col_sta1]);
                    double sta2 = Convert.ToDouble(dt1.Rows[i][col_sta2]);
                    double wt1 = Convert.ToDouble(dt1.Rows[i][col_wt]);
                    string mat1 = Convert.ToString(dt1.Rows[i][col_pipe_type]);
                    string coat1 = get_coating(dt_mat, mat1);


                    bool delete_row_j = false;
                    for (int j = dt_el_copy2.Rows.Count - 1; j >= 0; --j)
                    {

                        string just1 = "Adjacent Pipe";
                        double wt2 = Convert.ToDouble(dt_el_copy2.Rows[j][col_wt]);
                        string mat2 = Convert.ToString(dt_el_copy2.Rows[j][col_pipe_type]);
                        string coat2 = get_coating(dt_mat, mat2);

                        if (dt_el_copy2.Rows[j][col_sta2] != DBNull.Value && dt_el_copy2.Rows[j][col_end_elbow_adjacent] != DBNull.Value)
                        {
                            double sta_pipe1 = Convert.ToDouble(dt_el_copy2.Rows[j][col_sta2]);
                            double sta_pipe2 = Convert.ToDouble(dt_el_copy2.Rows[j][col_end_elbow_adjacent]);
                            if (wt2 > wt1 || (wt2 == wt1 && coat1 == "FBE" && coat1 != coat2))
                            {
                                if (sta1 == sta_pipe1 && sta2 == sta_pipe2)
                                {
                                    dt1.Rows[i][col_wt] = dt_el_copy2.Rows[j][col_wt];
                                    dt1.Rows[i][col_pipe_type] = dt_el_copy2.Rows[j][col_pipe_type];
                                    dt1.Rows[i][col_just] = just1;
                                    delete_row_j = true;

                                }

                                else
                                {
                                    #region sta_pipe1 > sta1 && sta_pipe2 < sta2
                                    if (sta_pipe1 > sta1 && sta_pipe2 < sta2)
                                    {
                                        System.Data.DataRow row1 = dt1.NewRow();
                                        row1.ItemArray = dt1.Rows[i].ItemArray;
                                        row1[col_sta1] = sta_pipe2;

                                        System.Data.DataRow row2 = dt1.NewRow();
                                        row2[col_sta1] = sta_pipe1;
                                        row2[col_sta2] = sta_pipe2;
                                        row2[col_wt] = dt_el_copy2.Rows[j][col_wt];
                                        row2[col_pipe_type] = dt_el_copy2.Rows[j][col_pipe_type];
                                        row2[col_just] = just1;

                                        dt1.Rows.Add(row1);
                                        dt1.Rows.Add(row2);

                                        dt1.Rows[i][col_sta2] = sta_pipe1;
                                        delete_row_j = true;

                                    }
                                    #endregion
                                    #region sta_pipe1 == sta1 && sta_pipe2 < sta2
                                    else if (sta_pipe1 == sta1 && sta_pipe2 < sta2)
                                    {
                                        System.Data.DataRow row2 = dt1.NewRow();
                                        row2[col_sta1] = sta_pipe1;
                                        row2[col_sta2] = sta_pipe2;
                                        row2[col_wt] = dt_el_copy2.Rows[j][col_wt];
                                        row2[col_pipe_type] = dt_el_copy2.Rows[j][col_pipe_type];
                                        row2[col_just] = just1;

                                        dt1.Rows.Add(row2);

                                        dt1.Rows[i][col_sta1] = sta_pipe2;
                                        delete_row_j = true;


                                    }
                                    #endregion
                                    #region sta_pipe1 > sta1 && sta_pipe2 == sta2
                                    else if (sta_pipe1 > sta1 && sta_pipe2 == sta2)
                                    {
                                        System.Data.DataRow row2 = dt1.NewRow();
                                        row2[col_sta1] = sta_pipe1;
                                        row2[col_sta2] = sta_pipe2;
                                        row2[col_wt] = dt_el_copy2.Rows[j][col_wt];
                                        row2[col_pipe_type] = dt_el_copy2.Rows[j][col_pipe_type];
                                        row2[col_just] = just1;
                                        dt1.Rows.Add(row2);
                                        dt1.Rows[i][col_sta2] = sta_pipe1;
                                        delete_row_j = true;

                                    }
                                    #endregion

                                    #region sta_pipe1 > sta1 && sta_pipe2 > sta2 && sta_pipe1 < sta2
                                    else if (sta_pipe1 > sta1 && sta_pipe2 > sta2 && sta_pipe1 < sta2)
                                    {

                                        System.Data.DataRow row2 = dt1.NewRow();
                                        row2[col_sta1] = sta_pipe1;
                                        row2[col_sta2] = sta_pipe2;
                                        row2[col_wt] = dt_el_copy2.Rows[j][col_wt];
                                        row2[col_pipe_type] = dt_el_copy2.Rows[j][col_pipe_type];
                                        row2[col_just] = just1;

                                        dt1.Rows.Add(row2);
                                        dt1.Rows[i][col_sta2] = sta_pipe1;
                                        if (i < dt1.Rows.Count - 1)
                                        {
                                            dt1.Rows[i + 1][col_sta1] = sta_pipe2;
                                        }

                                        delete_row_j = true;
                                    }
                                    #endregion

                                }
                            }
                        }






                        if (delete_row_j == true)
                        {
                            dt_el_copy2.Rows[j].Delete();

                        }

                    }

                }


                dt1 = Functions.Sort_data_table(dt1, col_sta1);
                //Functions.Transfer_datatable_to_new_excel_spreadsheet(dt1, "0.0");
            }

        }

        private System.Data.DataTable create_dt_transition(System.Data.DataTable dt_comp, System.Data.DataTable dt_t)
        {
            System.Data.DataTable dt1 = new System.Data.DataTable();

            dt1.Columns.Add(col_sta, typeof(double));
            dt1.Columns.Add("x1", typeof(double));
            dt1.Columns.Add("y1", typeof(double));
            dt1.Columns.Add("layer", typeof(string));
            dt1.Columns.Add("ci", typeof(short));
            if (dt_comp.Rows.Count > 0 && dt_t != null && dt_t.Rows.Count > 0)
            {
                for (int i = 0; i < dt_comp.Rows.Count - 1; ++i)
                {
                    string mat1 = Convert.ToString(dt_comp.Rows[i][col_pipe_type]);
                    string mat2 = Convert.ToString(dt_comp.Rows[i + 1][col_pipe_type]);
                    int index1 = -1;
                    int index2 = -1;
                    if (dt_t.Columns.Contains(mat1) == true && dt_t.Columns.Contains(mat2) == true)
                    {
                        index1 = dt_t.Columns[mat1].Ordinal;
                        index2 = dt_t.Columns[mat2].Ordinal;

                        if (Convert.ToString(dt_t.Rows[index1 - 1][index2]) == "T")//first row is 1 not zero, first column is NA so is zero
                        {
                            dt1.Rows.Add();
                            dt1.Rows[dt1.Rows.Count - 1][0] = dt_comp.Rows[i][col_sta2];
                        }
                    }
                }
            }
            return dt1;
        }

        private string get_coating(System.Data.DataTable dt_mat, string mat1)
        {
            string coating = "FBE";
            string col_coating = "Coating";


            if (dt_mat != null && dt_mat.Rows.Count > 0 && dt_mat.Columns.Contains(col_pipe_type) == true && dt_mat.Columns.Contains(col_coating) == true)
            {
                for (int i = 0; i < dt_mat.Rows.Count; ++i)
                {
                    if (dt_mat.Rows[i][col_pipe_type] != DBNull.Value && dt_mat.Rows[i][col_coating] != DBNull.Value)
                    {
                        string mat0 = Convert.ToString(dt_mat.Rows[i][col_pipe_type]);
                        string coat0 = Convert.ToString(dt_mat.Rows[i][col_coating]);
                        if (mat0.ToLower() == mat1.ToLower())
                        {
                            coating = coat0;
                            i = dt_mat.Rows.Count;
                        }
                    }
                }
            }
            return coating;
        }
        private bool check_wall_thickness(System.Data.DataTable dt1, System.Data.DataTable dt_materials, string col_pipe_type, string col_wt, int start1, ref int line1)
        {
            if (dt1.Rows.Count == 0) return true;
            DataSet dataset1 = new DataSet();
            dataset1.Tables.Add(dt1);
            dataset1.Tables.Add(dt_materials);

            DataRelation relation1 = new DataRelation("xxx", dt1.Columns[col_pipe_type], dt_materials.Columns[col_pipe_type], false);
            dataset1.Relations.Add(relation1);

            for (int i = 0; i < dt1.Rows.Count; ++i)
            {
                if (dt1.Rows[i].GetChildRows(relation1).Length == 1)
                {
                    if (dt1.Rows[i][col_pipe_type] != DBNull.Value && dt1.Rows[i][col_wt] != DBNull.Value && dt1.Rows[i].GetChildRows(relation1)[0][col_wt] != DBNull.Value)
                    {
                        string mat1 = Convert.ToString(dt1.Rows[i][col_pipe_type]);
                        double wt1 = Convert.ToDouble(dt1.Rows[i][col_wt]);
                        double wt2 = Convert.ToDouble(dt1.Rows[i].GetChildRows(relation1)[0][col_wt]);

                        if (wt1 != wt2)
                        {
                            if (dt1.Rows[i][excel_cell] != DBNull.Value)
                            {
                                line1 = start1 + Convert.ToInt32(dt1.Rows[i][excel_cell]) - 1;
                            }

                            if (dt1.Columns.Contains(excel_cell1) == true)
                            {
                                if (dt1.Rows[i][excel_cell1] != DBNull.Value)
                                {
                                    line1 = start1 + Convert.ToInt32(dt1.Rows[i][excel_cell1]) - 1;
                                }
                            }
                            dataset1.Relations.Remove(relation1);
                            dataset1.Tables.Remove(dt1);
                            dataset1.Tables.Remove(dt_materials);
                            return false;
                        }
                    }
                    else
                    {
                        line1 = start1 + Convert.ToInt32(dt1.Rows[i][excel_cell]) - 1;
                        return false;
                    }
                }
                else
                {
                    if (dt1.Rows[i][col_wt] == DBNull.Value) MessageBox.Show("No wall tickness");
                    if (dt1.Rows[i][col_pipe_type] == DBNull.Value) MessageBox.Show("No material number");
                    string mat1 = Convert.ToString(dt1.Rows[i][col_pipe_type]);
                    double wt1 = Convert.ToDouble(dt1.Rows[i][col_wt]);

                    line1 = start1 + Convert.ToInt32(dt1.Rows[i][excel_cell]) - 1;
                    dataset1.Relations.Remove(relation1);
                    dataset1.Tables.Remove(dt1);
                    dataset1.Tables.Remove(dt_materials);
                    return false;
                }
            }

            dataset1.Relations.Remove(relation1);
            dataset1.Tables.Remove(dt1);
            dataset1.Tables.Remove(dt_materials);

            return true;
        }

        private bool check_overlaps(System.Data.DataTable dt1, string col_start, string col_end, int start1, ref int line1)
        {
            for (int i = 0; i < dt1.Rows.Count - 1; ++i)
            {
                if (dt1.Rows[i][col_end] != DBNull.Value)
                {
                    double sta2 = Convert.ToDouble(dt1.Rows[i][col_end]);
                    double sta11 = -123.4567;

                    for (int j = i + 1; j < dt1.Rows.Count; ++j)
                    {
                        if (dt1.Rows[j][col_start] != DBNull.Value)
                        {
                            sta11 = Convert.ToDouble(dt1.Rows[j][col_start]);
                            j = dt1.Rows.Count;
                        }
                    }


                    if (sta11 != -123.4567)
                    {
                        if (sta11 < sta2)
                        {
                            if (dt1.Rows[i][excel_cell] != DBNull.Value)
                            {
                                line1 = start1 + Convert.ToInt32(dt1.Rows[i][excel_cell]) - 1;
                            }


                            if (dt1.Columns.Contains(excel_cell1) == true)
                            {
                                if (dt1.Rows[i][excel_cell1] != DBNull.Value)
                                {
                                    line1 = start1 + Convert.ToInt32(dt1.Rows[i][excel_cell1]) - 1;
                                }
                            }


                            return false;
                        }
                    }

                }

            }
            return true;
        }

        private bool check_sta2_bigger_than_sta1(System.Data.DataTable dt1, string col1, string col2, int start1, ref int line1)
        {
            for (int i = 0; i < dt1.Rows.Count; ++i)
            {

                if (dt1.Rows[i][col1] != DBNull.Value && dt1.Rows[i][col2] != DBNull.Value)
                {
                    double sta2 = Convert.ToDouble(dt1.Rows[i][col2]);
                    double sta1 = Convert.ToDouble(dt1.Rows[i][col1]);



                    if (sta1 > sta2)
                    {
                        if (dt1.Rows[i][excel_cell] != DBNull.Value)
                        {
                            line1 = start1 + Convert.ToInt32(dt1.Rows[i][excel_cell]) - 1;
                        }


                        if (dt1.Columns.Contains(excel_cell1) == true)
                        {
                            if (dt1.Rows[i][excel_cell1] != DBNull.Value)
                            {
                                line1 = start1 + Convert.ToInt32(dt1.Rows[i][excel_cell1]) - 1;
                            }
                        }

                        return false;
                    }
                }

            }
            return true;
        }


        private bool check_gaps_and_overlaps(System.Data.DataTable dt1, int start1, ref int line1)
        {
            for (int i = 0; i < dt1.Rows.Count - 1; ++i)
            {
                double sta2 = Convert.ToDouble(dt1.Rows[i][col_sta2]);
                double sta11 = Convert.ToDouble(dt1.Rows[i + 1][col_sta1]);
                if (sta11 != sta2)
                {
                    line1 = start1 + Convert.ToInt32(dt1.Rows[i][excel_cell]) - 1;

                    return false;
                }
            }
            return true;
        }

        private bool check_overlaps_for_elbows(System.Data.DataTable dt1, int start1, ref int line1, ref string comment)
        {
            for (int i = 0; i < dt1.Rows.Count - 1; ++i)
            {

                double sta02 = -1;
                double sta011 = -1;
                if (dt1.Rows[i][col_end_elbow_adjacent] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt1.Rows[i][col_end_elbow_adjacent])) == true && Convert.ToString(dt1.Rows[i][col_end_elbow_adjacent]) != "-")
                {
                    sta02 = Convert.ToDouble(dt1.Rows[i][col_end_elbow_adjacent]);
                }

                if (dt1.Rows[i + 1][col_start_elbow_adjacent] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt1.Rows[i + 1][col_start_elbow_adjacent])) == true && Convert.ToString(dt1.Rows[i + 1][col_start_elbow_adjacent]) != "-")
                {
                    sta011 = Convert.ToDouble(dt1.Rows[i + 1][col_start_elbow_adjacent]);
                }


                double sta2 = Convert.ToDouble(dt1.Rows[i][col_sta2]);
                double sta11 = Convert.ToDouble(dt1.Rows[i + 1][col_sta1]);

                if (sta2 > sta11)
                {
                    line1 = start1 + Convert.ToInt32(dt1.Rows[i][excel_cell]) - 1;

                    comment = "error at " + Convert.ToString(sta2) + " has to be smaller or equal than " + Convert.ToString(sta11);
                    return false;


                }

                if (sta011 >= 0)
                {
                    if (sta2 > sta011)
                    {
                        line1 = start1 + Convert.ToInt32(dt1.Rows[i][excel_cell]) - 1;

                        comment = "error at " + Convert.ToString(sta2) + " has to be smaller or equal than " + Convert.ToString(sta011);
                        return false;
                    }
                    if (sta011 >= sta11)
                    {
                        line1 = start1 + Convert.ToInt32(dt1.Rows[i][excel_cell]) - 1;

                        comment = "error at " + Convert.ToString(sta011) + " has to be bigger than " + Convert.ToString(sta11);
                        return false;
                    }
                }


                if (sta02 >= 0)
                {
                    if (sta02 > sta11)
                    {
                        line1 = start1 + Convert.ToInt32(dt1.Rows[i][excel_cell]) - 1;
                        comment = "error at " + Convert.ToString(sta02) + " has to be smaller or equal than " + Convert.ToString(sta11);
                        return false;
                    }
                    if (sta2 >= sta02)
                    {
                        line1 = start1 + Convert.ToInt32(dt1.Rows[i][excel_cell]) - 1;
                        comment = "error at " + Convert.ToString(sta2) + " has to be smaller than " + Convert.ToString(sta02);
                        return false;
                    }

                    if (sta011 >= 0)
                    {
                        if (sta02 > sta011)
                        {
                            line1 = start1 + Convert.ToInt32(dt1.Rows[i][excel_cell]) - 1;
                            comment = "error at " + Convert.ToString(sta02) + " has to be smaller or equal than " + Convert.ToString(sta011);
                            return false;
                        }
                    }

                }

            }
            return true;
        }


        public void create_pipe_od_table()
        {

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



                            List1.Add(col_od_pipe_type);
                            List2.Add("Material");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            List1.Add(col_od_wt);
                            List2.Add("Wall Thickness");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            List1.Add(col_od_coat);
                            List2.Add("Coating");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            List1.Add(col_od_class);
                            List2.Add("Class");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            List1.Add(col_od_mat_descr);
                            List2.Add("Description");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            List1.Add(col_od_descr);
                            List2.Add("Description");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            List1.Add(col_od_start);
                            List2.Add("Start");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            List1.Add(col_od_end);
                            List2.Add("End");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            List1.Add(col_len);
                            List2.Add("Length");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);


                            List1.Add(col_notes);
                            List2.Add("pipe notes");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            Functions.Get_object_data_table(pipes_od, "Generated by MD", List1, List2, List3);


                            Trans1.Commit();
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

        }
        private void attach_od_to_pipes(System.Data.DataTable dt_od)
        {

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                if (dt_od.Rows.Count > 0)
                {
                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                    BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                    for (int i = 0; i < dt_od.Rows.Count; ++i)
                    {

                        ObjectId id1 = (ObjectId)dt_od.Rows[i]["id"];

                        List<object> lista_val = new List<object>();

                        string pipe_type = null;
                        if (dt_od.Rows[i][col_od_pipe_type] != DBNull.Value)
                        {
                            pipe_type = Convert.ToString(dt_od.Rows[i][col_od_pipe_type]);
                        }
                        lista_val.Add(pipe_type);

                        string od_wt = null;
                        if (dt_od.Rows[i][col_od_wt] != DBNull.Value)
                        {
                            od_wt = Convert.ToString(dt_od.Rows[i][col_od_wt]);
                        }
                        lista_val.Add(od_wt);

                        string od_coat = null;
                        if (dt_od.Rows[i][col_od_coat] != DBNull.Value)
                        {
                            od_coat = Convert.ToString(dt_od.Rows[i][col_od_coat]);
                        }
                        lista_val.Add(od_coat);


                        string od_class = null;
                        if (dt_od.Rows[i][col_od_class] != DBNull.Value)
                        {
                            od_class = Convert.ToString(dt_od.Rows[i][col_od_class]);
                        }
                        lista_val.Add(od_class);


                        string od_mat_descr = null;
                        if (dt_od.Rows[i][col_od_mat_descr] != DBNull.Value)
                        {
                            od_mat_descr = Convert.ToString(dt_od.Rows[i][col_od_mat_descr]);
                        }
                        lista_val.Add(od_mat_descr);


                        string od_descr = null;
                        if (dt_od.Rows[i][col_od_descr] != DBNull.Value)
                        {
                            od_descr = Convert.ToString(dt_od.Rows[i][col_od_descr]);
                        }
                        lista_val.Add(od_descr);


                        string od_start = null;
                        if (dt_od.Rows[i][col_od_start] != DBNull.Value)
                        {
                            od_start = Convert.ToString(dt_od.Rows[i][col_od_start]);
                        }
                        lista_val.Add(od_start);


                        string od_end = null;
                        if (dt_od.Rows[i][col_od_end] != DBNull.Value)
                        {
                            od_end = Convert.ToString(dt_od.Rows[i][col_od_end]);
                        }
                        lista_val.Add(od_end);

                        string od_len = null;
                        if (dt_od.Rows[i][col_len] != DBNull.Value)
                        {
                            od_len = Convert.ToString(dt_od.Rows[i][col_len]);
                        }
                        lista_val.Add(od_len);

                        string od_notes = null;
                        if (dt_od.Rows[i][col_notes] != DBNull.Value)
                        {
                            od_notes = Convert.ToString(dt_od.Rows[i][col_notes]);
                        }
                        lista_val.Add(od_notes);

                        Polyline atws1 = Trans1.GetObject(id1, OpenMode.ForWrite) as Polyline;


                        List<Autodesk.Gis.Map.Constants.DataType> lista_types = new List<Autodesk.Gis.Map.Constants.DataType>();

                        for (int k = 1; k <= lista_val.Count; ++k)
                        {
                            lista_types.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                        }

                        Functions.add_od_table_to_object(id1, pipes_od, lista_val, lista_types);
                    }
                }
                Trans1.Commit();
            }

        }

        public void create_elbow_od_table()
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

                        List1.Add(col_od_elbow_id);
                        List2.Add("Elbow ID");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);


                        List1.Add(col_od_elbow_angle);
                        List2.Add("Elbow Angle");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_od_bend_type);
                        List2.Add("Elbow bend type");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_od_descr);
                        List2.Add("Description");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_od_elbow_ref_dwg);
                        List2.Add("Elbow Reference dwg");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_elbow_pi);
                        List2.Add("Elbow PI");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_od_start);
                        List2.Add("Elbow Start");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_od_end);
                        List2.Add("Elbow End");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_length);
                        List2.Add("Elbow length");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_od_pipe_type);
                        List2.Add("Elbow type");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_od_pipe_class);
                        List2.Add("Elbow class");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_od_wt);
                        List2.Add("Elbow wall thickness");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_elbow_notes);
                        List2.Add("Elbow notes");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        Functions.Get_object_data_table(elbows_od, "Generated by MD", List1, List2, List3);


                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }




        private void attach_od_to_objects(System.Data.DataTable dt_od, string name_of_table)
        {

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                if (dt_od.Rows.Count > 0)
                {
                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                    BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                    for (int i = 0; i < dt_od.Rows.Count; ++i)
                    {

                        ObjectId id1 = (ObjectId)dt_od.Rows[i]["id"];

                        List<object> lista_val = new List<object>();

                        for (int j = 0; j < dt_od.Columns.Count - 1; ++j)
                        {
                            string val1 = null;
                            if (dt_od.Rows[i][j] != DBNull.Value)
                            {
                                val1 = Convert.ToString(dt_od.Rows[i][j]);
                            }
                            lista_val.Add(val1);
                        }


                        List<Autodesk.Gis.Map.Constants.DataType> lista_types = new List<Autodesk.Gis.Map.Constants.DataType>();

                        for (int k = 1; k <= lista_val.Count; ++k)
                        {
                            lista_types.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                        }


                        Polyline poly1 = Trans1.GetObject(id1, OpenMode.ForWrite) as Polyline;
                        DBPoint pt1 = Trans1.GetObject(id1, OpenMode.ForWrite) as DBPoint;
                        if (poly1 != null || pt1 != null)
                        {
                            Functions.add_od_table_to_object(id1, name_of_table, lista_val, lista_types);
                        }


                    }
                }
                Trans1.Commit();
            }

        }

        private void attach_od_to_geotech(System.Data.DataTable dt_od, string name_of_table)
        {

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                if (dt_od.Rows.Count > 0)
                {
                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                    BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                    for (int i = 0; i < dt_od.Rows.Count; ++i)
                    {

                        ObjectId id1 = (ObjectId)dt_od.Rows[i]["id"];

                        List<object> lista_val = new List<object>();

                        for (int j = 0; j < dt_od.Columns.Count - 3; ++j)
                        {
                            string val1 = null;
                            if (dt_od.Rows[i][j] != DBNull.Value)
                            {
                                val1 = Convert.ToString(dt_od.Rows[i][j]);
                            }
                            lista_val.Add(val1);
                        }


                        List<Autodesk.Gis.Map.Constants.DataType> lista_types = new List<Autodesk.Gis.Map.Constants.DataType>();

                        for (int k = 1; k <= lista_val.Count; ++k)
                        {
                            lista_types.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                        }


                        Polyline poly1 = Trans1.GetObject(id1, OpenMode.ForWrite) as Polyline;
                        DBPoint pt1 = Trans1.GetObject(id1, OpenMode.ForWrite) as DBPoint;
                        if (poly1 != null || pt1 != null)
                        {
                            Functions.add_od_table_to_object(id1, name_of_table, lista_val, lista_types);
                        }


                    }
                }
                Trans1.Commit();
            }

        }

        private void attach_od_to_geotech_start(System.Data.DataTable dt_od)
        {

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                if (dt_od.Rows.Count > 0)
                {
                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                    BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;


                    for (int i = 0; i < dt_od.Rows.Count; ++i)
                    {


                        ObjectId id2 = (ObjectId)dt_od.Rows[i]["id1"];


                        List<object> lista_val1 = new List<object>();


                        string f1 = null;
                        if (dt_od.Rows[i][col_geotech_od_sta1] != DBNull.Value)
                        {
                            f1 = Convert.ToString(dt_od.Rows[i][col_geotech_od_sta1]);
                        }

                        lista_val1.Add(f1);


                        string descr = null;
                        if (dt_od.Rows[i][col_geotech_od_descr1] != DBNull.Value)
                        {
                            descr = Convert.ToString(dt_od.Rows[i][col_geotech_od_descr1]);
                        }

                        lista_val1.Add(descr);



                        string j1 = null;
                        if (dt_od.Rows[i][col_geotech_od_class] != DBNull.Value)
                        {
                            j1 = Convert.ToString(dt_od.Rows[i][col_geotech_od_class]);
                        }

                        lista_val1.Add(j1);

                        string j2 = null;
                        if (dt_od.Rows[i][col_geotech_od_type] != DBNull.Value)
                        {
                            j2 = Convert.ToString(dt_od.Rows[i][col_geotech_od_type]);
                        }

                        lista_val1.Add(j2);

                        string j3 = null;
                        if (dt_od.Rows[i][col_geotech_od_label] != DBNull.Value)
                        {
                            j3 = Convert.ToString(dt_od.Rows[i][col_geotech_od_label]);
                        }

                        lista_val1.Add(j3);


                        string n1 = null;
                        if (dt_od.Rows[i][col_notes] != DBNull.Value)
                        {
                            n1 = Convert.ToString(dt_od.Rows[i][col_notes]);
                        }

                        lista_val1.Add(n1);



                        DBPoint dbp1 = Trans1.GetObject(id2, OpenMode.ForWrite) as DBPoint;


                        if (dbp1 != null)
                        {
                            List<Autodesk.Gis.Map.Constants.DataType> lista_types = new List<Autodesk.Gis.Map.Constants.DataType>();

                            for (int k = 1; k <= lista_val1.Count; ++k)
                            {
                                lista_types.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                            }

                            Functions.add_od_table_to_object(id2, geotech_od_pt, lista_val1, lista_types);
                        }


                    }
                }
                Trans1.Commit();
            }

        }

        private void attach_od_to_geotech_end(System.Data.DataTable dt_od)
        {

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                if (dt_od.Rows.Count > 0)
                {
                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                    BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;


                    for (int i = 0; i < dt_od.Rows.Count; ++i)
                    {


                        ObjectId id2 = (ObjectId)dt_od.Rows[i]["id2"];


                        List<object> lista_val1 = new List<object>();


                        string f1 = null;
                        if (dt_od.Rows[i][col_geotech_od_sta2] != DBNull.Value)
                        {
                            f1 = Convert.ToString(dt_od.Rows[i][col_geotech_od_sta2]);
                        }

                        lista_val1.Add(f1);


                        string descr = null;
                        if (dt_od.Rows[i][col_geotech_od_descr2] != DBNull.Value)
                        {
                            descr = Convert.ToString(dt_od.Rows[i][col_geotech_od_descr2]);
                        }

                        lista_val1.Add(descr);



                        string j1 = null;
                        if (dt_od.Rows[i][col_geotech_od_class] != DBNull.Value)
                        {
                            j1 = Convert.ToString(dt_od.Rows[i][col_geotech_od_class]);
                        }

                        lista_val1.Add(j1);

                        string j2 = null;
                        if (dt_od.Rows[i][col_geotech_od_type] != DBNull.Value)
                        {
                            j2 = Convert.ToString(dt_od.Rows[i][col_geotech_od_type]);
                        }

                        lista_val1.Add(j2);

                        string j3 = null;
                        if (dt_od.Rows[i][col_geotech_od_label] != DBNull.Value)
                        {
                            j3 = Convert.ToString(dt_od.Rows[i][col_geotech_od_label]);
                        }

                        lista_val1.Add(j3);


                        string n1 = null;
                        if (dt_od.Rows[i][col_notes] != DBNull.Value)
                        {
                            n1 = Convert.ToString(dt_od.Rows[i][col_notes]);
                        }

                        lista_val1.Add(n1);



                        DBPoint dbp1 = Trans1.GetObject(id2, OpenMode.ForWrite) as DBPoint;


                        if (dbp1 != null)
                        {
                            List<Autodesk.Gis.Map.Constants.DataType> lista_types = new List<Autodesk.Gis.Map.Constants.DataType>();

                            for (int k = 1; k <= lista_val1.Count; ++k)
                            {
                                lista_types.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                            }

                            Functions.add_od_table_to_object(id2, geotech_od_pt, lista_val1, lista_types);
                        }


                    }
                }
                Trans1.Commit();
            }

        }


        public void create_facility_od_table()
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


                        List1.Add(col_fac_name);
                        List2.Add("Name");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_start);
                        List2.Add("Assembly start");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);


                        List1.Add(col_end);
                        List2.Add("Assembly End");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_len);
                        List2.Add("length");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);


                        List1.Add(col_descr1);
                        List2.Add("Description1");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_descr2);
                        List2.Add("Description2");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);


                        List1.Add(col_just);
                        List2.Add("Justification");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        Functions.Get_object_data_table(fab_od, "Generated by MD", List1, List2, List3);


                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void create_pre_existing_od_table()
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

                        List1.Add(col_start);
                        List2.Add("Pre-Existing start");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);


                        List1.Add(col_end);
                        List2.Add("Pre-Existing End");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);


                        List1.Add(col_len);
                        List2.Add("Length");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_descr);
                        List2.Add("Description");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_just);
                        List2.Add("Justification");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);


                        List1.Add(col_notes);
                        List2.Add("Notes");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        Functions.Get_object_data_table(preexisting_od, "Generated by MD", List1, List2, List3);


                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void create_class_od_table()
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


                        List1.Add(col_od_pipe_type);
                        List2.Add("Pipe type");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_start);
                        List2.Add("Assembly start");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_od_descr1);
                        List2.Add("Description");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);


                        List1.Add(col_end);
                        List2.Add("Assembly End");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_od_descr2);
                        List2.Add("Description");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);


                        List1.Add(col_od_wt);
                        List2.Add("Wall Thickness");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_just);
                        List2.Add("Justification");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        Functions.Get_object_data_table(class_od, "Generated by MD", List1, List2, List3);


                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void create_buoyancy_od_table()
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



                        List1.Add(col_feature);
                        List2.Add("Feature Type");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_spacing);
                        List2.Add("Spacing");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_count);
                        List2.Add("Feature Count");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_start);
                        List2.Add("Feature Start");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_end);
                        List2.Add("Feature End");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_length);
                        List2.Add("Feature length");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_just);
                        List2.Add("Justification");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_notes);
                        List2.Add("Notes");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);


                        Functions.Get_object_data_table(buoyancy_od, "Generated by MD", List1, List2, List3);




                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }


        public void create_buoyancy_pt_od_table()
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



                        List1.Add(col_feature);
                        List2.Add("Feature Type");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_station);
                        List2.Add("Feature Station");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_just);
                        List2.Add("Justification");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_notes);
                        List2.Add("Notes");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);


                        Functions.Get_object_data_table(buoyancy_pt_od, "Generated by MD", List1, List2, List3);




                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        public void create_geotech_od_table()
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



                        List1.Add(col_geotech_od_sta1);
                        List2.Add("Start");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_geotech_od_descr1);
                        List2.Add("Description start");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_geotech_od_sta2);
                        List2.Add("End");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_geotech_od_descr2);
                        List2.Add("Description End");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_len);
                        List2.Add("Length");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_geotech_od_class);
                        List2.Add("Geohazard Class");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_geotech_od_type);
                        List2.Add("Geohazard Type");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);


                        List1.Add(col_geotech_od_label);
                        List2.Add("Golder Hazard Label");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_notes);
                        List2.Add("Notes");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);


                        Functions.Get_object_data_table(geotech_od, "Generated by MD", List1, List2, List3);




                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        public void create_geotech_pt_od_table()
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


                        List1.Add(col_station);
                        List2.Add("Station");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_descr);
                        List2.Add("Description");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);


                        List1.Add(col_geotech_od_class);
                        List2.Add("Geohazard Class");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_geotech_od_type);
                        List2.Add("Geohazard Type");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);


                        List1.Add(col_geotech_od_label);
                        List2.Add("Golder Hazard Label");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_notes);
                        List2.Add("Notes");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);


                        Functions.Get_object_data_table(geotech_od_pt, "Generated by MD", List1, List2, List3);




                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        public void create_doc_od_table()
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



                        List1.Add(col_doc_od_sta1);
                        List2.Add("Start");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_doc_od_sta2);
                        List2.Add("End");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_doc_od_min_cvr);
                        List2.Add("Minimum Depth of Cover");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_len);
                        List2.Add("Length");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);


                        List1.Add(col_just);
                        List2.Add("Justification");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_notes);
                        List2.Add("Notes");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);


                        Functions.Get_object_data_table(doc_od, "Generated by MD", List1, List2, List3);




                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        public void create_muskeg_od_table()
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



                        List1.Add(col_station);
                        List2.Add("Station");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_descr);
                        List2.Add("Description");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_muskeg_od_label);
                        List2.Add("Golder Label");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);



                        Functions.Get_object_data_table(muskeg_od, "Generated by MD", List1, List2, List3);




                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        private void attach_od_to_buoyancy(System.Data.DataTable dt_od)
        {

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                if (dt_od.Rows.Count > 0)
                {
                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                    BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;


                    for (int i = 0; i < dt_od.Rows.Count; ++i)
                    {

                        ObjectId id1 = (ObjectId)dt_od.Rows[i]["id"];
                        ObjectId id2 = (ObjectId)dt_od.Rows[i]["id1"];
                        ObjectId id3 = (ObjectId)dt_od.Rows[i]["id2"];

                        List<object> lista_val = new List<object>();
                        List<object> lista_val1 = new List<object>();
                        List<object> lista_val2 = new List<object>();

                        string f1 = null;
                        if (dt_od.Rows[i][col_feature] != DBNull.Value)
                        {
                            f1 = Convert.ToString(dt_od.Rows[i][col_feature]);
                        }
                        lista_val.Add(f1);
                        lista_val1.Add(f1);
                        lista_val2.Add(f1);



                        string spp1 = null;
                        if (dt_od.Rows[i][col_spacing] != DBNull.Value)
                        {
                            spp1 = Convert.ToString(dt_od.Rows[i][col_spacing]);
                        }
                        lista_val.Add(spp1);

                        string count1 = null;
                        if (dt_od.Rows[i][col_count] != DBNull.Value)
                        {
                            count1 = Convert.ToString(dt_od.Rows[i][col_count]);
                        }
                        lista_val.Add(count1);

                        string od_start = null;
                        if (dt_od.Rows[i][col_start] != DBNull.Value)
                        {
                            od_start = Convert.ToString(dt_od.Rows[i][col_start]);
                        }
                        lista_val.Add(od_start);
                        lista_val1.Add(od_start);
                        lista_val2.Add(od_start);

                        string od_end = null;
                        if (dt_od.Rows[i][col_end] != DBNull.Value)
                        {
                            od_end = Convert.ToString(dt_od.Rows[i][col_end]);
                        }
                        lista_val.Add(od_end);
                        lista_val1.Add(od_end);
                        lista_val2.Add(od_end);

                        string l1 = null;
                        if (dt_od.Rows[i][col_len] != DBNull.Value)
                        {
                            l1 = Convert.ToString(dt_od.Rows[i][col_len]);
                        }
                        lista_val.Add(l1);


                        string j1 = null;
                        if (dt_od.Rows[i][col_just] != DBNull.Value)
                        {
                            j1 = Convert.ToString(dt_od.Rows[i][col_just]);
                        }
                        lista_val.Add(j1);
                        lista_val1.Add(j1);
                        lista_val2.Add(j1);

                        string n1 = null;
                        if (dt_od.Rows[i][col_notes] != DBNull.Value)
                        {
                            n1 = Convert.ToString(dt_od.Rows[i][col_notes]);
                        }
                        lista_val.Add(n1);
                        lista_val1.Add(n1);
                        lista_val2.Add(n1);


                        Polyline atws1 = Trans1.GetObject(id1, OpenMode.ForWrite) as Polyline;


                        if (atws1 != null)
                        {
                            List<Autodesk.Gis.Map.Constants.DataType> lista_types = new List<Autodesk.Gis.Map.Constants.DataType>();

                            for (int k = 1; k <= lista_val.Count; ++k)
                            {
                                lista_types.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                            }

                            Functions.add_od_table_to_object(id1, buoyancy_od, lista_val, lista_types);
                        }



                    }
                }
                Trans1.Commit();
            }

        }

        private void attach_od_to_buoyancy_start(System.Data.DataTable dt_od)
        {

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                if (dt_od.Rows.Count > 0)
                {
                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                    BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;


                    for (int i = 0; i < dt_od.Rows.Count; ++i)
                    {


                        ObjectId id2 = (ObjectId)dt_od.Rows[i]["id1"];


                        List<object> lista_val1 = new List<object>();


                        string f1 = null;
                        if (dt_od.Rows[i][col_feature] != DBNull.Value)
                        {
                            f1 = Convert.ToString(dt_od.Rows[i][col_feature]);
                        }

                        lista_val1.Add(f1);


                        string od_start = null;
                        if (dt_od.Rows[i][col_start] != DBNull.Value)
                        {
                            od_start = Convert.ToString(dt_od.Rows[i][col_start]);
                        }

                        lista_val1.Add(od_start);



                        string j1 = null;
                        if (dt_od.Rows[i][col_just] != DBNull.Value)
                        {
                            j1 = Convert.ToString(dt_od.Rows[i][col_just]);
                        }

                        lista_val1.Add(j1);


                        string n1 = null;
                        if (dt_od.Rows[i][col_notes] != DBNull.Value)
                        {
                            n1 = Convert.ToString(dt_od.Rows[i][col_notes]);
                        }

                        lista_val1.Add(n1);



                        DBPoint dbp1 = Trans1.GetObject(id2, OpenMode.ForWrite) as DBPoint;


                        if (dbp1 != null)
                        {
                            List<Autodesk.Gis.Map.Constants.DataType> lista_types = new List<Autodesk.Gis.Map.Constants.DataType>();

                            for (int k = 1; k <= lista_val1.Count; ++k)
                            {
                                lista_types.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                            }

                            Functions.add_od_table_to_object(id2, buoyancy_pt_od, lista_val1, lista_types);
                        }


                    }
                }
                Trans1.Commit();
            }

        }

        private void attach_od_to_buoyancy_end(System.Data.DataTable dt_od)
        {

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                if (dt_od.Rows.Count > 0)
                {
                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                    BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;


                    for (int i = 0; i < dt_od.Rows.Count; ++i)
                    {


                        ObjectId id3 = (ObjectId)dt_od.Rows[i]["id2"];


                        List<object> lista_val2 = new List<object>();

                        string f1 = null;
                        if (dt_od.Rows[i][col_feature] != DBNull.Value)
                        {
                            f1 = Convert.ToString(dt_od.Rows[i][col_feature]);
                        }

                        lista_val2.Add(f1);



                        string od_end = null;
                        if (dt_od.Rows[i][col_end] != DBNull.Value)
                        {
                            od_end = Convert.ToString(dt_od.Rows[i][col_end]);
                        }

                        lista_val2.Add(od_end);


                        string j1 = null;
                        if (dt_od.Rows[i][col_just] != DBNull.Value)
                        {
                            j1 = Convert.ToString(dt_od.Rows[i][col_just]);
                        }

                        lista_val2.Add(j1);

                        string n1 = null;
                        if (dt_od.Rows[i][col_notes] != DBNull.Value)
                        {
                            n1 = Convert.ToString(dt_od.Rows[i][col_notes]);
                        }

                        lista_val2.Add(n1);



                        DBPoint dbp2 = Trans1.GetObject(id3, OpenMode.ForWrite) as DBPoint;



                        if (dbp2 != null)
                        {
                            List<Autodesk.Gis.Map.Constants.DataType> lista_types = new List<Autodesk.Gis.Map.Constants.DataType>();

                            for (int k = 1; k <= lista_val2.Count; ++k)
                            {
                                lista_types.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                            }

                            Functions.add_od_table_to_object(id3, buoyancy_pt_od, lista_val2, lista_types);
                        }

                    }
                }
                Trans1.Commit();
            }

        }


        public void create_cpac_od_table()
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



                        List1.Add(col_station);
                        List2.Add("Station");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_descr);
                        List2.Add("Description");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_descr2);
                        List2.Add("Description2");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);


                        List1.Add(col_eq1);
                        List2.Add("Equipment type 1");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_eq2);
                        List2.Add("Equipment type 2");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_eq3);
                        List2.Add("Equipment type 3");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_just);
                        List2.Add("Justification");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_notes);
                        List2.Add("Notes");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        Functions.Get_object_data_table(cpac_od, "Generated by MD", List1, List2, List3);




                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        private void attach_od_to_cpac(System.Data.DataTable dt_od)
        {

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                if (dt_od.Rows.Count > 0)
                {
                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                    BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;


                    for (int i = 0; i < dt_od.Rows.Count; ++i)
                    {

                        ObjectId id1 = (ObjectId)dt_od.Rows[i]["id"];

                        List<object> lista_val = new List<object>();

                        string sta = null;
                        if (dt_od.Rows[i][col_sta] != DBNull.Value)
                        {
                            sta = Convert.ToString(dt_od.Rows[i][col_sta]);
                        }
                        lista_val.Add(sta);



                        string descr = null;
                        if (dt_od.Rows[i][col_descr] != DBNull.Value)
                        {
                            descr = Convert.ToString(dt_od.Rows[i][col_descr]);
                        }
                        lista_val.Add(descr);

                        string descr2 = null;
                        if (dt_od.Rows[i][col_descr2] != DBNull.Value)
                        {
                            descr2 = Convert.ToString(dt_od.Rows[i][col_descr2]);
                        }
                        lista_val.Add(descr2);

                        string eq1 = null;
                        if (dt_od.Rows[i][col_eq1] != DBNull.Value)
                        {
                            eq1 = Convert.ToString(dt_od.Rows[i][col_eq1]);
                        }
                        lista_val.Add(eq1);

                        string eq2 = null;
                        if (dt_od.Rows[i][col_eq2] != DBNull.Value)
                        {
                            eq2 = Convert.ToString(dt_od.Rows[i][col_eq2]);
                        }
                        lista_val.Add(eq2);

                        string eq3 = null;
                        if (dt_od.Rows[i][col_eq3] != DBNull.Value)
                        {
                            eq3 = Convert.ToString(dt_od.Rows[i][col_eq3]);
                        }
                        lista_val.Add(eq3);

                        string j1 = null;
                        if (dt_od.Rows[i][col_just] != DBNull.Value)
                        {
                            j1 = Convert.ToString(dt_od.Rows[i][col_just]);
                        }
                        lista_val.Add(j1);

                        string n1 = null;
                        if (dt_od.Rows[i][col_notes] != DBNull.Value)
                        {
                            n1 = Convert.ToString(dt_od.Rows[i][col_notes]);
                        }
                        lista_val.Add(n1);


                        DBPoint atws1 = Trans1.GetObject(id1, OpenMode.ForWrite) as DBPoint;


                        List<Autodesk.Gis.Map.Constants.DataType> lista_types = new List<Autodesk.Gis.Map.Constants.DataType>();

                        for (int k = 1; k <= lista_val.Count; ++k)
                        {
                            lista_types.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                        }

                        Functions.add_od_table_to_object(id1, cpac_od, lista_val, lista_types);
                    }
                }
                Trans1.Commit();
            }

        }

        public void create_es_od_table()
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



                        List1.Add(col_station);
                        List2.Add("Station");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_ditchplug);
                        List2.Add("Ditch Plug");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_es_spacing);
                        List2.Add("Feature Spacing");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_notes);
                        List2.Add("Notes");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        Functions.Get_object_data_table(es_od, "Generated by MD", List1, List2, List3);




                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        private void attach_od_to_es(System.Data.DataTable dt_od)
        {

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                if (dt_od.Rows.Count > 0)
                {
                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                    BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;


                    for (int i = 0; i < dt_od.Rows.Count; ++i)
                    {

                        ObjectId id1 = (ObjectId)dt_od.Rows[i]["id"];

                        List<object> lista_val = new List<object>();

                        string sta = null;
                        if (dt_od.Rows[i][col_sta] != DBNull.Value)
                        {
                            sta = Convert.ToString(dt_od.Rows[i][col_sta]);
                        }
                        lista_val.Add(sta);



                        string dp = null;
                        if (dt_od.Rows[i][col_ditchplug] != DBNull.Value)
                        {
                            dp = Convert.ToString(dt_od.Rows[i][col_ditchplug]);
                        }
                        lista_val.Add(dp);


                        string dp1 = null;
                        if (dt_od.Rows[i][col_es_spacing] != DBNull.Value)
                        {
                            dp1 = Convert.ToString(dt_od.Rows[i][col_es_spacing]);
                        }
                        lista_val.Add(dp1);

                        string nt1 = null;
                        if (dt_od.Rows[i][col_notes] != DBNull.Value)
                        {
                            nt1 = Convert.ToString(dt_od.Rows[i][col_notes]);
                        }
                        lista_val.Add(nt1);



                        DBPoint atws1 = Trans1.GetObject(id1, OpenMode.ForWrite) as DBPoint;


                        List<Autodesk.Gis.Map.Constants.DataType> lista_types = new List<Autodesk.Gis.Map.Constants.DataType>();

                        for (int k = 1; k <= lista_val.Count; ++k)
                        {
                            lista_types.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                        }

                        Functions.add_od_table_to_object(id1, es_od, lista_val, lista_types);
                    }
                }
                Trans1.Commit();
            }

        }

        public void create_hydrotest_point_od_table()
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



                        List1.Add(col_sta);
                        List2.Add("Station");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_descr);
                        List2.Add("Description");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        Functions.Get_object_data_table(hydrotest_odPT, "Generated by MD", List1, List2, List3);




                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        public void create_hydrotest_lines_od_table()
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



                        List1.Add(col_start);
                        List2.Add("Start");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_end);
                        List2.Add("End");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_len);
                        List2.Add("Length");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_descr);
                        List2.Add("Description");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        Functions.Get_object_data_table(hydrotest_od, "Generated by MD", List1, List2, List3);




                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        private void attach_od_to_hydrotestPT(System.Data.DataTable dt_od)
        {

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                if (dt_od.Rows.Count > 0)
                {
                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                    BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;


                    for (int i = 0; i < dt_od.Rows.Count; ++i)
                    {

                        ObjectId id1 = (ObjectId)dt_od.Rows[i]["id"];

                        List<object> lista_val = new List<object>();

                        string sta = null;
                        if (dt_od.Rows[i][col_sta] != DBNull.Value)
                        {
                            sta = Convert.ToString(dt_od.Rows[i][col_sta]);
                        }
                        lista_val.Add(sta);



                        string descr1 = null;
                        if (dt_od.Rows[i][col_descr] != DBNull.Value)
                        {
                            descr1 = Convert.ToString(dt_od.Rows[i][col_descr]);
                        }
                        lista_val.Add(descr1);





                        DBPoint atws1 = Trans1.GetObject(id1, OpenMode.ForWrite) as DBPoint;


                        List<Autodesk.Gis.Map.Constants.DataType> lista_types = new List<Autodesk.Gis.Map.Constants.DataType>();

                        for (int k = 1; k <= lista_val.Count; ++k)
                        {
                            lista_types.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                        }

                        Functions.add_od_table_to_object(id1, hydrotest_odPT, lista_val, lista_types);
                    }
                }
                Trans1.Commit();
            }

        }

        private void attach_od_to_hydrotest(System.Data.DataTable dt_od)
        {

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                if (dt_od.Rows.Count > 0)
                {
                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                    BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;


                    for (int i = 0; i < dt_od.Rows.Count; ++i)
                    {

                        ObjectId id1 = (ObjectId)dt_od.Rows[i]["id"];

                        List<object> lista_val = new List<object>();

                        string sta1 = null;
                        if (dt_od.Rows[i][col_start] != DBNull.Value)
                        {
                            sta1 = Convert.ToString(dt_od.Rows[i][col_start]);
                        }
                        lista_val.Add(sta1);

                        string sta2 = null;
                        if (dt_od.Rows[i][col_end] != DBNull.Value)
                        {
                            sta2 = Convert.ToString(dt_od.Rows[i][col_end]);
                        }
                        lista_val.Add(sta2);

                        string len1 = null;
                        if (dt_od.Rows[i][col_len] != DBNull.Value)
                        {
                            len1 = Convert.ToString(dt_od.Rows[i][col_len]);
                        }
                        lista_val.Add(len1);


                        string descr1 = null;
                        if (dt_od.Rows[i][col_descr] != DBNull.Value)
                        {
                            descr1 = Convert.ToString(dt_od.Rows[i][col_descr]);
                        }
                        lista_val.Add(descr1);


                        DBPoint atws1 = Trans1.GetObject(id1, OpenMode.ForWrite) as DBPoint;


                        List<Autodesk.Gis.Map.Constants.DataType> lista_types = new List<Autodesk.Gis.Map.Constants.DataType>();

                        for (int k = 1; k <= lista_val.Count; ++k)
                        {
                            lista_types.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                        }

                        Functions.add_od_table_to_object(id1, hydrotest_od, lista_val, lista_types);
                    }
                }
                Trans1.Commit();
            }

        }

        public void create_xing_od_table()
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

                        List1.Add(col_od_xingid);
                        List2.Add("Crossing ID");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_od_xingtype);
                        List2.Add("Crossing type");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_station);
                        List2.Add("Crossing station");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_descr1);
                        List2.Add("Crossing description line1");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_descr2);
                        List2.Add("Crossing description line2");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_od_ref_dwg_id);
                        List2.Add("Reference drawing ID");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_od_min_depth);
                        List2.Add("Minimum depth");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_od_xing_method);
                        List2.Add("Crossing method");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_od_pipe_type);
                        List2.Add("Pipe type");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_od_pipe_class);
                        List2.Add("Pipe class");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_od_wt);
                        List2.Add("Wall Thickness");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_just);
                        List2.Add("Justification");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        Functions.Get_object_data_table(xing_od, "Generated by MD", List1, List2, List3);




                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }


        public void create_transition_od_table()
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



                        List1.Add(col_sta);
                        List2.Add("Station");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col_descr);
                        List2.Add("Description");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);




                        Functions.Get_object_data_table(transition_od, "Generated by MD", List1, List2, List3);




                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        private void attach_od_to_transition(System.Data.DataTable dt_od)
        {

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                if (dt_od.Rows.Count > 0)
                {
                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                    BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;


                    for (int i = 0; i < dt_od.Rows.Count; ++i)
                    {

                        ObjectId id1 = (ObjectId)dt_od.Rows[i]["id"];

                        List<object> lista_val = new List<object>();

                        string sta = null;
                        if (dt_od.Rows[i][col_sta] != DBNull.Value)
                        {
                            sta = Convert.ToString(dt_od.Rows[i][col_sta]);
                        }
                        lista_val.Add(sta);


                        lista_val.Add("Transition");



                        DBPoint atws1 = Trans1.GetObject(id1, OpenMode.ForWrite) as DBPoint;


                        List<Autodesk.Gis.Map.Constants.DataType> lista_types = new List<Autodesk.Gis.Map.Constants.DataType>();

                        for (int k = 1; k <= lista_val.Count; ++k)
                        {
                            lista_types.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                        }

                        Functions.add_od_table_to_object(id1, transition_od, lista_val, lista_types);
                    }
                }
                Trans1.Commit();
            }

        }

        private void comboBox_hv_DropDown(object sender, EventArgs e)
        {
            
        }

        public double get_from_NPS_radius_for_pipes_from_inches_to_milimeters(double NPS_inches)
        {
            switch (NPS_inches)
            {
                case 0.125:
                    return 5.15;
                case 0.25:
                    return 6.85;
                case 0.375:
                    return 8.55;
                case 0.5:
                    return 10.65;
                case 0.75:
                    return 13.35;
                case 1:
                    return 16.7;
                case 1.25:
                    return 21.1;
                case 1.5:
                    return 24.15;
                case 2:
                    return 30.15;
                case 2.5:
                    return 36.5;
                case 3:
                    return 44.45;
                case 3.5:
                    return 50.8;
                case 4:
                    return 57.15;
                case 5:
                    return 70.65;
                case 6:
                    return 84.15;
                case 8:
                    return 109.55;
                case 10:
                    return 136.55;
                case 12:
                    return 161.95;
                case 14:
                    return 177.8;
                case 16:
                    return 203.2;
                case 18:
                    return 228.5;
                case 20:
                    return 254;
                case 22:
                    return 279.5;
                case 24:
                    return 304.8;
                case 26:
                    return 330;
                case 28:
                    return 355.5;
                case 30:
                    return 381;
                case 32:
                    return 406.5;
                case 34:
                    return 432;
                case 36:
                    return 457.2;
                case 38:
                    return 482.5;
                case 40:
                    return 508;
                case 42:
                    return 533.5;
                case 44:
                    return 559;
                case 46:
                    return 584;
                case 48:
                    return 609.5;
                case 50:
                    return 635;
                case 52:
                    return 660.5;
                case 54:
                    return 686;
                case 56:
                    return 711;
                case 58:
                    return 736.5;
                case 60:
                    return 762.0;
                case 62:
                    return 787.5;
                case 64:
                    return 813;
                case 66:
                    return 838.0;
                case 68:
                    return 863.5;
                case 70:
                    return 889;
                case 72:
                    return 914.5;
                case 74:
                    return 940;
                case 76:
                    return 965;
                case 78:
                    return 990.5;
                case 80:
                    return 1016;
                default:
                    return 0;

            }
        }

        private void ComboBox_nps_SelectedIndexChanged(object sender, EventArgs e)
        {
            string text1 = ComboBox_nps.Text;
            if (text1 == "NPS 1/2") text1 = "0.5";
            double val1 = Convert.ToDouble(text1.Replace("NPS ", ""));
            label_result_nps_radius.Text = Convert.ToString(get_from_NPS_radius_for_pipes_from_inches_to_milimeters(val1) / 1000);
            if (Functions.IsNumeric(textBox_field_bend_multiplier.Text) == true)
            {
                double fbm = Convert.ToDouble(textBox_field_bend_multiplier.Text);
                label_result_field_bend_radius.Text = Convert.ToString(2 * fbm * get_from_NPS_radius_for_pipes_from_inches_to_milimeters(val1) / 1000);

            }

            if (Functions.IsNumeric(textBox_elbow_multiplier.Text) == true)
            {
                double em = Convert.ToDouble(textBox_elbow_multiplier.Text);
                label_result_elbow_radius.Text = Convert.ToString(2 * em * get_from_NPS_radius_for_pipes_from_inches_to_milimeters(val1) / 1000);

            }
        }

        private void button_generate_angles_report_Click(object sender, EventArgs e)
        {
            try
            {

                string string1 = comboBox_horizontal.Text;
                string string2 = comboBox_vertical.Text;


                int startH1 = 0;
                int endH1 = 0;
                if (Functions.IsNumeric(textBoxH_start.Text) == true)
                {
                    startH1 = Convert.ToInt32(textBoxH_start.Text);
                }

                if (Functions.IsNumeric(textBoxH_end.Text) == true)
                {
                    endH1 = Convert.ToInt32(textBoxH_end.Text);
                }

                int startV1 = 0;
                int endV1 = 0;
                if (Functions.IsNumeric(textBoxV_start.Text) == true)
                {
                    startV1 = Convert.ToInt32(textBoxV_start.Text);
                }

                if (Functions.IsNumeric(textBoxV_end.Text) == true)
                {
                    endV1 = Convert.ToInt32(textBoxV_end.Text);
                }



                string nps_text = ComboBox_nps.Text;
                if (nps_text == "NPS 1/2") nps_text = "0.5";
                double nps_double = Convert.ToDouble(nps_text.Replace("NPS ", ""));

                double radius_of_nps = get_from_NPS_radius_for_pipes_from_inches_to_milimeters(nps_double) / 1000;

                double radius1 = 0;
                double radius2 = 0;


                double min_val = 0.1;
                if (Functions.IsNumeric(textBox_min_defl.Text) == true)
                {
                    min_val = Convert.ToDouble(textBox_min_defl.Text);
                }


                label_result_nps_radius.Text = Convert.ToString(radius_of_nps);
                if (Functions.IsNumeric(textBox_field_bend_multiplier.Text) == true)
                {
                    double fbm = Convert.ToDouble(textBox_field_bend_multiplier.Text);
                    radius1 = 2 * fbm * radius_of_nps;
                    label_result_field_bend_radius.Text = Convert.ToString(radius1);



                }

                if (Functions.IsNumeric(textBox_elbow_multiplier.Text) == true)
                {
                    double em = Convert.ToDouble(textBox_elbow_multiplier.Text);
                    radius2 = 2 * em * radius_of_nps;
                    label_result_elbow_radius.Text = Convert.ToString(radius2);
                }



                set_enable_false();

                System.Data.DataTable dt1 = new System.Data.DataTable();
                dt1.Columns.Add(col_sta, typeof(double));
                dt1.Columns.Add(col_defl, typeof(double));
                dt1.Columns.Add(col_is_elbow, typeof(bool));

                System.Data.DataTable dt2 = dt1.Clone();

                if ((string1.Contains("[") == true && string1.Contains("]") == true) || (string2.Contains("[") == true && string2.Contains("]") == true))
                {
                    string filename1 = string1.Substring(string1.IndexOf("]") + 4, string1.Length - (string1.IndexOf("]") + 4));
                    string filename2 = string2.Substring(string2.IndexOf("]") + 4, string2.Length - (string2.IndexOf("]") + 4));

                    string sheet_name1 = string1.Substring(1, string1.IndexOf("]") - 1);
                    string sheet_name2 = string2.Substring(1, string2.IndexOf("]") - 1);

                    Microsoft.Office.Interop.Excel.Worksheet W1 = null;
                    Microsoft.Office.Interop.Excel.Worksheet W2 = null;

                    if (filename1.Length > 0 && sheet_name1.Length > 0)
                    {
                        W1 = Functions.Get_opened_worksheet_from_Excel_by_name(filename1, sheet_name1);
                        if (W1 != null)
                        {
                            List<string> lista_col = new List<string>();
                            List<string> lista_colxl = new List<string>();
                            lista_col.Add(col_sta);
                            lista_col.Add(col_defl);
                            lista_col.Add(col_is_elbow);
                            lista_colxl.Add(textBoxH_sta.Text);
                            lista_colxl.Add(textBoxH_angle.Text);
                            lista_colxl.Add(textBoxH_is_elbow.Text);
                            dt1 = Functions.build_dt_from_excel(dt1, W1, startH1, endH1, lista_col, lista_colxl);
                        }
                    }

                    if (filename2.Length > 0 && sheet_name2.Length > 0)
                    {
                        W2 = Functions.Get_opened_worksheet_from_Excel_by_name(filename2, sheet_name2);
                        if (W2 != null)
                        {
                            List<string> lista_col = new List<string>();
                            List<string> lista_colxl = new List<string>();
                            lista_col.Add(col_sta);
                            lista_col.Add(col_defl);
                            lista_col.Add(col_is_elbow);
                            lista_colxl.Add(textBoxV_sta.Text);
                            lista_colxl.Add(textBoxV_angle.Text);
                            lista_colxl.Add(textBoxV_is_elbow.Text);
                            dt2 = Functions.build_dt_from_excel(dt2, W2, startV1, endV1, lista_col, lista_colxl);
                        }
                    }

                    string is_ovlp = "Is_overlap?";
                    string ovlp = "Overlap Value\r\n(6m or less IGNORE)";
                    string deltaPI = "DELTA PI";
                    string combined_defl = "Combined Bend Value";

                    if (dt1.Rows.Count > 0 || dt2.Rows.Count > 0)
                    {
                        System.Data.DataTable dt3 = new System.Data.DataTable();
                        dt3.Columns.Add(col_elbow_notes, typeof(string));
                        dt3.Columns.Add(col_bend_type, typeof(string));
                        dt3.Columns.Add(col_sta, typeof(double));
                        dt3.Columns.Add(col_sta1, typeof(double));
                        dt3.Columns.Add(col_sta2, typeof(double));
                        dt3.Columns.Add(col_defl, typeof(double));
                        dt3.Columns.Add(col_TL, typeof(double));
                        dt3.Columns.Add(col_radius, typeof(double));

                        dt3.Columns.Add(col_pup, typeof(double));
                        dt3.Columns.Add(is_ovlp, typeof(int));
                        dt3.Columns.Add(ovlp, typeof(double));
                        dt3.Columns.Add(deltaPI, typeof(double));
                        dt3.Columns.Add(combined_defl, typeof(double));

                        for (int i = 0; i < dt1.Rows.Count; ++i)
                        {
                            if (dt1.Rows[i][col_defl] != DBNull.Value)
                            {
                                double defl1 = Convert.ToDouble(dt1.Rows[i][col_defl]);
                                double X1 = radius1 * Math.Tan((defl1 * Math.PI / 180) / 2);
                                double X2 = radius2 * Math.Tan((defl1 * Math.PI / 180) / 2);

                                if (defl1 >= min_val && dt1.Rows[i][col_sta] != DBNull.Value)
                                {
                                    dt3.Rows.Add();
                                    dt3.Rows[dt3.Rows.Count - 1][col_bend_type] = "Horizontal";


                                    dt3.Rows[dt3.Rows.Count - 1][col_defl] = defl1;

                                    double sta = Convert.ToDouble(dt1.Rows[i][col_sta]);
                                    dt3.Rows[dt3.Rows.Count - 1][col_sta] = sta;

                                    bool is_elbow = false;
                                    if (dt1.Rows[i][col_is_elbow] != DBNull.Value)
                                    {
                                        if (Convert.ToBoolean(dt1.Rows[i][col_is_elbow]) == true)
                                        {
                                            is_elbow = true;
                                        }
                                    }


                                    double L = 0;
                                    double pup = 1;
                                    double radius = 0;







                                    if (is_elbow == false)
                                    {
                                        pup = 1.8;

                                        L = 2 * (pup + X1);
                                        dt3.Rows[dt3.Rows.Count - 1][col_pup] = pup;
                                        radius = radius1;

                                    }
                                    else
                                    {

                                        L = 2 * (pup + X2);
                                        dt3.Rows[dt3.Rows.Count - 1][col_elbow_notes] = "Elbow";
                                        dt3.Rows[dt3.Rows.Count - 1][col_pup] = pup;
                                        radius = radius2;
                                    }
                                    dt3.Rows[dt3.Rows.Count - 1][col_sta1] = sta - L / 2;
                                    dt3.Rows[dt3.Rows.Count - 1][col_sta2] = sta + L / 2;

                                    dt3.Rows[dt3.Rows.Count - 1][col_radius] = radius;
                                    dt3.Rows[dt3.Rows.Count - 1][col_TL] = L;

                                }

                            }

                        }

                        for (int i = 0; i < dt2.Rows.Count; ++i)
                        {
                            if (dt2.Rows[i][col_defl] != DBNull.Value)
                            {
                                double defl1 = Convert.ToDouble(dt2.Rows[i][col_defl]);
                                double X1 = radius1 * Math.Tan((defl1 * Math.PI / 180) / 2);
                                double X2 = radius2 * Math.Tan((defl1 * Math.PI / 180) / 2);

                                if (defl1 >= min_val && dt2.Rows[i][col_sta] != DBNull.Value)
                                {
                                    dt3.Rows.Add();
                                    dt3.Rows[dt3.Rows.Count - 1][col_bend_type] = "Vertical";

                                    dt3.Rows[dt3.Rows.Count - 1][col_defl] = defl1;

                                    double sta = Convert.ToDouble(dt2.Rows[i][col_sta]);
                                    dt3.Rows[dt3.Rows.Count - 1][col_sta] = sta;

                                    bool is_elbow = false;
                                    if (dt2.Rows[i][col_is_elbow] != DBNull.Value)
                                    {
                                        if (Convert.ToBoolean(dt2.Rows[i][col_is_elbow]) == true)
                                        {
                                            is_elbow = true;
                                        }
                                    }


                                    double L = 0;
                                    double pup = 1;
                                    double radius = 0;

                                    if (is_elbow == false)
                                    {
                                        pup = 1.8;

                                        L = 2 * (pup + X1);
                                        dt3.Rows[dt3.Rows.Count - 1][col_pup] = pup;
                                        radius = radius1;

                                    }
                                    else
                                    {

                                        L = 2 * (pup + X2);
                                        dt3.Rows[dt3.Rows.Count - 1][col_elbow_notes] = "Elbow";
                                        dt3.Rows[dt3.Rows.Count - 1][col_pup] = pup;
                                        radius = radius2;

                                    }
                                    dt3.Rows[dt3.Rows.Count - 1][col_sta1] = sta - L / 2;
                                    dt3.Rows[dt3.Rows.Count - 1][col_sta2] = sta + L / 2;

                                    dt3.Rows[dt3.Rows.Count - 1][col_radius] = radius;
                                    dt3.Rows[dt3.Rows.Count - 1][col_TL] = L;

                                }
                            }
                        }

                        dt3 = Functions.Sort_data_table(dt3, col_sta);


                        for (int i = 1; i < dt3.Rows.Count; ++i)
                        {
                            double sta2 = Convert.ToDouble(dt3.Rows[i - 1][col_sta2]);
                            double sta1 = Convert.ToDouble(dt3.Rows[i][col_sta1]);

                            double pi2 = Convert.ToDouble(dt3.Rows[i - 1][col_sta]);
                            double pi1 = Convert.ToDouble(dt3.Rows[i][col_sta]);

                            double defl2 = Convert.ToDouble(dt3.Rows[i - 1][col_defl]);
                            double defl1 = Convert.ToDouble(dt3.Rows[i][col_defl]);

                            if (sta2 > sta1)
                            {

                                if (Convert.ToString(dt3.Rows[i][col_bend_type]) == "Vertical" && Convert.ToString(dt3.Rows[i - 1][col_bend_type]) == "Horizontal")
                                {
                                    double t = defl1;
                                    defl1 = defl2;
                                    defl2 = t;
                                }

                                if (Convert.ToString(dt3.Rows[i][col_bend_type]) == Convert.ToString(dt3.Rows[i - 1][col_bend_type]))
                                {
                                    dt3.Rows[i][col_elbow_notes] = "Issue on " + Convert.ToString(dt3.Rows[i - 1][col_bend_type]);
                                    dt3.Rows[i - 1][col_elbow_notes] = "Issue on " + Convert.ToString(dt3.Rows[i - 1][col_bend_type]);
                                }


                                dt3.Rows[i][is_ovlp] = 1;
                                dt3.Rows[i][ovlp] = sta2 - sta1;
                                dt3.Rows[i][deltaPI] = pi1 - pi2;

                                if (Convert.ToString(dt3.Rows[i][col_bend_type]) == Convert.ToString(dt3.Rows[i - 1][col_bend_type]))
                                {
                                    dt3.Rows[i][col_elbow_notes] = "Issue on " + Convert.ToString(dt3.Rows[i - 1][col_bend_type]);
                                    dt3.Rows[i - 1][col_elbow_notes] = "Issue on " + Convert.ToString(dt3.Rows[i - 1][col_bend_type]);

                                    dt3.Rows[i][combined_defl] = 666;
                                }
                                else
                                {
                                    double combined = (Math.Acos((Math.Cos(defl1 * Math.PI / 180) * Math.Cos(0 * Math.PI / 180) * Math.Cos(defl2 * Math.PI / 180) + (Math.Sin(0 * Math.PI / 180) * Math.Sin(defl2 * Math.PI / 180))))) * 180 / Math.PI;
                                    dt3.Rows[i][combined_defl] = combined;
                                }


                            }
                        }

                        List<string> lista_col = new List<string>();
                        List<double> lista_w = new List<double>();
                        lista_col.Add("A");
                        lista_w.Add(6);
                        lista_col.Add("B");
                        lista_w.Add(10);
                        lista_col.Add("C");
                        lista_w.Add(10);
                        lista_col.Add("D");
                        lista_w.Add(10);
                        lista_col.Add("E");
                        lista_w.Add(10);
                        lista_col.Add("F");
                        lista_w.Add(10);
                        lista_col.Add("G");
                        lista_w.Add(10);
                        lista_col.Add("H");
                        lista_w.Add(10);

                        lista_col.Add("I");
                        lista_w.Add(10);

                        lista_col.Add("J");
                        lista_w.Add(10);

                        lista_col.Add("K");
                        lista_w.Add(10);

                        lista_col.Add("L");
                        lista_w.Add(10);

                        lista_col.Add("M");
                        lista_w.Add(10);

                        lista_col.Add("N");
                        lista_w.Add(20);


                        string nume1 = "Report - " + DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + "_" + DateTime.Now.Hour + "h" + DateTime.Now.Minute + "m";
                        Functions.Transfer_datatable_to_new_excel_spreadsheet(dt3, nume1, lista_col, lista_w);

                    }

                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }


    }
}

