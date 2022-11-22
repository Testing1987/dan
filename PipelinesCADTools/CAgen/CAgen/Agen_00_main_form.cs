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
    public partial class _AGEN_mainform : Form
    {
        private bool clickdragdown;
        private Point lastLocation;

        public static bool Exista_viewport_main = false;
        public static bool Exista_viewport_prof = false;
        public static bool Exista_viewport_owner = false;
        public static bool Exista_viewport_cross = false;
        public static bool Exista_viewport_mat = false;
        public static bool Exista_viewport_prof_band = false;
        public static bool Exista_viewport_tblk = false;

        public static string Project_type = "2D";

        public static AGEN_Home tpage_blank = null;
        public static AGEN_Viewport_Settings tpage_viewport_settings = null;
        public static AGEN_SheetIndex tpage_sheetindex = null;
        public static AGEN_CrossingScan tpage_crossing_scan = null;
        public static AGEN_OwnershipDraw tpage_owner_draw = null;

        public static AGEN_MaterialBand tpage_mat = null;
        public static AGEN_MaterialCount tpage_mat_count = null;

        public static AGEN_Owner_Band_Scan tpage_owner_scan = null;

        public static AGEN_Layer_alias tpage_layer_alias = null;
        public static AGEN_Sheet_Generation tpage_sheet_gen = null;
        public static AGEN_Project_Setup tpage_setup = null;
        public static AGEN_processing tpage_processing = null;
        public static AGEN_TBLK_Attributes tpage_tblk_attrib = null;
        public static AGEN_custom_band_scan tpage_cust_scan = null;
        public static AGEN_custom_band_draw tpage_cust_draw = null;


        public static AGEN_Profile_draw tpage_profdraw = null;
        public static AGEN_ProfileScan tpage_profilescan = null;
        public static AGEN_Crossing_draw tpage_crossing_draw = null;
        public static AGEN_band_analyze tpage_band_analize = null;
        public static AGEN_tools tpage_tools = null;
        public static Toolz_form tpage_toolz = null;
        public static VP2poly_form tpage_vp2poly = null;
        public static image_form tpage_image = null;

        public static AGEN_station_equations tpage_st_eq = null;
        public static Agen_load_cl_from_xl tpage_cl_xl = null;


        public static int round1 = 0;
        public static string units_of_measurement = "f";
        public static double Vw_scale = 1;

        public static double Vw_height = 0;
        public static double Vw_width = 0;
        public static double Vw_ps_x = 0;
        public static double Vw_ps_y = 0;
        public static bool Left_to_Right = true;


        public static double Vw_ps_tblk_x = 0;
        public static double Vw_ps_tblk_y = 0;

        public static double Vw_ps_prof_x = 0;
        public static double Vw_ps_prof_y = 0;

        public static double Vw_ps_mat_x = 0;
        public static double Vw_ps_mat_y = 0;

        public static double Vw_ps_cross_x = 0;
        public static double Vw_ps_cross_y = 0;

        public static double Vw_ps_prop_x = 0;
        public static double Vw_ps_prop_y = 0;

        public static double Vw_ps_profband_x = 0;
        public static double Vw_ps_profband_y = 0;

        public static double Vw_ps_slope_x = 0;
        public static double Vw_ps_slope_y = 0;

        public static double Vw_prof_height = 0;
        public static double Vw_profband_height = 0;



        public static double Vw_slope_height = 0;
        public static double Vw_slope_width = 0;

        public static bool ExcelVisible = false;
        public static double Match_distance = 5280;
        public static string Layer_name_ML_rectangle = "AGEN_Index_ML";
        public static string Layer_name_VP_rectangle = "AGEN_Index_VP";
        public static string Layer_North_Arrow = "NORTH";
        public static string Layer_Matchline_PaperSpace = "Agen_Matchline_PS";

        public static string layer_no_plot = "NO PLOT";
        public static string layer_crossing_band_text = "Agen_STA_Band_Text";
        public static string layer_crossing_band_pi = "Agen_STA_Band_PI";
        public static string layer_crossing_band_matchline = "Agen_matchline";
        public static string layer_prof_grid = "Agen_prof_Grid";
        public static string layer_prof_text = "Agen_prof_Text";
        public static string layer_prof_ground = "Agen_prof_grade";
        public static string layer_prof_pipe = "Agen_prof_pipe";
        public static string layer_prof_smys = "Agen_prof_smys";

        public static string layer_stationing = "Agen_sta";
        public static string layer_stationing_original = "Agen_sta";
        public static string layer_eq_blocks = "Agen_eq_blocks";
        public static string layer_pi_blocks = "Agen_pi_blocks";
        public static string layer_mp_blocks = "Agen_mp_blocks";
        public static string layer_prof_block_labels = "Agen_profile_block_labels";
        public static string layer_ownership_band = "Agen_band_ownership";
        public static string layer_ownership_band_no_plot = "Agen_no_plot_prop";
        public static string layer_centerline = "P_PL_CL";
        public static string layer_centerline_original = "P_PL_CL";
      
        public static short color_index_cl = 1;
        public static LineWeight lw_cl = LineWeight.ByLineWeightDefault;

        public static double texth = 5;



        public static string COUNTRY = "USA";

        public static Polyline Poly2D;
        public static Polyline3d Poly3D;

        public static string Layer_name_Main_Viewport = "AGEN_mainVP";
        public static string Layer_name_prof_main_viewport = "AGEN_VP_Prof_ON";
        public static string Layer_name_prof_side_viewport = "AGEN_VP_Prof_OFF";
        public static string Layer_name_ownership_Viewport = "AGEN_ownerVP";
        public static string Layer_name_crossing_Viewport = "AGEN_crossingVP";
        public static string Layer_name_material_Viewport = "AGEN_materialVP";
        public static string Layer_name_profband_Viewport = "VP_Prof_Band";
        public static string Layer_name_no_data_band_Viewport = "VP_ND_Band";
        public static string Layer_name_tblk_Viewport = "AGEN_tblkVP";

        public static string Layer_name_extra1_Viewport = "AGEN_VP1";
        public static string Layer_name_extra2_Viewport = "AGEN_VP2";
        public static string Layer_name_extra3_Viewport = "AGEN_VP3";
        public static string Layer_name_extra4_Viewport = "AGEN_VP4";
        public static string Layer_name_extra5_Viewport = "AGEN_VP5";


        public static string NA_name = "";
        public static string NorthArrowMS = "NorthArrow";
        public static string Layer_even = "BLKS_Even";
        public static string Layer_odd = "BLKS_Odd";
        public static string matchline_block = "AGEN_Matchline";
        public static string insertNAtoMS = "Insert into Sheet Index basefile";


        public static double NA_x = 0;
        public static double NA_y = 0;
        public static double NA_scale = 0;
        public static string Matchline_BlockName_in_PaperSpace = "";

        public static bool Freeze_operations = false;
        public static bool Template_is_open = false;

        public static string template1 = "";
        public static string template2 = "";


        public static System.Data.DataTable dt_sheet_index;
        public static System.Data.DataTable Data_table_Main_VP;
        public static System.Data.DataTable dt_centerline;

        public static System.Data.DataTable Data_table_blocks;

        public static System.Data.DataTable dt_station_equation;
        public static System.Data.DataTable dt_prof;
        public static System.Data.DataTable Data_Table_profile_band;


        public static System.Data.DataTable Data_Table_regular_bands;
        public static System.Data.DataTable Data_Table_custom_bands;
        public static double custom_band_scale = 1;
        public static System.Data.DataTable Data_Table_extra_mainVP;

        public static System.Data.DataTable Data_Table_display_bands;


        public static System.Data.DataTable Data_Table_property;
        public static System.Data.DataTable Data_Table_crossings;

        public static System.Data.DataTable dt_layer_alias;
        public static System.Data.DataTable Data_table_dwg_for_attributes;

        public static System.Data.DataTable dt_mat_lin;
        public static System.Data.DataTable dt_mat_lin_extra;
        public static System.Data.DataTable dt_mat2;

        public static System.Data.DataTable dt_mat_pt;
        public static System.Data.DataTable dt_settings_custom;
        public static System.Data.DataTable dt_ref;
        public static string ProjFolder = "";




        public static int Start_row_CL = 9;
        public static int Start_row_Sheet_index = 11;
        public static int Start_row_profile_band = 11;
        public static int Start_row_station_equation = 8;
        public static int Start_row_graph_profile = 8;
        public static int Start_row_1 = 1;
        public static int Start_row_property = 8;
        public static int Start_row_crossing = 8;
        public static int Start_row_layer_alias = 8;
        public static int Start_row_mat_lin = 13;
        public static int Start_row_mat_point = 12;
        public static int Start_row_block_attributes = 7;
        public static int Start_row_custom = 8;

        public static string Col_x = "X";
        public static string Col_y = "Y";
        public static string Col_z = "Z";
        public static string Col_station = "Station";
        public static string Col_descr = "Description";

        public static string Col_handle = "AcadHandle";
        public static string Col_dwg_name = "DwgNo";
        public static string Col_M1 = "StaBeg";
        public static string Col_M2 = "StaEnd";
        public static string Col_length = "Length";
        public static string Col_rot = "Rotation";
        public static string Col_Width = "Width";
        public static string Col_Height = "Height";

        public static string Col_DeflAng = "DeflAng";
        public static string Col_DeflAngDMS = "DeflAngDMS";
        public static string Col_Bearing = "Bearing";
        public static string Col_Distance = "Distance";
        public static string Col_2DSta = "2DSta";
        public static string Col_3DSta = "3DSta";

        public static string Col_MMid = "MMID";
        public static string Col_Type = "Type";
        public static string Col_Elev = "Elev";
        public static string Col_Elev1 = "Elev1";
        public static string Col_Elev2 = "Elev2";
        public static string Col_Sta_ahead = "Station Ahead";
        public static string Col_Sta_back = "Station Back";

        public static string Col_station_eq = "StationEq";
        public static string Col_Layer_name = "AcadLayer";


        public static string Col_offset = "Offset";
        public static string Col_block_name = "BlockName";
        public static string Col_left_right = "Side";

        public static string col_Full_name_dwg = "Drawing";
        public static string col_vpid1 = "VPId1";


        public static string col_desc = "Desc";
        public static string crossing_type_pi = "PI";

        public static string cl_excel_name = "centerline.xlsx";
        public static string sheet_index_excel_name = "sheet_index.xlsx";
        public static string prof_excel_name = "profile.xlsx";
        public static string imagery_excel_name = "imagery.xlsx";
        public static string band_prof_excel_name = "profile_band.xlsx";
        public static string property_excel_name = "property.xlsx";
        public static string crossing_excel_name = "crossing.xlsx";
        public static string layer_alias_excel_name = "layer alias.xlsx";
        public static string mat_linear_excel_name = "Material_Linear.xlsx";
        public static string mat_linear_extra_excel_name = "Material_Linear_extra.xlsx";
        public static string materials_excel_name = "materials.xlsx";
        public static string mat_points_excel_name = "Material_Points.xlsx";
        public static string block_attributes_excel_name = "TBLK_attributes.xlsx";
        public static string od2block_excel_name = "od2block.xlsx";
        public static string prof_labels_excel_name = "below_grade_profile_labels.xlsx";

        public static double prof_x0 = -1.123;
        public static double prof_y0 = -1.123;
        public static double prof_x_left = -1.123;
        public static double prof_x_right = -1.123;
        public static double prof_y_down = -1.123;
        public static double prof_width_lr = -1;
        public static double prof_texth = -1;
        public static double prof_hexag = 0;
        public static double prof_vexag = 0;
        public static double prof_down_el = 0;
        public static double prof_up_el = 0;
        public static double prof_start_sta = 0;
        public static double prof_end_sta = 0;



        public static string Col_2DSta1 = "2DStaBeg";
        public static string Col_3DSta1 = "3DStaBeg";
        public static string Col_2DSta2 = "2DStaEnd";
        public static string Col_3DSta2 = "3DStaEnd";
        public static string Col_EqSta1 = "EqStaBeg";
        public static string Col_EqSta2 = "EqStaEnd";
        public static string Col_Owner = "Owner";
        public static string Col_Linelist = "ParcelId";

        public static string Col_Material = "ItemNo";

        public static string Col_eqsta = "EqSta";
        public static string Col_DisplaySta = "DisplaySta";

        public static Point3d Point0_prop = new Point3d();
        public static Point3d Point0_tblk = new Point3d();
        public static Point3d Point0_cross = new Point3d();
        public static Point3d Point0_mat = new Point3d();
        public static Point3d Point0_slope = new Point3d();

        public static double Band_Separation = 1000;

        public static double Vw_cross_height = 0;
        public static double Vw_cross_width = 0;

        public static double Vw_mat_height = 0;
        public static double Vw_mat_width = 0;

        public static double Vw_prop_height = 0;
        public static double Vw_prop_width = 0;

        public static double Vw_tblk_height = 0;
        public static double Vw_tblk_width = 2;

        public static double tblk_separation = 150;
        public static double tblk_twist = 150;

  


        public static string version = "";


        public static string config_path = "";

        public static string layer_crossing = "";
        public static string current_segment = "";
        public static string first_custom_band = "";

        #region ownership attributes
        public static string owner_sta1_atr = "";
        public static string owner_sta2_atr = "";
        public static string owner_len_atr = "";
        public static string owner_linelist_atr = "";
        public static string owner_owner_atr = "";
        #endregion

        #region band analyse
        public static System.Data.DataTable dt_config_ownership = null;
        public static System.Data.DataTable dt_config_crossing = null;
        #endregion

        #region crossing draw parameters
        public static double XingDeltay1 = -123.456;
        public static double XingDeltay2 = -123.456;
        public static double XingDeltay3 = -123.456;
        #endregion

        #region nume bands
        public static string nume_banda_prof = "";
        public static string nume_banda_prop = "";
        public static string nume_banda_cross = "";
        public static string nume_banda_mat = "";
        public static string nume_main_vp = "";
        public static string nume_banda_profband = "";
        public static string nume_banda_tblk_band = "";
        public static string nume_banda_slope_band = "Slope Band";
        public static string nume_banda_no_data = "No Data Band";
        #endregion

        public static List<string> lista_segments = null;

        public static System.Data.DataTable dt_vp = null;

        public static List<int> lista_gen_prof_band = null;

        public _AGEN_mainform()
        {
            InitializeComponent();

            tpage_image = new image_form();
            tpage_image.MdiParent = this;
            tpage_image.Dock = DockStyle.Fill;
            tpage_image.Hide();


            tpage_crossing_scan = new AGEN_CrossingScan();
            tpage_crossing_scan.MdiParent = this;
            tpage_crossing_scan.Dock = DockStyle.Fill;
            tpage_crossing_scan.Hide();

            tpage_toolz = new Toolz_form();
            tpage_toolz.MdiParent = this;
            tpage_toolz.Dock = DockStyle.Fill;
            tpage_crossing_scan.Hide();

            tpage_vp2poly = new VP2poly_form();
            tpage_vp2poly.MdiParent = this;
            tpage_vp2poly.Dock = DockStyle.Fill;
            tpage_vp2poly.Hide();

            tpage_owner_draw = new AGEN_OwnershipDraw();
            tpage_owner_draw.MdiParent = this;
            tpage_owner_draw.Dock = DockStyle.Fill;
            tpage_owner_draw.Hide();

            tpage_profilescan = new AGEN_ProfileScan();
            tpage_profilescan.MdiParent = this;
            tpage_profilescan.Dock = DockStyle.Fill;
            tpage_profilescan.Hide();


            tpage_mat = new AGEN_MaterialBand();
            tpage_mat.MdiParent = this;
            tpage_mat.Dock = DockStyle.Fill;
            tpage_mat.Hide();

            tpage_mat_count = new AGEN_MaterialCount();
            tpage_mat_count.MdiParent = this;
            tpage_mat_count.Dock = DockStyle.Fill;
            tpage_mat_count.Hide();

            tpage_setup = new AGEN_Project_Setup();
            tpage_setup.MdiParent = this;
            tpage_setup.Dock = DockStyle.Fill;

            tpage_setup.Show();


            tpage_blank = new AGEN_Home();
            tpage_blank.MdiParent = this;
            tpage_blank.Dock = DockStyle.Fill;

            if (Functions.is_dan_popescu() == true)
            {
                tpage_blank.Hide();
            }
            else
            {
                tpage_blank.Show();
            }


            tpage_sheetindex = new AGEN_SheetIndex();
            tpage_sheetindex.MdiParent = this;
            tpage_sheetindex.Dock = DockStyle.Fill;
            tpage_sheetindex.Hide();


            tpage_sheet_gen = new AGEN_Sheet_Generation();
            tpage_sheet_gen.MdiParent = this;
            tpage_sheet_gen.Dock = DockStyle.Fill;
            tpage_sheet_gen.Hide();

            tpage_profdraw = new AGEN_Profile_draw();
            tpage_profdraw.MdiParent = this;
            tpage_profdraw.Dock = DockStyle.Fill;
            tpage_profdraw.Hide();

            tpage_layer_alias = new AGEN_Layer_alias();
            tpage_layer_alias.MdiParent = this;
            tpage_layer_alias.Dock = DockStyle.Fill;
            tpage_layer_alias.Hide();

            tpage_owner_scan = new AGEN_Owner_Band_Scan();
            tpage_owner_scan.MdiParent = this;
            tpage_owner_scan.Dock = DockStyle.Fill;
            tpage_owner_scan.Hide();

            tpage_viewport_settings = new AGEN_Viewport_Settings();
            tpage_viewport_settings.MdiParent = this;
            tpage_viewport_settings.Dock = DockStyle.Fill;
            tpage_viewport_settings.Hide();

            tpage_processing = new AGEN_processing();
            tpage_processing.MdiParent = this;
            tpage_processing.Dock = DockStyle.Fill;
            tpage_processing.Hide();

            tpage_tblk_attrib = new AGEN_TBLK_Attributes();
            tpage_tblk_attrib.MdiParent = this;
            tpage_tblk_attrib.Dock = DockStyle.Fill;
            tpage_tblk_attrib.Hide();

            tpage_cust_scan = new AGEN_custom_band_scan();
            tpage_cust_scan.MdiParent = this;
            tpage_cust_scan.Dock = DockStyle.Fill;
            tpage_cust_scan.Hide();

            tpage_cust_draw = new AGEN_custom_band_draw();
            tpage_cust_draw.MdiParent = this;
            tpage_cust_draw.Dock = DockStyle.Fill;
            tpage_cust_draw.Hide();


            tpage_crossing_draw = new AGEN_Crossing_draw();
            tpage_crossing_draw.MdiParent = this;
            tpage_crossing_draw.Dock = DockStyle.Fill;
            tpage_crossing_draw.Hide();

            tpage_band_analize = new AGEN_band_analyze();
            tpage_band_analize.MdiParent = this;
            tpage_band_analize.Dock = DockStyle.Fill;
            tpage_band_analize.Hide();

            tpage_tools = new AGEN_tools();
            tpage_tools.MdiParent = this;
            tpage_tools.Dock = DockStyle.Fill;
            tpage_tools.Hide();

            tpage_st_eq = new AGEN_station_equations();
            tpage_st_eq.MdiParent = this;
            tpage_st_eq.Dock = DockStyle.Fill;
            tpage_st_eq.Hide();

            tpage_cl_xl = new Agen_load_cl_from_xl();
            tpage_cl_xl.MdiParent = this;
            tpage_cl_xl.Dock = DockStyle.Fill;
            tpage_cl_xl.Hide();

            //sets the mdi background color at runtime
            foreach (Control ctrl in this.Controls)
            {
                if (ctrl is MdiClient)
                {
                    ctrl.BackColor = Color.FromArgb(62, 62, 66);
                }
            }

            tpage_profdraw.set_checkBox_prof_use_default_grid_val(true);
            tpage_profdraw.set_comboBox_prof_el_lbl_loc(0);


            treeView1.ShowPlusMinus = false;

            tpage_sheet_gen.Hide_panel_custom_bands();
            tpage_sheet_gen.Hide_panel_extra_bands();

            Data_Table_regular_bands = Functions.creeaza_regular_band_data_table_structure();
            Data_Table_custom_bands = Functions.creeaza_custom_band_data_table_structure();


            if (Functions.is_dan_popescu() == true)
            {
                ExcelVisible = true;


            }
            else
            {
                ExcelVisible = false;

                //tpage_profilescan.Hide_checkBox_draft_profile_in_ps2();
                // treeView1.Nodes[0].Nodes[2].Remove();
            }
        }

        protected override void OnLoad(EventArgs e)
        {
            // Hides the ugly border around the mdi container (main form)
            var mdiclient = this.Controls.OfType<MdiClient>().Single();
            this.SuspendLayout();
            mdiclient.SuspendLayout();
            var hdiff = mdiclient.Size.Width - mdiclient.ClientSize.Width;
            var vdiff = mdiclient.Size.Height - mdiclient.ClientSize.Height;
            var size = new Size(mdiclient.Width + hdiff, mdiclient.Height + vdiff);
            var location = new Point(mdiclient.Left - (hdiff / 2), mdiclient.Top - (vdiff / 2));
            mdiclient.Dock = DockStyle.None;
            mdiclient.Size = size;
            mdiclient.Location = location;
            mdiclient.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom;
            mdiclient.ResumeLayout(true);
            this.ResumeLayout(true);
            base.OnLoad(e);
        }


        private void clickmove_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            clickdragdown = true;
            lastLocation = e.Location;
        }

        private void clickmove_MouseMove(object sender, MouseEventArgs e)
        {
            if (clickdragdown)
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
            try
            {

                lista_segments = null;
                current_segment = "";
                COUNTRY = "USA";
                units_of_measurement = "f";
                Poly2D = null;
                Poly3D = null;
                _AGEN_mainform.ProjFolder = "";
                _AGEN_mainform.dt_centerline = null;
                _AGEN_mainform.dt_sheet_index = null;
                _AGEN_mainform.dt_station_equation = null;
                _AGEN_mainform.dt_prof = null;
                template1 = "";
                template2 = "";

                int i = 0;

                do
                {
                    System.Windows.Forms.Form Forma1 = System.Windows.Forms.Application.OpenForms[i];
                    if (Forma1 is AGEN_custom_band_form)
                    {
                        Forma1.Close();
                    }

                    i = i + 1;

                } while (i < System.Windows.Forms.Application.OpenForms.Count);

            }
            catch (InvalidOperationException ex)
            {

            }

            this.Close();


        }
        private void button_minimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }





        public void Set_textBox_config_file_location(string file1)
        {
            textBox_config_file_path.Text = file1;
            config_path = file1;
        }





        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            var hitTest = e.Node.TreeView.HitTest(e.Location);
            if (hitTest.Location == TreeViewHitTestLocations.PlusMinus)
            {
                return;
            }

            if (e.Node.IsExpanded)
            {
                e.Node.Collapse();
            }
            else
            {
                e.Node.Expand();
            }

            if (e.Node.Text == "Project" || e.Node.Text == "Settings")
            {
                tpage_processing.Hide();
                tpage_blank.Hide();
                tpage_setup.Show();
                tpage_viewport_settings.Hide();
                tpage_tblk_attrib.Hide();
                tpage_band_analize.Hide();
                tpage_sheetindex.Hide();
                tpage_layer_alias.Hide();
                tpage_crossing_scan.Hide();
                tpage_crossing_draw.Hide();
                tpage_profilescan.Hide();
                tpage_profdraw.Hide();
                tpage_owner_scan.Hide();
                tpage_owner_draw.Hide();
                tpage_mat.Hide();
                tpage_cust_scan.Hide();
                tpage_cust_draw.Hide();
                tpage_sheet_gen.Hide();
                tpage_tools.Hide();
                tpage_st_eq.Hide();
                tpage_cl_xl.Hide();
                tpage_toolz.Hide();
                tpage_vp2poly.Hide();
                tpage_image.Hide();
            }



            if (e.Node.Text == "Border Definition")
            {
                tpage_processing.Hide();
                tpage_blank.Hide();
                tpage_setup.Hide();
                tpage_viewport_settings.Show();
                tpage_tblk_attrib.Hide();
                tpage_band_analize.Hide();
                tpage_sheetindex.Hide();
                tpage_layer_alias.Hide();
                tpage_crossing_scan.Hide();
                tpage_crossing_draw.Hide();
                tpage_profilescan.Hide();
                tpage_profdraw.Hide();
                tpage_owner_scan.Hide();
                tpage_owner_draw.Hide();
                tpage_mat.Hide();
                tpage_cust_scan.Hide();
                tpage_cust_draw.Hide();
                tpage_sheet_gen.Hide();
                tpage_tools.Hide();
                tpage_st_eq.Hide();
                tpage_cl_xl.Hide();
                tpage_toolz.Hide();
                tpage_vp2poly.Hide();
                tpage_image.Hide();

            }

            if (e.Node.Text == "Block Attributes")
            {
                tpage_processing.Hide();
                tpage_blank.Hide();
                tpage_setup.Hide();
                tpage_viewport_settings.Hide();
                tpage_tblk_attrib.Show();
                tpage_band_analize.Hide();
                tpage_sheetindex.Hide();
                tpage_layer_alias.Hide();
                tpage_crossing_scan.Hide();
                tpage_crossing_draw.Hide();
                tpage_profilescan.Hide();
                tpage_profdraw.Hide();
                tpage_owner_scan.Hide();
                tpage_owner_draw.Hide();
                tpage_mat.Hide();
                tpage_cust_scan.Hide();
                tpage_cust_draw.Hide();
                tpage_sheet_gen.Hide();
                tpage_tools.Hide();
                tpage_st_eq.Hide();
                tpage_cl_xl.Hide();
                tpage_toolz.Hide();
                tpage_vp2poly.Hide();
                tpage_image.Hide();

            }

            if (e.Node.Text == "Band Analysis")
            {
                tpage_processing.Hide();
                tpage_blank.Hide();
                tpage_setup.Hide();
                tpage_viewport_settings.Hide();
                tpage_tblk_attrib.Hide();
                tpage_band_analize.Show();
                tpage_sheetindex.Hide();
                tpage_layer_alias.Hide();
                tpage_crossing_scan.Hide();
                tpage_crossing_draw.Hide();
                tpage_profilescan.Hide();
                tpage_profdraw.Hide();
                tpage_owner_scan.Hide();
                tpage_owner_draw.Hide();
                tpage_mat.Hide();
                tpage_cust_scan.Hide();
                tpage_cust_draw.Hide();
                tpage_sheet_gen.Hide();
                tpage_tools.Hide();
                tpage_st_eq.Hide();
                tpage_cl_xl.Hide();
                tpage_toolz.Hide();
                tpage_vp2poly.Hide();
                tpage_image.Hide();

            }

            if (e.Node.Text == "Plan View" || e.Node.Text == "Sheet Index Setup")
            {
                tpage_processing.Hide();
                tpage_blank.Hide();
                tpage_setup.Hide();
                tpage_viewport_settings.Hide();
                tpage_tblk_attrib.Hide();
                tpage_band_analize.Hide();
                tpage_sheetindex.Show();
                tpage_layer_alias.Hide();
                tpage_crossing_scan.Hide();
                tpage_crossing_draw.Hide();
                tpage_profilescan.Hide();
                tpage_profdraw.Hide();
                tpage_owner_scan.Hide();
                tpage_owner_draw.Hide();
                tpage_mat.Hide();
                tpage_cust_scan.Hide();
                tpage_cust_draw.Hide();
                tpage_sheet_gen.Hide();
                tpage_tools.Hide();
                tpage_st_eq.Hide();
                tpage_cl_xl.Hide();
                tpage_toolz.Hide();
                tpage_vp2poly.Hide();
                tpage_image.Hide();

            }

            if (e.Node.Text == "Band Builder")
            {
                tpage_processing.Hide();
                tpage_blank.Show();
                tpage_setup.Hide();
                tpage_viewport_settings.Hide();
                tpage_tblk_attrib.Hide();
                tpage_band_analize.Hide();
                tpage_sheetindex.Hide();
                tpage_layer_alias.Hide();
                tpage_crossing_scan.Hide();
                tpage_crossing_draw.Hide();
                tpage_profilescan.Hide();
                tpage_profdraw.Hide();
                tpage_owner_scan.Hide();
                tpage_owner_draw.Hide();
                tpage_mat.Hide();
                tpage_cust_scan.Hide();
                tpage_cust_draw.Hide();
                tpage_sheet_gen.Hide();
                tpage_tools.Hide();
                tpage_st_eq.Hide();
                tpage_cl_xl.Hide();
                tpage_toolz.Hide();
                tpage_vp2poly.Hide();
                tpage_image.Hide();

            }


            if (e.Node.Text == "Ownership")
            {
                tpage_processing.Hide();
                tpage_blank.Hide();
                tpage_setup.Hide();
                tpage_viewport_settings.Hide();
                tpage_tblk_attrib.Hide();
                tpage_band_analize.Hide();
                tpage_sheetindex.Hide();
                tpage_layer_alias.Hide();
                tpage_crossing_scan.Hide();
                tpage_crossing_draw.Hide();
                tpage_profilescan.Hide();
                tpage_profdraw.Hide();
                tpage_owner_scan.Show();
                tpage_owner_draw.Hide();
                tpage_mat.Hide();
                tpage_cust_scan.Hide();
                tpage_cust_draw.Hide();
                tpage_sheet_gen.Hide();
                tpage_tools.Hide();
                tpage_st_eq.Hide();
                tpage_cl_xl.Hide();
                tpage_toolz.Hide();
                tpage_vp2poly.Hide();
                tpage_image.Hide();

            }

            if (e.Node.Text == "Crossing")
            {
                tpage_processing.Hide();
                tpage_blank.Hide();
                tpage_setup.Hide();
                tpage_viewport_settings.Hide();
                tpage_tblk_attrib.Hide();
                tpage_band_analize.Hide();
                tpage_sheetindex.Hide();
                tpage_layer_alias.Show();
                tpage_crossing_scan.Hide();
                tpage_crossing_draw.Hide();
                tpage_profilescan.Hide();
                tpage_profdraw.Hide();
                tpage_owner_scan.Hide();
                tpage_owner_draw.Hide();
                tpage_mat.Hide();
                tpage_cust_scan.Hide();
                tpage_cust_draw.Hide();
                tpage_sheet_gen.Hide();
                tpage_tools.Hide();
                tpage_st_eq.Hide();
                tpage_cl_xl.Hide();
                tpage_toolz.Hide();
                tpage_vp2poly.Hide();
                tpage_image.Hide();

            }
            if (e.Node.Text == "Material")
            {
                tpage_processing.Hide();
                tpage_blank.Hide();
                tpage_setup.Hide();
                tpage_viewport_settings.Hide();
                tpage_tblk_attrib.Hide();
                tpage_band_analize.Hide();
                tpage_sheetindex.Hide();
                tpage_layer_alias.Hide();
                tpage_crossing_scan.Hide();
                tpage_crossing_draw.Hide();
                tpage_profilescan.Hide();
                tpage_profdraw.Hide();
                tpage_owner_scan.Hide();
                tpage_owner_draw.Hide();
                tpage_mat.Show();
                tpage_cust_scan.Hide();
                tpage_cust_draw.Hide();
                tpage_sheet_gen.Hide();
                tpage_tools.Hide();
                tpage_st_eq.Hide();
                tpage_cl_xl.Hide();
                tpage_toolz.Hide();
                tpage_vp2poly.Hide();
                tpage_image.Hide();

            }
            if (e.Node.Text == "Profile")
            {
                tpage_processing.Hide();
                tpage_blank.Hide();
                tpage_setup.Hide();
                tpage_viewport_settings.Hide();
                tpage_tblk_attrib.Hide();
                tpage_band_analize.Hide();
                tpage_sheetindex.Hide();
                tpage_layer_alias.Hide();
                tpage_crossing_scan.Hide();
                tpage_crossing_draw.Hide();
                tpage_profilescan.Show();
                tpage_profdraw.Hide();
                tpage_owner_scan.Hide();
                tpage_owner_draw.Hide();
                tpage_mat.Hide();
                tpage_cust_scan.Hide();
                tpage_cust_draw.Hide();
                tpage_sheet_gen.Hide();
                tpage_tools.Hide();
                tpage_st_eq.Hide();
                tpage_cl_xl.Hide();
                tpage_toolz.Hide();
                tpage_vp2poly.Hide();
                tpage_image.Hide();

            }

            if (e.Node.Text == "Sheet Generation" || e.Node.Text == "Create Alignment Sheets")
            {
                tpage_processing.Hide();
                tpage_blank.Hide();
                tpage_setup.Hide();
                tpage_viewport_settings.Hide();
                tpage_tblk_attrib.Hide();
                tpage_band_analize.Hide();
                tpage_sheetindex.Hide();
                tpage_layer_alias.Hide();
                tpage_crossing_scan.Hide();
                tpage_crossing_draw.Hide();
                tpage_profilescan.Hide();
                tpage_profdraw.Hide();
                tpage_owner_scan.Hide();
                tpage_owner_draw.Hide();
                tpage_mat.Hide();
                tpage_cust_scan.Hide();
                tpage_cust_draw.Hide();
                tpage_sheet_gen.Show();
                tpage_tools.Hide();
                tpage_st_eq.Hide();
                tpage_cl_xl.Hide();
                tpage_toolz.Hide();
                tpage_vp2poly.Hide();
                tpage_image.Hide();

            }

            if (e.Node.Text == "Custom")
            {
                tpage_processing.Hide();
                tpage_blank.Hide();
                tpage_setup.Hide();
                tpage_viewport_settings.Hide();
                tpage_tblk_attrib.Hide();
                tpage_band_analize.Hide();
                tpage_sheetindex.Hide();
                tpage_layer_alias.Hide();
                tpage_crossing_scan.Hide();
                tpage_crossing_draw.Hide();
                tpage_profilescan.Hide();
                tpage_profdraw.Hide();
                tpage_owner_scan.Hide();
                tpage_owner_draw.Hide();
                tpage_mat.Hide();
                tpage_cust_scan.Show();
                tpage_cust_draw.Hide();
                tpage_sheet_gen.Hide();
                tpage_tools.Hide();
                tpage_st_eq.Hide();
                tpage_cl_xl.Hide();
                tpage_toolz.Hide();
                tpage_vp2poly.Hide();
                tpage_image.Hide();

            }

            if (e.Node.Text == "Station Equations")
            {
                tpage_processing.Hide();
                tpage_blank.Hide();
                tpage_setup.Hide();
                tpage_viewport_settings.Hide();
                tpage_tblk_attrib.Hide();
                tpage_band_analize.Hide();
                tpage_sheetindex.Hide();
                tpage_layer_alias.Hide();
                tpage_crossing_scan.Hide();
                tpage_crossing_draw.Hide();
                tpage_profilescan.Hide();
                tpage_profdraw.Hide();
                tpage_owner_scan.Hide();
                tpage_owner_draw.Hide();
                tpage_mat.Hide();
                tpage_cust_scan.Hide();
                tpage_cust_draw.Hide();
                tpage_sheet_gen.Hide();
                tpage_tools.Hide();
                tpage_st_eq.Show();
                tpage_cl_xl.Hide();
                tpage_toolz.Hide();
                tpage_vp2poly.Hide();
                tpage_image.Hide();

            }

            if (e.Node.Text == "Extra tools" || e.Node.Text == "Rename Layout")
            {
                tpage_processing.Hide();
                tpage_blank.Hide();
                tpage_setup.Hide();
                tpage_viewport_settings.Hide();
                tpage_tblk_attrib.Hide();
                tpage_band_analize.Hide();
                tpage_sheetindex.Hide();
                tpage_layer_alias.Hide();
                tpage_crossing_scan.Hide();
                tpage_crossing_draw.Hide();
                tpage_profilescan.Hide();
                tpage_profdraw.Hide();
                tpage_owner_scan.Hide();
                tpage_owner_draw.Hide();
                tpage_mat.Hide();
                tpage_cust_scan.Hide();
                tpage_cust_draw.Hide();
                tpage_sheet_gen.Hide();
                tpage_tools.Hide();
                tpage_st_eq.Hide();
                tpage_cl_xl.Hide();
                tpage_toolz.Show();
                tpage_vp2poly.Hide();
                tpage_image.Hide();

            }
            if (e.Node.Name == "Node42")
            {
                tpage_processing.Hide();
                tpage_blank.Hide();
                tpage_setup.Hide();
                tpage_viewport_settings.Hide();
                tpage_tblk_attrib.Hide();
                tpage_band_analize.Hide();
                tpage_sheetindex.Hide();
                tpage_layer_alias.Hide();
                tpage_crossing_scan.Hide();
                tpage_crossing_draw.Hide();
                tpage_profilescan.Hide();
                tpage_profdraw.Hide();
                tpage_owner_scan.Hide();
                tpage_owner_draw.Hide();
                tpage_mat.Hide();
                tpage_cust_scan.Hide();
                tpage_cust_draw.Hide();
                tpage_sheet_gen.Hide();
                tpage_tools.Hide();
                tpage_st_eq.Hide();
                tpage_cl_xl.Hide();
                tpage_toolz.Hide();
                tpage_vp2poly.Show();
                tpage_image.Hide();

            }
            if (e.Node.Name == "Node43")
            {
                tpage_processing.Hide();
                tpage_blank.Hide();
                tpage_setup.Hide();
                tpage_viewport_settings.Hide();
                tpage_tblk_attrib.Hide();
                tpage_band_analize.Hide();
                tpage_sheetindex.Hide();
                tpage_layer_alias.Hide();
                tpage_crossing_scan.Hide();
                tpage_crossing_draw.Hide();
                tpage_profilescan.Hide();
                tpage_profdraw.Hide();
                tpage_owner_scan.Hide();
                tpage_owner_draw.Hide();
                tpage_mat.Hide();
                tpage_cust_scan.Hide();
                tpage_cust_draw.Hide();
                tpage_sheet_gen.Hide();
                tpage_tools.Hide();
                tpage_st_eq.Hide();
                tpage_cl_xl.Hide();
                tpage_toolz.Hide();
                tpage_vp2poly.Hide();
                tpage_image.Show();

            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process proc = new System.Diagnostics.Process();
            proc.StartInfo.FileName = "mailto:Support.CADTechnolgies@mottmacna.com?subject=Agen Help";
            proc.Start();
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            MessageBox.Show("Please contact Hector Morales / Richard Pangburn");
        }


    }
}
