using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace weld_sheet
{
    public partial class Form_main : Form
    {
        //Global Variables
        System.Data.DataTable dt_weld = null;

        //System.Data.DataTable DT_Missing_From_Spot = null;
        public Form_main FN_Form = null;



        string col_pnt = "PNT";
        string col_northing = "NORTHING";
        string col_easting = "EASTING";
        string col_elevation = "ELEVATION";
        string col_feature_code = "FEATURE_CODE";
        string col_description = "DESCRIPTION";
        string col_station = "STATION";
        string col_3d_stations = "3D_STATIONS";
        string col_mm_bk = "MM_BK";
        string col_wall_bk = "WALL_BK";
        string col_pipe_bk = "PIPE_BK";
        string col_heat_bk = "HEAT_BK";
        string col_x_ray = "X-RAY#";
        string col_mm_ahd = "MM_AHD";
        string col_wall_ahd = "WALL_AHD";
        string col_pipe_ahd = "PIPE_AHD";
        string col_heat_ahd = "HEAT_AHD";
        string col_length = "LENGTH";
        string col_ng = "NG";
        string col_ng_northing = "NG_NORTHING";
        string col_ng_easting = "NG_EASTING";
        string col_ng_elevation = "NG_ELEVATION";
        string col_cover = "COVER";
        string col_location = "LOCATION";
        string col_filename = "FILENAME";
        string col_spread = "SPREAD";
        string col_v_angle = "V_ANGLE";
        string col_h_angle = "H_ANGLE";
        string col_restored_ng = "RESTORED_NG";
        string col_r_ng_northing = "R_NG_NORTHING";
        string col_r_ng_easting = "R_NG_EASTING";
        string col_r_ng_elevation = "R_NG_ELEVATION";
        string col_restored_cover = "RESTORED COVER";

        string xl_col_pnt = "A";
        string xl_col_northing = "B";
        string xl_col_easting = "C";
        string xl_col_elevation = "D";
        string xl_col_feature_code = "E";
        string xl_col_description = "F";
        string xl_col_station = "G";
        string xl_col_3d_stations = "H";
        string xl_col_mm_bk = "I";
        string xl_col_wall_bk = "J";
        string xl_col_pipe_bk = "K";
        string xl_col_heat_bk = "L";
        string xl_col_x_ray = "M";
        string xl_col_mm_ahd = "N";
        string xl_col_wall_ahd = "O";
        string xl_col_pipe_ahd = "P";
        string xl_col_heat_ahd = "Q";
        string xl_col_length = "R";
        string xl_col_ng = "S";
        string xl_col_ng_northing = "T";
        string xl_col_ng_easting = "U";
        string xl_col_ng_elevation = "V";
        string xl_col_cover = "W";
        string xl_col_location = "X";
        string xl_col_filename = "Y";
        string xl_col_spread = "Z";
        string xl_col_v_angle = "AA";
        string xl_col_h_angle = "AB";
        string xl_col_restored_ng = "AC";
        string xl_col_r_ng_northing = "AD";
        string xl_col_r_ng_easting = "AE";
        string xl_col_r_ng_elevation = "AF";
        string xl_col_restored_cover = "AG";

        int start_row = 2;

        float extra1 = 0;
        float a = 0;

        string WGEN_folder = Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments) + "\\WGEN\\";


        Microsoft.Office.Interop.Excel.Application Excel1 = Functions.Excel1;
        private bool clickdragdown;
        private System.Drawing.Point lastLocation;

        public Form_main()
        {
            InitializeComponent();
            this.FormBorderStyle = FormBorderStyle.None;
            this.DoubleBuffered = true;
            this.SetStyle(ControlStyles.ResizeRedraw, true);
            FN_Form = this;
        }

        #region resize window
        private const int cGrip = 8;      // Grip size
        private const int cCaption = 32;   // Caption bar height;

        protected override void OnPaint(PaintEventArgs e)
        {
            System.Drawing.Rectangle rc = new System.Drawing.Rectangle(this.ClientSize.Width - cGrip, this.ClientSize.Height - cGrip, cGrip, cGrip);
            ControlPaint.DrawSizeGrip(e.Graphics, this.BackColor, rc);
            rc = new System.Drawing.Rectangle(0, 0, this.ClientSize.Width, cCaption);
            e.Graphics.FillRectangle(Brushes.Transparent, rc);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x84)
            {  // Trap WM_NCHITTEST
                System.Drawing.Point pos = new System.Drawing.Point(m.LParam.ToInt32());
                pos = this.PointToClient(pos);
                if (pos.Y < cCaption)
                {
                    m.Result = (IntPtr)2;  // HTCAPTION
                    return;
                }
                if (pos.X >= this.ClientSize.Width - cGrip && pos.Y >= this.ClientSize.Height - cGrip)
                {
                    m.Result = (IntPtr)17; // HTBOTTOMRIGHT
                    return;
                }
            }
            base.WndProc(ref m);
        }
        #endregion

        #region set enable true or false    
        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_Export_Field_Notes);
            lista_butoane.Add(button_minimize);
            lista_butoane.Add(comboBox_wm);


            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = false;
            }
        }
        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_Export_Field_Notes);
            lista_butoane.Add(button_minimize);
            lista_butoane.Add(comboBox_wm);


            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }
        #endregion

        #region Mouse Commands
        private void clickmove_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {

            clickdragdown = true;
            lastLocation = e.Location;
        }


        private void clickmove_MouseMove(object sender, MouseEventArgs e)
        {
            if (clickdragdown)
            {
                this.Location = new System.Drawing.Point(
                  (this.Location.X - lastLocation.X) + e.X, (this.Location.Y - lastLocation.Y) + e.Y);

                this.Update();
            }
        }

        private void clickmove_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            clickdragdown = false;
        }

        private void button_minimize_Click(object sender, EventArgs e)
        //Minimizes Form
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void button_Exit_Click(object sender, EventArgs e)
        //Closes Form
        {
            this.Close();
        }
        #endregion

        public void Page_Setup(Worksheet W1)
        {

            #region Margins
            W1.PageSetup.LeftMargin = .25 * 72;
            W1.PageSetup.RightMargin = .25 * 72;
            W1.PageSetup.TopMargin = .4 * 72;
            W1.PageSetup.BottomMargin = .25 * 72;
            W1.PageSetup.HeaderMargin = .1 * 72;
            W1.PageSetup.FooterMargin = .1 * 72;
            W1.PageSetup.Orientation = XlPageOrientation.xlPortrait;
            W1.PageSetup.PaperSize = XlPaperSize.xlPaperLetter;
            #endregion

            #region Column_Widths
            W1.Columns["A:A"].ColumnWidth = .5;
            W1.Columns["B:B"].ColumnWidth = .5;
            W1.Columns["C:C"].ColumnWidth = 7;
            W1.Columns["D:D"].ColumnWidth = .5;
            W1.Columns["E:F"].ColumnWidth = 4;
            W1.Columns["G:G"].ColumnWidth = .5;
            W1.Columns["H:H"].ColumnWidth = 8;
            W1.Columns["I:I"].ColumnWidth = .5;
            W1.Columns["J:J"].ColumnWidth = 5;
            W1.Columns["K:K"].ColumnWidth = 2;
            W1.Columns["L:L"].ColumnWidth = 5;
            W1.Columns["M:M"].ColumnWidth = .5;
            W1.Columns["N:N"].ColumnWidth = 2;
            W1.Columns["O:O"].ColumnWidth = .5;
            W1.Columns["P:P"].ColumnWidth = 6;
            W1.Columns["Q:Q"].ColumnWidth = 8;
            W1.Columns["R:R"].ColumnWidth = .5;
            W1.Columns["S:S"].ColumnWidth = 2;
            W1.Columns["T:T"].ColumnWidth = 2;
            W1.Columns["U:V"].ColumnWidth = 12;
            #endregion

            #region Header_Row_Heights
            //header
            W1.Rows[1].RowHeight = 14.4;
            W1.Rows[2].RowHeight = 4.1;
            W1.Rows[3].RowHeight = 14.4;
            W1.Rows[4].RowHeight = 4.1;
            W1.Rows[5].RowHeight = 14.4;
            W1.Rows[6].RowHeight = 4.1;
            W1.Rows[7].RowHeight = 3;
            W1.Rows[8].RowHeight = 4.1;
            W1.Rows[9].RowHeight = 14.4;
            W1.Rows[10].RowHeight = 14.4;
            W1.Rows[11].RowHeight = 14.4;
            W1.Rows[12].RowHeight = 4.1;
            #endregion

            W1.Range["A7:Q7"].Interior.Color = Color.Black;

            #region Add_Arrow

            Shape shp = W1.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeUpArrow, 230, 60, 30, 825);
            shp.ShapeStyle = Microsoft.Office.Core.MsoShapeStyleIndex.msoLineStylePreset8;
            shp.Fill.ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorText1;

            #endregion

            #region Add_Legend

            //Picture pic1 = (Picture);


            #endregion

            #region DataRow_Heights
            int Row_start = 13;
            int Increment = 9;


            int Row_Index = Row_start;

            W1.Rows[Row_Index].RowHeight = 4.1;
            W1.Rows[Row_Index + 1].RowHeight = 14.4;
            W1.Rows[Row_Index + 2].RowHeight = 12;
            W1.Rows[Row_Index + 3].RowHeight = 14.4;
            W1.Rows[Row_Index + 4].RowHeight = 12;
            W1.Rows[Row_Index + 5].RowHeight = 14.4;
            W1.Rows[Row_Index + 6].RowHeight = 12;
            W1.Rows[Row_Index + 7].RowHeight = 14.4;
            W1.Rows[Row_Index + 8].RowHeight = 4.1;
            #endregion

            #region Merge_Cells
            //header
            W1.Range[W1.Cells[1, 5], W1.Cells[1, 6]].merge(); // INTIALS Text Box
            W1.Range[W1.Cells[3, 5], W1.Cells[3, 6]].merge(); // CREW Text Box
            W1.Range[W1.Cells[5, 5], W1.Cells[5, 6]].merge(); // FROM Text Box
            W1.Range[W1.Cells[11, 5], W1.Cells[11, 6]].merge(); // Point Text Box
            //W1.Range[W1.Cells[1, 8], W1.Cells[1, 10]].merge(); // SHEET Label
            //W1.Range[W1.Cells[3, 8], W1.Cells[3, 10]].merge(); // FILENAME Label
            //W1.Range[W1.Cells[5, 8], W1.Cells[5, 10]].merge(); // TO Label
            //W1.Range[W1.Cells[1, 11], W1.Cells[1, 12]].merge(); // SHEET Textbox
            W1.Range[W1.Cells[3, 10], W1.Cells[3, 14]].merge(); // FILENAME Textbox
            W1.Range[W1.Cells[11, 10], W1.Cells[11, 12]].merge(); // Weld Textbox
            W1.Range[W1.Cells[5, 10], W1.Cells[5, 12]].merge(); // TO Textbox
            W1.Range[W1.Cells[1, 13], W1.Cells[1, 14]].merge(); // OF Label
            //W1.Range[W1.Cells[1, 15], W1.Cells[1, 17]].merge(); // OF Textbox
            W1.Range[W1.Cells[10, 16], W1.Cells[10, 17]].merge(); // STATION Label
            W1.Range[W1.Cells[11, 16], W1.Cells[11, 17]].merge(); // STATION Textbox

            W1.Range["U14:V16"].Merge();
            W1.Range["U18:V20"].Merge();
            W1.Range["U23:V25"].Merge();
            W1.Range["U27:V29"].Merge();
            W1.Range["U32:V34"].Merge();
            W1.Range["U36:V38"].Merge();
            W1.Range["U41:V43"].Merge();
            W1.Range["U45:V47"].Merge();
            W1.Range["U50:V52"].Merge();
            W1.Range["U54:V56"].Merge();
            W1.Range["U59:V61"].Merge();
            W1.Range["U63:V65"].Merge();

            //datarow
            W1.Range[W1.Cells[Row_Index + 1, 5], W1.Cells[Row_Index + 1, 6]].merge(); // NG Text Box
            W1.Range[W1.Cells[Row_Index + 3, 5], W1.Cells[Row_Index + 3, 6]].merge(); // Cut Off Text Box
            W1.Range[W1.Cells[Row_Index + 7, 5], W1.Cells[Row_Index + 7, 6]].merge(); // Point Text Box
            W1.Range[W1.Cells[Row_Index + 1, 10], W1.Cells[Row_Index + 1, 12]].merge(); // MMID Text Box
            W1.Range[W1.Cells[Row_Index + 3, 10], W1.Cells[Row_Index + 3, 12]].merge(); // Pipe Text Box
            W1.Range[W1.Cells[Row_Index + 5, 10], W1.Cells[Row_Index + 5, 12]].merge(); // Heat Text Box
            W1.Range[W1.Cells[Row_Index + 7, 10], W1.Cells[Row_Index + 7, 12]].merge(); // Weld Text Box
            W1.Range[W1.Cells[Row_Index + 1, 16], W1.Cells[Row_Index + 1, 17]].merge(); // W.T. Label
            W1.Range[W1.Cells[Row_Index + 2, 16], W1.Cells[Row_Index + 2, 17]].merge(); // W.T. Textbox
            W1.Range[W1.Cells[Row_Index + 3, 16], W1.Cells[Row_Index + 3, 17]].merge(); // Length Label
            W1.Range[W1.Cells[Row_Index + 4, 16], W1.Cells[Row_Index + 4, 17]].merge(); // Length Textbox
            W1.Range[W1.Cells[Row_Index + 5, 16], W1.Cells[Row_Index + 5, 17]].merge(); // Station Label
            W1.Range[W1.Cells[Row_Index + 6, 16], W1.Cells[Row_Index + 6, 17]].merge(); // Station Textbox
            #endregion

            #region Add_Labels
            //header
            //Label_Color(W1.Range["C1:T11"]);
            W1.Range["C1"].Value = "INITIALS";
            W1.Range["C3"].Value = "CREW";
            W1.Range["C5"].Value = "FROM";
            W1.Range["H1"].Value = "SHEET";
            W1.Range["H3"].Value = "FILENAME";
            W1.Range["H5"].Value = "TO";
            W1.Range["K1"].Value = "OF";
            W1.Range["P1"].Value = "DRAG";
            W1.Range["P3"].Value = "DATE";
            W1.Range["C11"].Value = "POINT";
            W1.Range["H11"].Value = "WELD";
            W1.Range["P10"].Value = "STATION";

            //datarow
            //Label_Color(W1.Range["B" + Row_Index.ToString() + ":Z" + (Row_Index + 8).ToString()]);
            W1.Range["C" + (Row_Index + 1).ToString()].Value = "NG:";
            W1.Range["C" + (Row_Index + 3).ToString()].Value = "CUT OFF:";
            W1.Range["C" + (Row_Index + 7).ToString()].Value = "POINT:";
            W1.Range["H" + (Row_Index + 1).ToString()].Value = "MMID:";
            W1.Range["H" + (Row_Index + 3).ToString()].Value = "PIPE:";
            W1.Range["H" + (Row_Index + 5).ToString()].Value = "HEAT:";
            W1.Range["H" + (Row_Index + 7).ToString()].Value = "WELD:";
            W1.Range["P" + (Row_Index + 1).ToString()].Value = "W.T.:";
            W1.Range["P" + (Row_Index + 3).ToString()].Value = "LENGTH:";
            W1.Range["P" + (Row_Index + 5).ToString()].Value = "STATION";
            #endregion

            #region Add_Borders
            //header
            Range TB_Intials = W1.Range["E1:F1"];
            Range TB_Crew = W1.Range["E3:F3"];
            Range TB_From = W1.Range["E5:F5"];
            Range TB_Sheet = W1.Range["J1"];
            Range TB_Filename = W1.Range["J3:N3"];
            Range TB_To = W1.Range["J5:L5"];
            Range TB_OF = W1.Range["L1"];
            Range TB_Drag = W1.Range["Q1"];
            Range TB_Date = W1.Range["Q3"];
            Range TB_Point = W1.Range["E11:F11"];
            Range TB_Weld = W1.Range["J11:L11"];
            Range TB_Station = W1.Range["P11:Q11"];

            Excel_Formatting.Border_Style_Bottom_Thin(TB_Intials);
            Excel_Formatting.Border_Style_Bottom_Thin(TB_Crew);
            Excel_Formatting.Border_Style_Bottom_Thin(TB_From);
            Excel_Formatting.Border_Style_Bottom_Thin(TB_Sheet);
            Excel_Formatting.Border_Style_Bottom_Thin(TB_Filename);
            Excel_Formatting.Border_Style_Bottom_Thin(TB_To);
            Excel_Formatting.Border_Style_Bottom_Thin(TB_OF);
            Excel_Formatting.Border_Style_Bottom_Thin(TB_Drag);
            Excel_Formatting.Border_Style_Bottom_Thin(TB_Date);
            Excel_Formatting.Border_Style_Bottom_Thin(TB_Point);
            Excel_Formatting.Border_Style_Bottom_Thin(TB_Weld);
            Excel_Formatting.Border_Style_Bottom_Thin(TB_Station);

            //DataRow
            Range Panel_1 = W1.Range["B" + Row_Index.ToString() + ":R" + (Row_Index + 8).ToString()]; // Thick Border
            Range Panel_2 = W1.Range["S" + Row_Index.ToString() + ":S" + (Row_Index + 8).ToString()]; // Thin Border
            Range Panel_3 = W1.Range["S" + Row_Index.ToString() + ":T" + (Row_Index + 8).ToString()]; // Thick Border
            Range Panel_4 = W1.Range["U" + Row_Index.ToString() + ":V" + (Row_Index + 8).ToString()]; // Thick Border Missing Right Side
            Range TB_NG = W1.Range["E" + (Row_Index + 1).ToString() + ":F" + (Row_Index + 1).ToString()]; // Thin Bottom Line
            Range TB_Cut_Off = W1.Range["E" + (Row_Index + 3).ToString() + ":F" + (Row_Index + 3).ToString()]; // Thin Bottom Line
            Range TB_Point2 = W1.Range["E" + (Row_Index + 7).ToString() + ":F" + (Row_Index + 7).ToString()]; // Thin Bottom Line
            Range TB_MMID = W1.Range["J" + (Row_Index + 1).ToString() + ":L" + (Row_Index + 1).ToString()]; // Thin Bottom Line
            Range TB_Pipe = W1.Range["J" + (Row_Index + 3).ToString() + ":L" + (Row_Index + 3).ToString()]; // Thin Bottom Line
            Range TB_Heat = W1.Range["J" + (Row_Index + 5).ToString() + ":L" + (Row_Index + 5).ToString()]; // Thin Bottom Line
            Range TB_Weld2 = W1.Range["J" + (Row_Index + 7).ToString() + ":L" + (Row_Index + 7).ToString()]; // Thin Bottom Line
            Range TB_WT = W1.Range["P" + (Row_Index + 2).ToString() + ":Q" + (Row_Index + 2).ToString()]; // Thin Bottom Line
            Range TB_LEN = W1.Range["P" + (Row_Index + 4).ToString() + ":Q" + (Row_Index + 4).ToString()]; // Thin Bottom Line
            Range TB_STA = W1.Range["P" + (Row_Index + 6).ToString() + ":Q" + (Row_Index + 6).ToString()]; // Thin Bottom Line

            Excel_Formatting.Border_Style_Thick(Panel_1);
            Excel_Formatting.Border_Style_Thin(Panel_2);
            Excel_Formatting.Border_Style_Thick(Panel_3);
            Excel_Formatting.Border_Style_Thick(Panel_4);
            Excel_Formatting.Border_Style_Remove_RightSide(Panel_4);
            Excel_Formatting.Border_Style_Bottom_Thin(TB_NG);
            Excel_Formatting.Border_Style_Bottom_Thin(TB_Cut_Off);
            Excel_Formatting.Border_Style_Bottom_Thin(TB_Point2);
            Excel_Formatting.Border_Style_Bottom_Thin(TB_MMID);
            Excel_Formatting.Border_Style_Bottom_Thin(TB_Pipe);
            Excel_Formatting.Border_Style_Bottom_Thin(TB_Heat);
            Excel_Formatting.Border_Style_Bottom_Thin(TB_Weld2);
            Excel_Formatting.Border_Style_Bottom_Thin(TB_WT);
            Excel_Formatting.Border_Style_Bottom_Thin(TB_LEN);
            Excel_Formatting.Border_Style_Bottom_Thin(TB_STA);
            #endregion

            #region Text_Alignment

            W1.Range["B" + Row_Index.ToString() + ":V" + (Row_Index + 8).ToString()].Cells.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            W1.Range["P9:Q20"].Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            W1.Range["J1:L1"].Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            W1.Range["J11"].Cells.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            W1.Range["E11"].Cells.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            W1.Range["U1:V69"].Font.Size = 8;
            W1.Range["U1:V69"].Cells.HorizontalAlignment = XlHAlign.xlHAlignLeft;

            #endregion

            if (Functions.IsNumeric(textBox_extra.Text))
            {
                extra1 = Convert.ToInt32(textBox_extra.Text);
            }

            string folder1 = WGEN_folder + @"ICONS\";
            a = (float)(W1.Range["A1"].ColumnWidth + W1.Range["B1"].ColumnWidth + W1.Range["C1"].ColumnWidth + W1.Range["D1"].ColumnWidth + W1.Range["E1"].ColumnWidth +
               W1.Range["F1"].ColumnWidth + W1.Range["G1"].ColumnWidth + W1.Range["H1"].ColumnWidth + W1.Range["I1"].ColumnWidth + W1.Range["J1"].ColumnWidth +
               W1.Range["K1"].ColumnWidth + W1.Range["L1"].ColumnWidth + W1.Range["M1"].ColumnWidth + W1.Range["N1"].ColumnWidth + W1.Range["O1"].ColumnWidth +
               W1.Range["P1"].ColumnWidth + W1.Range["Q1"].ColumnWidth + W1.Range["R1"].ColumnWidth + W1.Range["S1"].ColumnWidth) * 6 + 5 - 2 + extra1;

            W1.Shapes.AddPicture(folder1 + "legend.jpg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, a, 0, 144, 72); ;


            int j = 1;

            Row_Index = Row_start + Increment * j;
            while (j < 6)
            {
                W1.Rows["13:21"].Copy(W1.Rows[Row_Index]);
                ++j;
                Row_Index = Row_start + Increment * j;
            }
        }


        private void import_data()
        {
            Workbook Workbook1 = null;
            Microsoft.Office.Interop.Excel.Application Excel1 = null;
            Worksheet W1 = null;

            try
            {


                dt_weld = new System.Data.DataTable();
                dt_weld.Columns.Add(col_pnt, typeof(string));
                dt_weld.Columns.Add(col_northing, typeof(double));
                dt_weld.Columns.Add(col_easting, typeof(double));
                dt_weld.Columns.Add(col_elevation, typeof(double));
                dt_weld.Columns.Add(col_feature_code, typeof(string));
                dt_weld.Columns.Add(col_description, typeof(string));
                dt_weld.Columns.Add(col_station, typeof(double));
                dt_weld.Columns.Add(col_3d_stations, typeof(double));
                dt_weld.Columns.Add(col_mm_bk, typeof(string));
                dt_weld.Columns.Add(col_wall_bk, typeof(string));
                dt_weld.Columns.Add(col_pipe_bk, typeof(string));
                dt_weld.Columns.Add(col_heat_bk, typeof(string));
                dt_weld.Columns.Add(col_x_ray, typeof(string));
                dt_weld.Columns.Add(col_mm_ahd, typeof(string));
                dt_weld.Columns.Add(col_wall_ahd, typeof(string));
                dt_weld.Columns.Add(col_pipe_ahd, typeof(string));
                dt_weld.Columns.Add(col_heat_ahd, typeof(string));
                dt_weld.Columns.Add(col_length, typeof(string));
                dt_weld.Columns.Add(col_ng, typeof(string));
                dt_weld.Columns.Add(col_ng_northing, typeof(double));
                dt_weld.Columns.Add(col_ng_easting, typeof(double));
                dt_weld.Columns.Add(col_ng_elevation, typeof(double));
                dt_weld.Columns.Add(col_cover, typeof(string));
                dt_weld.Columns.Add(col_location, typeof(string));
                dt_weld.Columns.Add(col_filename, typeof(string));
                dt_weld.Columns.Add(col_spread, typeof(string));
                dt_weld.Columns.Add(col_v_angle, typeof(string));
                dt_weld.Columns.Add(col_h_angle, typeof(string));
                dt_weld.Columns.Add(col_restored_ng, typeof(string));
                dt_weld.Columns.Add(col_r_ng_northing, typeof(string));
                dt_weld.Columns.Add(col_r_ng_easting, typeof(string));
                dt_weld.Columns.Add(col_r_ng_elevation, typeof(string));
                dt_weld.Columns.Add(col_restored_cover, typeof(string));


                List<string> lista_col = new List<string>();
                List<string> lista_colxl = new List<string>();


                lista_col.Add(col_pnt);
                lista_col.Add(col_northing);
                lista_col.Add(col_easting);
                lista_col.Add(col_elevation);
                lista_col.Add(col_feature_code);
                lista_col.Add(col_description);
                lista_col.Add(col_station);
                lista_col.Add(col_3d_stations);
                lista_col.Add(col_mm_bk);
                lista_col.Add(col_wall_bk);
                lista_col.Add(col_pipe_bk);
                lista_col.Add(col_heat_bk);
                lista_col.Add(col_x_ray);
                lista_col.Add(col_mm_ahd);
                lista_col.Add(col_wall_ahd);
                lista_col.Add(col_pipe_ahd);
                lista_col.Add(col_heat_ahd);
                lista_col.Add(col_length);
                lista_col.Add(col_ng);
                lista_col.Add(col_ng_northing);
                lista_col.Add(col_ng_easting);
                lista_col.Add(col_ng_elevation);
                lista_col.Add(col_cover);
                lista_col.Add(col_location);
                lista_col.Add(col_filename);
                lista_col.Add(col_spread);
                lista_col.Add(col_v_angle);
                lista_col.Add(col_h_angle);
                lista_col.Add(col_restored_ng);
                lista_col.Add(col_r_ng_northing);
                lista_col.Add(col_r_ng_easting);
                lista_col.Add(col_r_ng_elevation);
                lista_col.Add(col_restored_cover);


                lista_colxl.Add(xl_col_pnt);
                lista_colxl.Add(xl_col_northing);
                lista_colxl.Add(xl_col_easting);
                lista_colxl.Add(xl_col_elevation);
                lista_colxl.Add(xl_col_feature_code);
                lista_colxl.Add(xl_col_description);
                lista_colxl.Add(xl_col_station);
                lista_colxl.Add(xl_col_3d_stations);
                lista_colxl.Add(xl_col_mm_bk);
                lista_colxl.Add(xl_col_wall_bk);
                lista_colxl.Add(xl_col_pipe_bk);
                lista_colxl.Add(xl_col_heat_bk);
                lista_colxl.Add(xl_col_x_ray);
                lista_colxl.Add(xl_col_mm_ahd);
                lista_colxl.Add(xl_col_wall_ahd);
                lista_colxl.Add(xl_col_pipe_ahd);
                lista_colxl.Add(xl_col_heat_ahd);
                lista_colxl.Add(xl_col_length);
                lista_colxl.Add(xl_col_ng);
                lista_colxl.Add(xl_col_ng_northing);
                lista_colxl.Add(xl_col_ng_easting);
                lista_colxl.Add(xl_col_ng_elevation);
                lista_colxl.Add(xl_col_cover);
                lista_colxl.Add(xl_col_location);
                lista_colxl.Add(xl_col_filename);
                lista_colxl.Add(xl_col_spread);
                lista_colxl.Add(xl_col_v_angle);
                lista_colxl.Add(xl_col_h_angle);
                lista_colxl.Add(xl_col_restored_ng);
                lista_colxl.Add(xl_col_r_ng_northing);
                lista_colxl.Add(xl_col_r_ng_easting);
                lista_colxl.Add(xl_col_r_ng_elevation);
                lista_colxl.Add(xl_col_restored_cover);



                if (comboBox_wm.Text != "")
                {
                    string string1 = comboBox_wm.Text;
                    if (string1.Contains("[") == true && string1.Contains("]") == true)
                    {
                        string filename = string1.Substring(string1.IndexOf("]") + 4, string1.Length - (string1.IndexOf("]") + 4));
                        string sheet_name = string1.Substring(1, string1.IndexOf("]") - 1);
                        if (filename.Length > 0 && sheet_name.Length > 0)
                        {
                            W1 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, sheet_name);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Weld Map spreadsheet Not Selected", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }



                dt_weld = Functions.build_data_table_from_excel(dt_weld, W1, start_row, 30000, lista_col, lista_colxl);


            }

            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (W1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W1);
                if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
            }

        }


        private void button_Export_weld_sheets_Click(object sender, EventArgs e)
        {
            Workbook Workbook1 = null;
            Microsoft.Office.Interop.Excel.Application Excel1 = null;
            Worksheet W1 = null;
            try
            {
                set_enable_false();
                import_data();



                if (dt_weld != null && dt_weld.Rows.Count > 0)
                {

                    dt_weld = Functions.Sort_data_table(dt_weld, col_station);

                    Excel1 = Functions.Excel_Open();
                    Workbook1 = Excel1.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                    Excel1.ActiveWindow.DisplayGridlines = false;


                    int existing_number_of_sheets = Workbook1.Worksheets.Count;
                    if (existing_number_of_sheets == 0)
                    {
                        Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[1], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                    }

                    Excel1.DisplayAlerts = false;
                    if (existing_number_of_sheets > 1)
                    {
                        for (int i = 2; i <= existing_number_of_sheets; ++i)
                        {
                            Workbook1.Worksheets[i].Delete();
                        }
                    }

                    W1 = Workbook1.Worksheets[1];




                    Page_Setup(W1);


                    Populate_data(Workbook1, W1, dt_weld);

                    W1.Delete();


                    for (int i = 1; i <= Workbook1.Worksheets.Count; ++i)
                    {
                        W1 = Workbook1.Worksheets[i];
                        W1.Range["J1"].Value2 = i.ToString();
                        W1.Range["L1"].Value2 = Workbook1.Worksheets.Count.ToString();

                    }

                    Excel1.DisplayAlerts = true;



                    MessageBox.Show("Done", "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (W1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W1);
                if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
            }
            set_enable_true();
        }


        public void Populate_data(Workbook Workbook1, Worksheet W1, System.Data.DataTable dt1)
        {
            string folder1 = WGEN_folder + @"ICONS\";

            Worksheet W2 = Workbook1.Worksheets[Workbook1.Worksheets.Count];
            Worksheet W_previous = null;
            W1.Copy(After: W2);
            W2 = Workbook1.Worksheets[Workbook1.Worksheets.Count];

            int left_index = 58;

            int increment = 9;
            int tab_name = 1;
            W2.Name = tab_name.ToString();

            //Functions.Transfer_datatable_to_new_excel_spreadsheet(dt1);

            double sta_p = 0;

            int index_lista_right = 0;
            int no_right = 0;


            List<string> lista_right = new List<string>();
            lista_right.Add("U63");
            lista_right.Add("U59");
            lista_right.Add("U54");
            lista_right.Add("U50");
            lista_right.Add("U45");
            lista_right.Add("U41");
            lista_right.Add("U36");
            lista_right.Add("U32");
            lista_right.Add("U27");
            lista_right.Add("U23");
            lista_right.Add("U18");
            lista_right.Add("U14");


            bool add_new_page = false;
            bool write_on_previous_page = false;
            bool write_start_station = false;


            float size_symbol = 22;

            float r6_2 = 125;
            float r6_1 = 180;

            float r5_2 = 225;
            float r5_1 = 280;

            float r4_2 = 325;
            float r4_1 = 380;

            float r3_2 = 425;
            float r3_1 = 480;

            float r2_2 = 525;
            float r2_1 = 580;

            float r1_2 = 625;
            float r1_1 = 680;




            for (int i = 0; i < dt1.Rows.Count; ++i)
            {

                if (dt1.Rows[i][col_feature_code] != DBNull.Value && dt1.Rows[i][col_station] != DBNull.Value)
                {
                    string fc = Convert.ToString(dt1.Rows[i][col_feature_code]).ToUpper();
                    double sta = Convert.ToDouble(dt1.Rows[i][col_station]);

                    if (fc == "WELD")
                    {

                        no_right = 0;
                        if (left_index < 13)
                        {
                            add_new_page = true;
                        }
                        else
                        {
                            add_new_page = false;
                        }

                        if (add_new_page == true)
                        {
                            W_previous = W2;
                            W1.Copy(After: W2);
                            W2 = Workbook1.Worksheets[Workbook1.Worksheets.Count];
                            ++tab_name;
                            W2.Name = tab_name.ToString();
                            left_index = 58;
                            index_lista_right = -1;

                        }


                        double len = sta - sta_p;
                        W2.Range["E" + (left_index + 1).ToString()].Value = dt1.Rows[i][col_ng];
                        W2.Range["E" + (left_index + 3).ToString()].Value = "";
                        W2.Range["E" + (left_index + 7).ToString()].Value = dt1.Rows[i][col_pnt];
                        W2.Range["J" + (left_index + 1).ToString()].Value = dt1.Rows[i][col_mm_ahd];
                        W2.Range["J" + (left_index + 3).ToString()].Value = dt1.Rows[i][col_pipe_ahd];
                        W2.Range["J" + (left_index + 5).ToString()].Value = dt1.Rows[i][col_heat_ahd];
                        W2.Range["J" + (left_index + 7).ToString()].Value = dt1.Rows[i][col_description];
                        W2.Range["P" + (left_index + 2).ToString()].Value = dt1.Rows[i][col_wall_ahd];
                        W2.Range["P" + (left_index + 4).ToString()].Value = len;
                        W2.Range["P" + (left_index + 6).ToString()].Value = Functions.Get_chainage_from_double(sta, "f", 0);

                        if (write_on_previous_page == false && left_index == 58 && Workbook1.Worksheets.Count > 2)
                        {
                            write_on_previous_page = true;
                        }


                        if (write_on_previous_page == true)
                        {
                            W_previous.Range["E11"].Value = dt1.Rows[i][col_pnt];
                            W_previous.Range["J11"].Value = dt1.Rows[i][col_description];
                            W_previous.Range["P11"].Value = Functions.Get_chainage_from_double(sta, "f", 0);
                            write_on_previous_page = false;
                        }


                        if (left_index == 58 && dt1.Rows[i][col_filename] != DBNull.Value)
                        {
                            string filename = Convert.ToString(dt1.Rows[i][col_filename]);
                            W2.Range["J3"].Value = filename;
                            if (filename.Length > 2)
                            {
                                W2.Range["E1"].Value = filename.Substring(0, 2);

                                string date_string = filename.Substring(2, filename.Length - 2);
                                string new_date_string = "";
                                for (int k = 0; k < date_string.Length; ++k)
                                {
                                    string leterr1 = date_string.Substring(k, 1);
                                    if (Functions.IsNumeric(leterr1) == true)
                                    {
                                        new_date_string = new_date_string + leterr1;
                                    }
                                    else
                                    {
                                        k = date_string.Length;
                                    }
                                }



                                W2.Range["Q3"].Value = new_date_string;
                            }
                        }

                        if (left_index == 58)
                        {
                            W2.Range["E5"].Value = Functions.Get_chainage_from_double(sta, "f", 0);
                        }


                        if (write_start_station == false && left_index == 13)
                        {

                            write_start_station = true;
                        }


                        if (write_start_station == true)
                        {
                            W2.Range["J5"].Value = Functions.Get_chainage_from_double(sta, "f", 0);
                            write_start_station = false;
                        }



                        switch (left_index)
                        {
                            case 58:
                                index_lista_right = 0;
                                no_right = 0;
                                break;

                            case 49:
                                index_lista_right = 2;
                                no_right = 0;
                                break;

                            case 40:
                                index_lista_right = 4;
                                no_right = 0;
                                break;

                            case 31:
                                index_lista_right = 6;
                                no_right = 0;
                                break;

                            case 22:
                                index_lista_right = 8;
                                no_right = 0;
                                break;

                            case 13:
                                index_lista_right = 10;
                                no_right = 0;
                                break;
                            default:
                                index_lista_right = 0;
                                no_right = 0;
                                break;
                        }

                        left_index = left_index - increment;

                        sta_p = sta;
                    }
                    else
                    {


                        ++no_right;
                        switch (no_right)
                        {
                            case 1:
                                break;
                            case 2:
                                break;
                            case 3:
                                left_index = left_index - increment;
                                break;
                            case 4:
                                break;
                            case 5:
                                left_index = left_index - increment;
                                break;
                            case 6:
                                break;
                            case 7:
                                left_index = left_index - increment;
                                break;
                            case 8:
                                break;
                            case 9:
                                left_index = left_index - increment;
                                break;
                            case 10:
                                break;
                            case 11:
                                left_index = left_index - increment;
                                break;
                            case 12:
                                break;
                            case 13:
                                index_lista_right = 0;
                                no_right = 0;
                                W_previous = W2;
                                W1.Copy(After: W2);
                                W2 = Workbook1.Worksheets[Workbook1.Worksheets.Count];
                                ++tab_name;
                                W2.Name = tab_name.ToString();
                                left_index = 58 - increment;
                                write_on_previous_page = true;
                                write_start_station = true;
                                break;
                            default:
                                break;
                        }




                        if (index_lista_right == 12)
                        {
                            index_lista_right = 0;
                            no_right = 0;
                            W_previous = W2;
                            W1.Copy(After: W2);
                            W2 = Workbook1.Worksheets[Workbook1.Worksheets.Count];
                            ++tab_name;
                            W2.Name = tab_name.ToString();
                            left_index = 58 - increment;
                            write_on_previous_page = true;
                            write_start_station = true;

                        }



                        string description = "";
                        if (dt1.Rows[i][col_description] != DBNull.Value)
                        {
                            description = Convert.ToString(dt1.Rows[i][col_description]);
                        }


                        string pnt = "";
                        if (dt1.Rows[i][col_pnt] != DBNull.Value)
                        {
                            pnt = Convert.ToString(dt1.Rows[i][col_pnt]);
                        }


                        string chainage = Functions.Get_chainage_from_double(sta, "f", 0);

                        float ins_pt = r1_1;

                        switch (index_lista_right)
                        {
                            case 0:
                                ins_pt = r1_1;
                                break;
                            case 1:
                                ins_pt = r1_2;
                                break;


                            case 2:
                                ins_pt = r2_1;
                                break;
                            case 3:
                                ins_pt = r2_2;
                                break;


                            case 4:
                                ins_pt = r3_1;
                                break;
                            case 5:
                                ins_pt = r3_2;
                                break;


                            case 6:
                                ins_pt = r4_1;
                                break;
                            case 7:
                                ins_pt = r4_2;
                                break;


                            case 8:
                                ins_pt = r5_1;
                                break;
                            case 9:
                                ins_pt = r5_2;
                                break;


                            case 10:
                                ins_pt = r6_1;
                                break;
                            case 11:
                                ins_pt = r6_2;
                                break;


                            default:
                                ins_pt = r1_1;
                                break;
                        }


                        if (fc == "BEND")
                        {
                            string vangle = "";
                            if (dt1.Rows[i][col_v_angle] != DBNull.Value)
                            {
                                vangle = Convert.ToString(dt1.Rows[i][col_v_angle]);
                            }
                            string hangle = "";
                            if (dt1.Rows[i][col_h_angle] != DBNull.Value)
                            {
                                hangle = Convert.ToString(dt1.Rows[i][col_h_angle]);
                            }
                            double dbl_vangle = 0;
                            if (Functions.IsNumeric(vangle) == true)
                            {
                                dbl_vangle = Convert.ToDouble(vangle);
                            }
                            double dbl_hangle = 0;
                            if (Functions.IsNumeric(hangle) == true)
                            {
                                dbl_hangle = Convert.ToDouble(hangle);
                            }

                            if (dbl_hangle > 0)
                            {
                                description = description + " HOR=" + hangle;
                            }
                            if (dbl_vangle > 0)
                            {
                                description = description + " VER=" + vangle;
                            }

                            if (dbl_hangle > 0 && dbl_vangle == 0)
                            {
                                if (description.ToUpper().Contains("LEFT") == true)
                                {
                                    W2.Shapes.AddPicture(folder1 + "PI left.jpg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, a, ins_pt, size_symbol, size_symbol);
                                }
                                else
                                {
                                    W2.Shapes.AddPicture(folder1 + "PI right.jpg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, a, ins_pt, size_symbol, size_symbol);
                                }
                            }

                            if (dbl_hangle > 0 && dbl_vangle > 0)
                            {
                                if (description.ToUpper().Contains("SAG") == true)
                                {
                                    if (description.ToUpper().Contains("LEFT") == true)
                                    {
                                        W2.Shapes.AddPicture(folder1 + "Sag left.jpg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, a, ins_pt, size_symbol, size_symbol);
                                    }
                                    else
                                    {
                                        W2.Shapes.AddPicture(folder1 + "Sag right.jpg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, a, ins_pt, size_symbol, size_symbol);
                                    }
                                }
                                else
                                {
                                    if (description.ToUpper().Contains("LEFT") == true)
                                    {
                                        W2.Shapes.AddPicture(folder1 + "Overbend left.jpg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, a, ins_pt, size_symbol, size_symbol);
                                    }
                                    else
                                    {
                                        W2.Shapes.AddPicture(folder1 + "Overbend right.jpg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, a, ins_pt, size_symbol, size_symbol);
                                    }
                                }

                            }

                            if (dbl_hangle == 0 && dbl_vangle > 0)
                            {
                                if (description.ToUpper().Contains("SAG") == true)
                                {
                                    W2.Shapes.AddPicture(folder1 + "Sag.jpg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, a, ins_pt, size_symbol, size_symbol);
                                }
                                else
                                {
                                    W2.Shapes.AddPicture(folder1 + "Overbend.jpg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, a, ins_pt, size_symbol, size_symbol);
                                }
                            }


                        }

                        W2.Range[lista_right[index_lista_right]].Value = fc + " " + pnt + "\r\n" + chainage + " " + description;

                        if (fc.ToUpper().Contains("CAD") == true && fc.ToUpper().Contains("WELD") == true)
                        {
                            W2.Shapes.AddPicture(folder1 + "CAD weld.jpg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, a, ins_pt, size_symbol, size_symbol);
                        }

                        if (fc.ToUpper().Contains("PIPELINE") == true)
                        {
                            W2.Shapes.AddPicture(folder1 + "Existing pipeline.jpg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, a, ins_pt, size_symbol, size_symbol);
                        }

                        if (fc.ToUpper().Contains("ROCK") == true && fc.ToUpper().Contains("SHIELD") == true)
                        {
                            W2.Shapes.AddPicture(folder1 + "Rock Shield.jpg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, a, ins_pt, size_symbol, size_symbol);
                        }


                        if (fc.ToUpper().Contains("TRENCH") == true && fc.ToUpper().Contains("BREAKERS") == true)
                        {
                            W2.Shapes.AddPicture(folder1 + "Trench breaker.jpg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, a, ins_pt, size_symbol, size_symbol);
                        }

                        if (fc.ToUpper().Contains("WEIGHT") == true)
                        {
                            W2.Shapes.AddPicture(folder1 + "Weight.jpg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, a, ins_pt, size_symbol, size_symbol);
                        }


                        if (fc.ToUpper().Contains("LOOSE") == true && fc.ToUpper().Contains("END") == true)
                        {
                            W2.Shapes.AddPicture(folder1 + "Loose end.jpg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, a, ins_pt, size_symbol, size_symbol);
                        }

                        ++index_lista_right;

                    }

                }
            }
        }


        #region drop downs
        private void comboBox_Pipe_DropDown(object sender, EventArgs e)
        {
            Functions.Load_opened_worksheets_to_combobox(comboBox_wm);
        }


        #endregion


        private void label_mm_Click(object sender, EventArgs e)
        {
            comboBox_wm.Items.Clear();


        }

        private void button_load_settings_Click(object sender, EventArgs e)
        {
            if (textBox_extra.Visible == true)
            {
                textBox_extra.Visible = false;
            }
            else
            {
                textBox_extra.Visible = true;
            }

        }
    }
}
