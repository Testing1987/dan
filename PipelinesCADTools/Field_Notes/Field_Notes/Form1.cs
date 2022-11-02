using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;

namespace weld_sheet
{
    public partial class Form_Field_Notes : Form
    {
        //Global Variables
        System.Data.DataTable dt_weld = null;
        System.Data.DataTable dt_pipe = null;
        System.Data.DataTable DT_Bend = null;
        System.Data.DataTable dt_pipe_and_bend = null;
        System.Data.DataTable dt_compiled = null;
        //System.Data.DataTable DT_Missing_From_Spot = null;
        public Form_Field_Notes FN_Form = null;
        Worksheet W1 = null; // pipe data
        Worksheet W2 = null; // weld data
        Worksheet W3 = null; // bend data

        #region spot columns
        //Weld
        string spot_weld_col_Point_No = "WELD NUMBER";
        string spot_weld_col_STA = "STATION NUMBER";
        string spot_weld_col_Pipe_Bk = "BEHIND PIPE DESCRIPTION";
        string spot_weld_col_Pipe_Ahd = "AHEAD PIPE DESCRIPTION";
        string spot_weld_col_MM_Bk = "BEHIND PIPE USER NUMBER";
        string spot_weld_col_MM_Ahd = "AHEAD PIPE USER NUMBER";
        string spot_weld_col_PipeNo_Bk = "BEHIND PIPE FRIENDLY NUMBER";
        string spot_weld_col_PipeNo_Ahd = "AHEAD PIPE FRIENDLY NUMBER";

        //Pipe
        string spot_pipe_col_Desc = "DESCRIPTION";
        string spot_pipe_col_No = "NUMBER";
        string spot_pipe_col_PipeNo = "FRIENDLY NUMBER";
        string spot_pipe_col_Heat = "HEAT";
        string spot_pipe_col_Len = "MEASURED LENGTH";
        string spot_pipe_col_wthk = "WALL THICKNESS";

        //Bend
        string spot_bend_col_ang = "ANGLE";
        string spot_bend_col_type = "BEND TYPE";
        string spot_bend_col_Desc = "DESCRIPTION";

        #endregion

        #region Weld Tally Columns
        string col_NG = "NG";
        string col_Cutoff = "CUTOFF";
        string col_Pt = "POINT";
        string col_MMID = "MMID";
        string col_Pipe = "PIPE";
        string col_Heat = "HEAT";
        string col_Weld = "WELD";
        string col_Wallthk = "W.T.";
        string col_Len = "LEN";
        string col_STA = "STA";


        string col_source = "Source";
        string col_missing = "Missing From";
        string col_le = "Loose End";
        string col_drag = "DRAG";
        string col_st_dbl = "STATION AS DBL";

        string xl_back = "H";
        string xl_ahd = "E";

        #endregion

        Microsoft.Office.Interop.Excel.Application Excel1 = Functions.Excel1;
        private bool clickdragdown;
        private System.Drawing.Point lastLocation;

        public Form_Field_Notes()
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
            lista_butoane.Add(comboBox_Pipe);
            lista_butoane.Add(comboBox_Weld);
            lista_butoane.Add(comboBox_Bend);

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
            lista_butoane.Add(comboBox_Pipe);
            lista_butoane.Add(comboBox_Bend);
            lista_butoane.Add(comboBox_Weld);

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
            W1.Columns["E:E"].ColumnWidth = 7;
            W1.Columns["F:F"].ColumnWidth = 7;
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
            W1.Columns["U:U"].ColumnWidth = 11;
            W1.Columns["V:V"].ColumnWidth = 11;
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

            Shape shp = W1.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeUpArrow, 273, 60, 30, 825);
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
            W1.Range["J11"].Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            #endregion

            int j = 1;

            Row_Index = Row_start + Increment * j;
            while (j < 6)
            {
                W1.Rows["13:21"].Copy(W1.Rows[Row_Index]);
                ++j;
                Row_Index = Row_start + Increment * j;
            }

        }

        public void Sort_Table_By_Column(ref System.Data.DataTable dt1, string col_name, System.Data.DataTable dt3)
        {

            List<System.Data.DataTable> result = dt1.AsEnumerable().GroupBy(row => row.Field<int>(col_name)).Select(g => g.CopyToDataTable()).ToList();


            for (int i = 0; i < result.Count; ++i)
            {
                System.Data.DataTable dt2 = new System.Data.DataTable();
                dt2 = result[i];

                dt2 = Functions.Sort_data_table(dt2, col_st_dbl);


                string pipe_bk = "";
                string mm_bk = "";
                double pipe_length = -1000000000;

                if (dt2.Rows[0][spot_weld_col_Pipe_Bk] != DBNull.Value)
                {
                    pipe_bk = Convert.ToString(dt2.Rows[0][spot_weld_col_Pipe_Bk]);

                    for (int j = 0; j < dt3.Rows.Count; ++j)
                    {
                        if (dt3.Rows[j][spot_pipe_col_Desc] != DBNull.Value)
                        {
                            string pipe2 = Convert.ToString(dt3.Rows[j][spot_pipe_col_Desc]).ToUpper();

                            if (pipe_bk.ToUpper() == pipe2)
                            {
                                if (dt3.Rows[j][spot_pipe_col_Len] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt3.Rows[j][spot_pipe_col_Len])) == true)
                                {
                                    pipe_length = Convert.ToDouble(dt3.Rows[j][spot_pipe_col_Len]);

                                    j = dt3.Rows.Count;
                                }
                            }
                        }
                    }
                }

                if (dt2.Rows[0][spot_weld_col_MM_Bk] != DBNull.Value)
                {
                    mm_bk = Convert.ToString(dt2.Rows[0][spot_weld_col_MM_Bk]);
                }

                double sta1 = Convert.ToDouble(dt2.Rows[0]["STATION AS DBL"]);


                System.Data.DataRow row1 = dt2.NewRow();


                if (pipe_bk != "")
                {
                    row1[spot_weld_col_Pipe_Ahd] = pipe_bk;
                }

                if (mm_bk != "")
                {
                    row1[spot_weld_col_MM_Ahd] = mm_bk;
                }

                row1[spot_weld_col_Point_No] = "LE";
                row1[col_drag] = dt2.Rows[0][col_drag];
                row1[spot_weld_col_STA] = sta1 - pipe_length;

                dt2.Rows.InsertAt(row1, 0);

                System.Data.DataRow row2 = dt2.NewRow();
                row2[spot_weld_col_Point_No] = "LE";

                dt2.Rows.InsertAt(row2, dt2.Rows.Count - 1);

                dt2.Rows[0][col_le] = true;

                if (i == 0)
                {
                    dt1 = dt2.Clone();
                }

                for (int j = 0; j < dt2.Rows.Count; ++j)
                {
                    dt1.Rows.Add();
                    dt1.Rows[dt1.Rows.Count - 1].ItemArray = dt2.Rows[j].ItemArray;
                }

            }

        }

        private void import_data()
        {
            try
            {
                #region Scan Pipe File

                dt_pipe = new System.Data.DataTable();
                string Error_Message = "Pipe File Scan Error";
                W1 = null;
                if (comboBox_Pipe.Text != "")
                {
                    string string1 = comboBox_Pipe.Text;
                    if (string1.Contains("[") == true && string1.Contains("]") == true)
                    {
                        string filename = string1.Substring(string1.IndexOf("]") + 4, string1.Length - (string1.IndexOf("]") + 4));
                        string sheet_name = string1.Substring(1, string1.IndexOf("]") - 1);
                        if (filename.Length > 0 && sheet_name.Length > 0)
                        {
                            W1 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, sheet_name);
                        }
                        else
                        {

                        }
                    }
                    else
                    {

                    }
                }
                else
                {
                    MessageBox.Show("Pipe Data Not Selected", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                string start_row_string = "2";
                int start_row = Convert.ToInt32(start_row_string);
                string end_row_string = Convert.ToString(Functions.get_last_row_excel(W1, "A", 1));
                int end_row = Convert.ToInt32(end_row_string);
                string start_column_string = "A";
                string end_column_string = Functions.get_last_column_from_excel(W1, 1);

                //end_column_string = get_last_column_from_excel(W1, 1);

                Range sc_convert_start = W1.Range[start_column_string + "1"];
                Range sc_convert_end = W1.Range[end_column_string + "1"];
                int start_column = sc_convert_start.Column;
                int end_column = sc_convert_end.Column;
                Range Row1 = W1.Range[start_column_string + "1" + ":" + end_column_string + "1"];
                Range range1 = W1.Range[start_column_string + start_row_string + ":" + end_column_string + end_row_string];

                dt_pipe = Functions.Populate_data_table_with_excel_range("Pipe File", dt_pipe, start_row, end_row, start_column, end_column, Row1, range1, Error_Message);


                //Common_Methods.Transfer_datatable_to_new_excel_spreadsheet(DT_Pipe);

                #endregion

                #region Scan Weld File
                dt_weld = new System.Data.DataTable();

                Error_Message = "Weld File Scan Error";
                W2 = null;

                if (comboBox_Weld.Text != "")
                {
                    string string1 = comboBox_Weld.Text;
                    if (string1.Contains("[") == true && string1.Contains("]") == true)
                    {
                        string filename = string1.Substring(string1.IndexOf("]") + 4, string1.Length - (string1.IndexOf("]") + 4));
                        string sheet_name = string1.Substring(1, string1.IndexOf("]") - 1);
                        if (filename.Length > 0 && sheet_name.Length > 0)
                        {
                            W2 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, sheet_name);
                        }
                        else
                        {

                        }
                    }
                    else
                    {

                    }
                }
                else
                {
                    MessageBox.Show("Weld Data Not Selected", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }


                start_row_string = "2";
                start_row = Convert.ToInt32(start_row_string);
                end_row_string = Convert.ToString(Functions.get_last_row_excel(W2, "A", 1));
                end_row = Convert.ToInt32(end_row_string);
                start_column_string = "A";
                end_column_string = Functions.get_last_column_from_excel(W2, 1);

                sc_convert_start = W2.Range[start_column_string + "1"];
                sc_convert_end = W2.Range[end_column_string + "1"];
                start_column = sc_convert_start.Column;
                end_column = sc_convert_end.Column;
                Row1 = W2.Range[start_column_string + "1" + ":" + end_column_string + "1"];
                range1 = W2.Range[start_column_string + start_row_string + ":" + end_column_string + end_row_string];

                dt_weld = Functions.Populate_data_table_with_excel_range("Weld File", dt_weld, start_row, end_row, start_column, end_column, Row1, range1, Error_Message);

                dt_weld.Columns.Add(col_st_dbl, typeof(double));
                dt_weld.Columns.Add(col_le, typeof(bool));
                dt_weld.Columns.Add(col_drag, typeof(int));

                dt_weld.Columns[col_drag].SetOrdinal(0);
                dt_weld.Columns[col_le].SetOrdinal(0);


                bool is_le = false;
                int k = 1;

                for (int i = 0; i < dt_weld.Rows.Count; ++i)
                {
                    if (dt_weld.Rows[i][spot_weld_col_STA] != DBNull.Value)
                    {
                        double sta_dbl = Convert.ToDouble(Convert.ToString(dt_weld.Rows[i][spot_weld_col_STA]).Replace("+", ""));
                        dt_weld.Rows[i][spot_weld_col_STA] = sta_dbl;
                        dt_weld.Rows[i][col_st_dbl] = sta_dbl;
                    }

                    string assest_back = Convert.ToString(dt_weld.Rows[i][dt_weld.Columns[spot_weld_col_MM_Bk]]);
                    string assest_ahead = "";

                    if (i > 0)
                    {
                        assest_ahead = Convert.ToString(dt_weld.Rows[i - 1][dt_weld.Columns[spot_weld_col_MM_Ahd]]);

                    }

                    if (i > 0)
                    {
                        if (assest_ahead != assest_back)
                        {
                            k = k + 1;
                            dt_weld.Rows[i][col_drag] = k;
                        }
                        else
                        {
                            dt_weld.Rows[i][col_drag] = k;
                        }
                    }

                    if (i == 0)
                    {
                        dt_weld.Rows[i][col_drag] = k;
                    }

                }



                Sort_Table_By_Column(ref dt_weld, col_drag, dt_pipe);

                dt_weld.Columns.Remove(col_st_dbl);

                #endregion

                #region Scan Bend File
                DT_Bend = new System.Data.DataTable();
                Error_Message = "Bend File Scan Error";
                W3 = null;
                if (comboBox_Bend.Text != "")
                {
                    string string1 = comboBox_Bend.Text;
                    if (string1.Contains("[") == true && string1.Contains("]") == true)
                    {
                        string filename = string1.Substring(string1.IndexOf("]") + 4, string1.Length - (string1.IndexOf("]") + 4));
                        string sheet_name = string1.Substring(1, string1.IndexOf("]") - 1);
                        if (filename.Length > 0 && sheet_name.Length > 0)
                        {
                            W3 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, sheet_name);
                        }
                        else
                        {

                        }
                    }
                    else
                    {

                    }

                    start_row_string = "2";
                    start_row = Convert.ToInt32(start_row_string);
                    end_row_string = Convert.ToString(Functions.get_last_row_excel(W3, "A", 1));
                    end_row = Convert.ToInt32(end_row_string);
                    start_column_string = "A";
                    end_column_string = Functions.get_last_column_from_excel(W3, 1);

                    //end_column_string = get_last_column_from_excel(W3, 1);

                    sc_convert_start = W3.Range[start_column_string + "1"];
                    sc_convert_end = W3.Range[end_column_string + "1"];
                    start_column = sc_convert_start.Column;
                    end_column = sc_convert_end.Column;
                    Row1 = W3.Range[start_column_string + "1" + ":" + end_column_string + "1"];
                    range1 = W3.Range[start_column_string + start_row_string + ":" + end_column_string + end_row_string];

                    DT_Bend = Functions.Populate_data_table_with_excel_range("Bend File", DT_Bend, start_row, end_row, start_column, end_column, Row1, range1, Error_Message);


                    //Common_Methods.Transfer_datatable_to_new_excel_spreadsheet(DT_Bend);
                }
                else
                {
                    //MessageBox.Show("Bend Data Not Selected", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //return;
                }





                #endregion

                #region Combine Pipe & Bend Tables

                dt_pipe_and_bend = new System.Data.DataTable();
                string col_Type = "XTYPE";

                dt_pipe_and_bend.Columns.Add(spot_pipe_col_Desc, typeof(string));
                dt_pipe_and_bend.Columns.Add(spot_pipe_col_No, typeof(string));
                dt_pipe_and_bend.Columns.Add(spot_pipe_col_PipeNo, typeof(string));
                dt_pipe_and_bend.Columns.Add(spot_pipe_col_Heat, typeof(string));
                dt_pipe_and_bend.Columns.Add(spot_pipe_col_Len, typeof(string));
                dt_pipe_and_bend.Columns.Add(spot_pipe_col_wthk, typeof(string));
                dt_pipe_and_bend.Columns.Add(col_Type, typeof(string));
                dt_pipe_and_bend.Columns.Add(spot_bend_col_ang, typeof(string));
                dt_pipe_and_bend.Columns.Add(spot_bend_col_type, typeof(string));

                if (dt_pipe != null && dt_pipe.Rows.Count > 0)
                {
                    for (int i = 0; i < dt_pipe.Rows.Count; ++i)
                    {
                        dt_pipe_and_bend.Rows.Add();
                        dt_pipe_and_bend.Rows[dt_pipe_and_bend.Rows.Count - 1][spot_pipe_col_Desc] = dt_pipe.Rows[i][spot_pipe_col_Desc];
                        dt_pipe_and_bend.Rows[dt_pipe_and_bend.Rows.Count - 1][spot_pipe_col_No] = dt_pipe.Rows[i][spot_pipe_col_No];
                        dt_pipe_and_bend.Rows[dt_pipe_and_bend.Rows.Count - 1][spot_pipe_col_PipeNo] = dt_pipe.Rows[i][spot_pipe_col_PipeNo];
                        dt_pipe_and_bend.Rows[dt_pipe_and_bend.Rows.Count - 1][spot_pipe_col_Heat] = dt_pipe.Rows[i][spot_pipe_col_Heat];
                        dt_pipe_and_bend.Rows[dt_pipe_and_bend.Rows.Count - 1][spot_pipe_col_Len] = dt_pipe.Rows[i][spot_pipe_col_Len];
                        dt_pipe_and_bend.Rows[dt_pipe_and_bend.Rows.Count - 1][spot_pipe_col_wthk] = dt_pipe.Rows[i][spot_pipe_col_wthk];
                        dt_pipe_and_bend.Rows[dt_pipe_and_bend.Rows.Count - 1][col_Type] = "Pipe";

                    }
                }

                if (comboBox_Bend.Text != "" && DT_Bend != null && DT_Bend.Rows.Count > 0)
                {
                    for (int i = 0; i < DT_Bend.Rows.Count; ++i)
                    {
                        dt_pipe_and_bend.Rows.Add();
                        dt_pipe_and_bend.Rows[dt_pipe_and_bend.Rows.Count - 1][spot_pipe_col_Desc] = DT_Bend.Rows[i][spot_pipe_col_Desc];
                        dt_pipe_and_bend.Rows[dt_pipe_and_bend.Rows.Count - 1][spot_pipe_col_No] = DT_Bend.Rows[i][spot_pipe_col_Desc];
                        dt_pipe_and_bend.Rows[dt_pipe_and_bend.Rows.Count - 1][spot_pipe_col_PipeNo] = DT_Bend.Rows[i][spot_pipe_col_Desc];
                        dt_pipe_and_bend.Rows[dt_pipe_and_bend.Rows.Count - 1][col_Type] = "Bend";
                        dt_pipe_and_bend.Rows[dt_pipe_and_bend.Rows.Count - 1][spot_bend_col_ang] = DT_Bend.Rows[i][spot_bend_col_ang];
                        dt_pipe_and_bend.Rows[dt_pipe_and_bend.Rows.Count - 1][spot_bend_col_type] = DT_Bend.Rows[i][spot_bend_col_type];
                    }
                }

                //Functions.Transfer_datatable_to_new_excel_spreadsheet(dt_pipe_and_bend);

                #endregion

                #region Define New Data Table
                dt_compiled = new System.Data.DataTable();
                //DT_Missing_From_Spot = new System.Data.DataTable();

                dt_compiled.Columns.Add(col_NG, typeof(string));
                dt_compiled.Columns.Add(col_Cutoff, typeof(string));
                dt_compiled.Columns.Add(col_Pt, typeof(string));
                dt_compiled.Columns.Add(col_MMID, typeof(string));
                dt_compiled.Columns.Add(col_Pipe, typeof(string));
                dt_compiled.Columns.Add(col_Heat, typeof(string));
                dt_compiled.Columns.Add(col_Weld, typeof(string));
                dt_compiled.Columns.Add(col_Wallthk, typeof(string));
                dt_compiled.Columns.Add(col_Len, typeof(string));
                dt_compiled.Columns.Add(col_STA, typeof(string));
                dt_compiled.Columns.Add(col_le, typeof(bool));
                dt_compiled.Columns.Add(col_drag, typeof(int));
                dt_compiled.Columns.Add(spot_bend_col_ang, typeof(string));
                dt_compiled.Columns.Add(spot_bend_col_type, typeof(string));


                //DT_Missing_From_Spot.Columns.Add(col_Pipe, typeof(string));
                //DT_Missing_From_Spot.Columns.Add(col_source, typeof(string));
                //DT_Missing_From_Spot.Columns.Add(col_missing, typeof(string));

                #endregion

                #region Compile_Data

                dt_weld.TableName = "weld";
                dt_pipe_and_bend.TableName = "pipebend";
                DataSet dataset1 = new DataSet();
                dataset1.Tables.Add(dt_weld);
                dataset1.Tables.Add(dt_pipe_and_bend);


                DataRelation relation1 = new DataRelation("xxx", dt_weld.Columns[spot_weld_col_Pipe_Ahd], dt_pipe_and_bend.Columns[spot_pipe_col_Desc], false);
                dataset1.Relations.Add(relation1);



                for (int i = 0; i < dt_weld.Rows.Count; ++i)
                {
                    if (dt_weld.Rows[i].GetChildRows(relation1).Length > 0)
                    {
                        string weld_description = Convert.ToString(dt_weld.Rows[i][dt_weld.Columns[spot_weld_col_Pipe_Ahd]]);
                        string pipe_description = dt_weld.Rows[i].GetChildRows(relation1)[0]["DESCRIPTION"].ToString();
                        string row_MMID = Convert.ToString(dt_weld.Rows[i][dt_weld.Columns[spot_weld_col_MM_Ahd]]);
                        string row_drag = Convert.ToString(dt_weld.Rows[i][dt_weld.Columns["DRAG"]]);

                        is_le = false;

                        if (dt_weld.Rows[i][col_le] != DBNull.Value && Convert.ToBoolean(dt_weld.Rows[i][col_le]) == true)
                        {
                            is_le = true;
                        }

                        #region Missing Info
                        string row_NG = "";
                        string row_Cutoff = "";
                        string row_Pt = "";

                        #endregion

                        #region Pipe & Bend File
                        string row_Pipe = "";
                        string row_Heat = "";
                        string row_Wallthk = "";
                        string row_Len = "";
                        string row_bend_ang = "";
                        string row_bend_type = "";



                        if (dt_weld.Rows[i].GetChildRows(relation1)[0][spot_pipe_col_PipeNo] != DBNull.Value)
                        {
                            row_Pipe = dt_weld.Rows[i].GetChildRows(relation1)[0][spot_pipe_col_PipeNo].ToString();
                        }

                        if (dt_weld.Rows[i].GetChildRows(relation1)[0]["HEAT"] != DBNull.Value)
                        {
                            row_Heat = dt_weld.Rows[i].GetChildRows(relation1)[0]["HEAT"].ToString();
                        }

                        if (dt_weld.Rows[i].GetChildRows(relation1)[0]["WALL THICKNESS"] != DBNull.Value)
                        {
                            row_Wallthk = dt_weld.Rows[i].GetChildRows(relation1)[0]["WALL THICKNESS"].ToString();
                        }

                        if (dt_weld.Rows[i].GetChildRows(relation1)[0]["MEASURED LENGTH"] != DBNull.Value)
                        {
                            row_Len = dt_weld.Rows[i].GetChildRows(relation1)[0]["MEASURED LENGTH"].ToString();
                        }

                        if (dt_weld.Rows[i].GetChildRows(relation1)[0][spot_bend_col_ang] != DBNull.Value)
                        {
                            row_bend_ang = dt_weld.Rows[i].GetChildRows(relation1)[0][spot_bend_col_ang].ToString();
                        }

                        if (dt_weld.Rows[i].GetChildRows(relation1)[0][spot_bend_col_type] != DBNull.Value)
                        {
                            row_bend_type = dt_weld.Rows[i].GetChildRows(relation1)[0][spot_bend_col_type].ToString();
                        }

                        #endregion

                        #region Weld File

                        //string row_Wallthk = Convert.ToString(dt_weld.Rows[i][dt_weld.Columns["AHEAD PIPE WALL"]]);

                        string row_Weld = "";

                        if (dt_weld.Rows[i][spot_weld_col_Point_No] != DBNull.Value)

                        {
                            row_Weld = Convert.ToString(dt_weld.Rows[i][spot_weld_col_Point_No]);
                        }


                        double row_STA = 0;

                        if (dt_weld.Rows[i][spot_weld_col_STA] != DBNull.Value &&
                            Functions.IsNumeric(Convert.ToString(dt_weld.Rows[i][spot_weld_col_STA])) == true)
                        {
                            row_STA = Convert.ToDouble(dt_weld.Rows[i][spot_weld_col_STA]);
                        }

                        #endregion

                        if (weld_description.ToUpper() == pipe_description.ToUpper())
                        {
                            dt_compiled.Rows.Add();
                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][col_NG] = row_NG;
                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][col_Cutoff] = row_Cutoff;
                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][col_Pt] = row_Pt;
                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][col_MMID] = row_MMID;
                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][col_Pipe] = row_Pipe;
                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][col_Heat] = row_Heat;
                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][col_Weld] = row_Weld;
                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][col_Wallthk] = row_Wallthk;
                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][col_Len] = row_Len;
                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][col_drag] = row_drag;
                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][spot_bend_col_ang] = row_bend_ang;
                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][spot_bend_col_type] = row_bend_type;



                            if (i == 0 || is_le == true)
                            {
                                dt_compiled.Rows[dt_compiled.Rows.Count - 1][col_le] = true;
                                is_le = false;
                            }
                            dt_compiled.Rows[dt_compiled.Rows.Count - 1][col_STA] = Functions.Get_chainage_from_double(row_STA, "f", 0);

                        }

                    }

                    else
                    {
                        //string pipe_ahd = Convert.ToString(dt_weld.Rows[i][col_pipe_ahd]);
                        //DT_Missing_From_Spot.Rows.Add();
                        //DT_Missing_From_Spot.Rows[DT_Missing_From_Spot.Rows.Count - 1][col_Pipe] = pipe_ahd;
                        //DT_Missing_From_Spot.Rows[DT_Missing_From_Spot.Rows.Count - 1][col_source] = "Weld Data";
                        //DT_Missing_From_Spot.Rows[DT_Missing_From_Spot.Rows.Count - 1][col_missing] = "Pipe or Bend Data";
                    }
                }

                dataset1.Relations.Remove(relation1);
                dataset1.Tables.Remove(dt_weld);
                dataset1.Tables.Remove(dt_pipe_and_bend);



                System.Data.DataRow row1 = dt_compiled.NewRow();
                row1["WELD"] = "LE";
                row1["DRAG"] = Convert.ToInt32(dt_compiled.Rows[dt_compiled.Rows.Count - 1]["DRAG"]);
                dt_compiled.Rows.InsertAt(row1, dt_compiled.Rows.Count);

                for (int i = dt_compiled.Rows.Count - 1; i > 0; --i)
                {
                    row1 = dt_compiled.NewRow();
                    row1["WELD"] = "LE";
                    row1["DRAG"] = Convert.ToInt32(dt_compiled.Rows[i - 1]["DRAG"]);

                    int drag_cur = Convert.ToInt32(dt_compiled.Rows[i]["DRAG"]);
                    int drag_next = Convert.ToInt32(dt_compiled.Rows[i - 1]["DRAG"]);


                    if (drag_cur == drag_next)
                    {
                        //dt_compiled.Rows.Add();
                        //dt_compiled.Rows[dt_compiled.Rows.Count - 1].ItemArray = dt_compiled.Rows[i].ItemArray;

                    }
                    else
                    {
                        dt_compiled.Rows.InsertAt(row1, i);
                    }

                }

                #endregion


            }

            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

        }



        public void Add_Field_Note_DataRow(Worksheet W1, double Row_Count, int page_no, System.Data.DataTable dt1)
        {
            int Row_start = 58;
            int Increment = -9;

            int start1 = 6 * (page_no - 1);
            int end1 = start1 + 6;
            if (end1 > Row_Count) end1 = Convert.ToInt32(Row_Count);
            int j = 0;

            for (int i = start1; i < end1; ++i)
            {



                int Row_Index = Row_start + Increment * j;
                W1.Range["E" + (Row_Index + 1).ToString()].Value = dt1.Rows[i][col_NG];
                W1.Range["E" + (Row_Index + 3).ToString()].Value = dt1.Rows[i][col_Cutoff];
                W1.Range["E" + (Row_Index + 7).ToString()].Value = dt1.Rows[i][col_Pt];
                W1.Range["J" + (Row_Index + 1).ToString()].Value = dt1.Rows[i][col_MMID];
                W1.Range["J" + (Row_Index + 3).ToString()].Value = dt1.Rows[i][col_Pipe];
                W1.Range["J" + (Row_Index + 5).ToString()].Value = dt1.Rows[i][col_Heat];
                W1.Range["J" + (Row_Index + 7).ToString()].Value = dt1.Rows[i][col_Weld];
                W1.Range["P" + (Row_Index + 2).ToString()].Value = dt1.Rows[i][col_Wallthk];
                W1.Range["P" + (Row_Index + 4).ToString()].Value = dt1.Rows[i][col_Len];
                W1.Range["P" + (Row_Index + 6).ToString()].Value = dt1.Rows[i][col_STA];

                W1.Name = (page_no).ToString();

                ++j;
            }
        }


        private void button_Export_Field_Notes_Click(object sender, EventArgs e)
        {
            try
            {
                set_enable_false();
                import_data();

                //Common_Methods.Transfer_datatable_to_new_excel_spreadsheet(DT_Compiled);
                //return;

                if (dt_compiled != null && dt_compiled.Rows.Count > 0)
                {


                    Microsoft.Office.Interop.Excel.Application Excel1 = Functions.Excel_Open();
                    Workbook Workbook1 = Excel1.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
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

                    Worksheet W1 = Workbook1.Worksheets[1];

                    Page_Setup(W1);


                    Populate_data(Workbook1, W1, dt_compiled);

                    W1.Delete();
                    Excel1.DisplayAlerts = true;

                    //Functions.Transfer_datatable_to_new_excel_spreadsheet(dt_compiled);

                    //if (DT_Missing_From_Spot != null)
                    //{
                    //    Functions.Transfer_datatable_to_new_excel_spreadsheet(DT_Missing_From_Spot);
                    //}

                    //int count1 = dt_weld.Rows.Count;
                    //int count2 = DT_Bend.Rows.Count;
                    //int count3 = dt_pipe.Rows.Count;

                    //MessageBox.Show("Weld Count = " + Convert.ToString(count1)+ "\n Bend Count = " + Convert.ToString(count2) + "\n Pipe Count = " + Convert.ToString(count3), "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    MessageBox.Show("Done", "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            set_enable_true();
        }


        public void Populate_data(Workbook Workbook1, Worksheet W1, System.Data.DataTable dt1)
        {
            Worksheet W2 = Workbook1.Worksheets[Workbook1.Worksheets.Count];
            W1.Copy(After: W2);
            W2 = Workbook1.Worksheets[Workbook1.Worksheets.Count];

            int row1 = 58;
            int incr_xl = 9;

            int magic_no = 6; //there is room for 6 rows per page
            int tab_name = 1;


            bool is_first_on_page = false;



            int j = 0;


            //Functions.Transfer_datatable_to_new_excel_spreadsheet(dt1);

            for (int i = 0; i < dt1.Rows.Count; ++i)
            {

                {

                    bool le = false;

                    if (dt1.Rows[i][col_le] != DBNull.Value)
                    {
                        le = Convert.ToBoolean(dt1.Rows[i][col_le]);
                        if (le == true)
                        {
                            is_first_on_page = true;
                        }
                        else
                        {
                            is_first_on_page = false;
                        }
                    }
                    else
                    {
                        is_first_on_page = false;
                    }

                }

                if (is_first_on_page == true || i == 0)
                {
                    j = 1;

                }

                if (j == magic_no)
                {
                    j = 0;
                }


                if (i == dt1.Rows.Count - 1)
                {
                    W2.Name = tab_name.ToString();
                }

                if (i > 0 && (j == 0 || is_first_on_page == true))
                {
                    is_first_on_page = false;
                    W2.Name = tab_name.ToString();

                    if (j == 0)
                    {
                        #region Next Page Weld Info
                        // you don't need i+1.... first you write on current page then add a new one so is the same index for next on previous and current on current page
                        string pt_next = "";

                        if (dt_compiled.Rows[i][col_Pt] != DBNull.Value)
                        {
                            pt_next = Convert.ToString(dt_compiled.Rows[i][col_Pt]);
                        }

                        string weld_next = "";

                        if (dt_compiled.Rows[i][col_Weld] != DBNull.Value)
                        {
                            weld_next = Convert.ToString(dt_compiled.Rows[i][col_Weld]);
                        }

                        string sta_next = "";

                        if (dt_compiled.Rows[i][col_STA] != DBNull.Value)
                        {
                            sta_next = Convert.ToString(dt_compiled.Rows[i][col_STA]);
                        }

                        W2.Range["E11"].Value = pt_next;
                        W2.Range["J11"].Value = weld_next;
                        W2.Range["P11"].Value = sta_next;

                        #endregion
                    }

                    W1.Copy(After: W2);
                    W2 = Workbook1.Worksheets[Workbook1.Worksheets.Count];
                    ++tab_name;
                    W2.Name = tab_name.ToString();
                }




                #region read values
                string pipe1 = "";
                if (dt1.Rows[i][col_Pipe] != DBNull.Value)
                {
                    pipe1 = Convert.ToString(dt1.Rows[i][col_Pipe]);
                }


                string weld1 = "";
                if (dt1.Rows[i][col_Weld] != DBNull.Value)
                {
                    weld1 = Convert.ToString(dt1.Rows[i][col_Weld]);
                }

                string sta1 = "";
                if (dt1.Rows[i][col_STA] != DBNull.Value)
                {
                    sta1 = Convert.ToString(dt1.Rows[i][col_STA]);
                }
                #endregion

                int idx_xl = row1 - incr_xl * j;//here you start from the bottom of the page

                #region write values
                W2.Range["E" + (idx_xl + 1).ToString()].Value = dt1.Rows[i][col_NG];
                W2.Range["E" + (idx_xl + 3).ToString()].Value = dt1.Rows[i][col_Cutoff];
                W2.Range["E" + (idx_xl + 7).ToString()].Value = dt1.Rows[i][col_Pt];
                W2.Range["J" + (idx_xl + 1).ToString()].Value = dt1.Rows[i][col_MMID];
                W2.Range["J" + (idx_xl + 3).ToString()].Value = pipe1;
                W2.Range["J" + (idx_xl + 5).ToString()].Value = dt1.Rows[i][col_Heat];
                W2.Range["J" + (idx_xl + 7).ToString()].Value = weld1;
                W2.Range["P" + (idx_xl + 2).ToString()].Value = dt1.Rows[i][col_Wallthk];
                W2.Range["P" + (idx_xl + 4).ToString()].Value = dt1.Rows[i][col_Len];
                W2.Range["P" + (idx_xl + 6).ToString()].Value = sta1;

                W2.Range["U" + (idx_xl + 4).ToString()].Value = Convert.ToString(dt1.Rows[i]["ANGLE"]) + " " + Convert.ToString(dt1.Rows[i]["BEND TYPE"]);
                #endregion


                ++j;



            }
        }


        #region drop downs
        private void comboBox_Pipe_DropDown(object sender, EventArgs e)
        {
            Functions.Load_opened_worksheets_to_combobox(comboBox_Pipe);
        }

        private void comboBox_Weld_DropDown(object sender, EventArgs e)
        {
            Functions.Load_opened_worksheets_to_combobox(comboBox_Weld);
        }

        private void comboBox_Bend_DropDown(object sender, EventArgs e)
        {
            Functions.Load_opened_worksheets_to_combobox(comboBox_Bend);
        }

        #endregion

        #region transfer to excel
        private void label2_Click(object sender, EventArgs e)
        {
            Functions.Transfer_datatable_to_new_excel_spreadsheet(dt_pipe);
        }

        private void label1_Click(object sender, EventArgs e)
        {
            Functions.Transfer_datatable_to_new_excel_spreadsheet(dt_weld);
        }

        private void label4_Click(object sender, EventArgs e)
        {
            Functions.Transfer_datatable_to_new_excel_spreadsheet(dt_compiled);
        }


        #endregion

        private void label_mm_Click(object sender, EventArgs e)
        {
            comboBox_Pipe.Items.Clear();
            comboBox_Bend.Items.Clear();
            comboBox_Weld.Items.Clear();

        }
    }
}
