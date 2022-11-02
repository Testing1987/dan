using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CLTools.UI
{
    public partial class AGEN_01C_CL_tools : Form
    {
        #region Colors
        Color lightBackgroundColor = Color.FromArgb(62, 62, 66);
        Color darkBackgroundColor = Color.FromArgb(37, 37, 38);
        Color MMblueFont = Color.FromArgb(0, 122, 204);
        Color fontWhite = Color.White;
        #endregion

        #region Selected Form
        public string selectedForm = "Centerline";
        #endregion

        #region Help Messages

        public string centerlineHelpMessage = "TransCanada centerlines are provided by a third party survey company in the form of an Excel spreadsheet. " +
            "The purpose of this part of the program is to read the database provided by the survey company to create the AGEN centerline file." +
            " This has to be done because TransCanada project chainage/stationing is based on Combined Scale Factor, which is a different way of measuring" +
            " the length of the line. Please make sure you have the survey file open in a current Excel workbook and you select that worksheet from the Excel Mapper on the left side of the screen." +
            "\r\n\r\nFirst map the rows where the actual centerline data begins and ends. " +
            "\r\n\r\nNext map the columns for northing, easting, elevation, chainage, and CSF. " +
            "\r\n\r\nFinally press the Generate Centerline button to build the centerline file.";

        public string pipesummaryHelpMessage = "The purpose of this tool is to build a chainage-to-chainage summary of all linear pipe materials. " +
            "This is accomplished by reading the engineering database. Please make sure you have the engineering database loaded in the Excel Mapper " +
            "on the left side of this window. The program is designed to read the database automatically, however, if the database format has changed " +
            "please check the 'Map Manually' button and map the columns yourself.";

        public string csfstationingHelpMessage = "These controls are used to calculate chainages along the given centerline." +
            "\r\n\r\nExcel Tab - These buttons will read an open Excel worksheet and calculate the CSF chainages or US stations from the given coordinates. Conversely, it can also be used to read" +
            " a list of stations/chainages from Excel to give you the coordinates. You must use the Excel mapper to use these controls. " +
            "This control can be used to calculate a lot of points very quickly." +
            "\r\n\r\nAutoCAD Tab - These buttons will allow you to select a point in model space and calcualte the CSF chainage or station for you." +
            "\r\n\r\nManual Tab - These buttons will allow you to type in a chainage or a point manually and it will calculate the opposite for you. This is useful to quickly grab a chainage or point " +
            "from an email or IM.";

        public string utilitiesHelpMessage = "There are various utilities controls on this page.\r\n\r\n" +
            "Scan Heavy Wall - Select lines in AutoCAD that represent the heavy wall layout. A spreadsheet will be produced with the start/end chainages of it." +
            "\r\n\r\nDraw Heavy Wall Linework - Heavy wall will be drawn from the active Excel worksheet (set on the Excel mapper) from coordinates mapped above the button." +
            "\r\n\r\nBand Checks - This tool will ask you to select your engineering bands in the banding basefile and check both the length property of the block against the stationing," +
            " and the order in which the blocks are inserted, to make sure all blocks are placed in order.";

        #endregion

        public string helpTitleCenterline = "Centerline Help";
        public string helpTitlePipeSummary = "Pipe Summary Help";
        public string helpTitleCSFStationing = "Stationing Calc Help";
        public string helpTitleUtilities = "Utilities Help";
        public AGEN_01C_CL_tools()
        {
            InitializeComponent();
            DataTable dt_centerline = new DataTable();
            dt_centerline.Columns.Add("Generated centerline data will appear here", typeof(string));

            dataGrid_centerline.DataSource = dt_centerline;
            dataGrid_centerline.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            dataGrid_centerline.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGrid_centerline.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGrid_centerline.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGrid_centerline.DefaultCellStyle.ForeColor = Color.White;
            dataGrid_centerline.EnableHeadersVisualStyles = false;

            dataGridView_pipe_summary.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            dataGridView_pipe_summary.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_pipe_summary.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView_pipe_summary.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_pipe_summary.DefaultCellStyle.ForeColor = Color.White;
            dataGridView_pipe_summary.EnableHeadersVisualStyles = false;
        }

        #region Buttons

        private void button_centerline_Click(object sender, EventArgs e)
        {
            #region Highlighting
            button_centerline.BackColor = darkBackgroundColor;
            button_pipe_summary.BackColor = lightBackgroundColor;
            button_CSF_stationing.BackColor = lightBackgroundColor;
            button_utilities.BackColor = lightBackgroundColor;

            button_centerline.ForeColor = MMblueFont;
            button_pipe_summary.ForeColor = fontWhite;
            button_CSF_stationing.ForeColor = fontWhite;
            button_utilities.ForeColor = fontWhite;
            #endregion
            panel_excel_mapper_controls.Visible = true;
            panel_centerline.BringToFront();
            selectedForm = "Centerline";
        }

        private void button_pipe_summary_Click(object sender, EventArgs e)
        {
            #region Highlighting
            button_centerline.BackColor = lightBackgroundColor;
            button_pipe_summary.BackColor = darkBackgroundColor;
            button_CSF_stationing.BackColor = lightBackgroundColor;
            button_utilities.BackColor = lightBackgroundColor;

            button_centerline.ForeColor = fontWhite;
            button_pipe_summary.ForeColor = MMblueFont;
            button_CSF_stationing.ForeColor = fontWhite;
            button_utilities.ForeColor = fontWhite;
            #endregion
            panel_excel_mapper_controls.Visible = false;
            panel_pipe_summary.BringToFront();
            selectedForm = "PipeSummary";
        }

        private void button_CSF_stationing_Click(object sender, EventArgs e)
        {
            #region Highlighting
            button_centerline.BackColor = lightBackgroundColor;
            button_pipe_summary.BackColor = lightBackgroundColor;
            button_CSF_stationing.BackColor = darkBackgroundColor;
            button_utilities.BackColor = lightBackgroundColor;

            button_centerline.ForeColor = fontWhite;
            button_pipe_summary.ForeColor = fontWhite;
            button_CSF_stationing.ForeColor = MMblueFont;
            button_utilities.ForeColor = fontWhite;
            #endregion
            panel_excel_mapper_controls.Visible = true;
            panel_csf_stationing.BringToFront();
            selectedForm = "CSFStationing";
        }

        private void button_US_stationing_Click(object sender, EventArgs e)
        {
            #region Highlighting
            button_centerline.BackColor = lightBackgroundColor;
            button_pipe_summary.BackColor = lightBackgroundColor;
            button_CSF_stationing.BackColor = lightBackgroundColor;
            button_utilities.BackColor = lightBackgroundColor;

            button_centerline.ForeColor = fontWhite;
            button_pipe_summary.ForeColor = fontWhite;
            button_CSF_stationing.ForeColor = fontWhite;
            button_utilities.ForeColor = fontWhite;
            #endregion
            panel_excel_mapper_controls.Visible = true;
            panel_us_stationing.BringToFront();
            selectedForm = "USStationing";
        }

        private void button_utilities_Click(object sender, EventArgs e)
        {
            #region Highlighting
            button_centerline.BackColor = lightBackgroundColor;
            button_pipe_summary.BackColor = lightBackgroundColor;
            button_CSF_stationing.BackColor = lightBackgroundColor;
            button_utilities.BackColor = darkBackgroundColor;

            button_centerline.ForeColor = fontWhite;
            button_pipe_summary.ForeColor = fontWhite;
            button_CSF_stationing.ForeColor = fontWhite;
            button_utilities.ForeColor = MMblueFont;
            #endregion
            panel_excel_mapper_controls.Visible = true;
            panel_utilities.BringToFront();
            selectedForm = "Utilities";
        }

        private void button_nav_station_exel_Click(object sender, EventArgs e)
        {
            #region Highlighting
            button_nav_station_exel.BackColor = darkBackgroundColor;
            button_nav_station_AutoCAD.BackColor = lightBackgroundColor;
            button_nav_station_Manual.BackColor = lightBackgroundColor;

            button_nav_station_exel.ForeColor = MMblueFont;
            button_nav_station_AutoCAD.ForeColor = fontWhite;
            button_nav_station_Manual.ForeColor = fontWhite;
            #endregion
            panel_excel_mapper_controls.Visible = true;
            panel_station_excel.Visible = true;
            panel_station_manual.Visible = false;
            panel_station_autocad.Visible = false;
        }

        private void button_nav_station_AutoCAD_Click(object sender, EventArgs e)
        {
            #region Highlighting
            button_nav_station_exel.BackColor = lightBackgroundColor;
            button_nav_station_AutoCAD.BackColor = darkBackgroundColor;
            button_nav_station_Manual.BackColor = lightBackgroundColor;

            button_nav_station_exel.ForeColor = fontWhite;
            button_nav_station_AutoCAD.ForeColor = MMblueFont;
            button_nav_station_Manual.ForeColor = fontWhite;
            #endregion
            panel_excel_mapper_controls.Visible = false;
            panel_station_manual.Visible = false;
            panel_station_excel.Visible = false;
            panel_station_autocad.Visible = true;
        }

        private void button_nav_station_Manual_Click(object sender, EventArgs e)
        {
            #region Highlighting
            button_nav_station_exel.BackColor = lightBackgroundColor;
            button_nav_station_AutoCAD.BackColor = lightBackgroundColor;
            button_nav_station_Manual.BackColor = darkBackgroundColor;

            button_nav_station_exel.ForeColor = fontWhite;
            button_nav_station_AutoCAD.ForeColor = fontWhite;
            button_nav_station_Manual.ForeColor = MMblueFont;
            #endregion
            panel_excel_mapper_controls.Visible = false;
            panel_station_autocad.Visible = false;
            panel_station_manual.Visible = true;
            panel_station_excel.Visible = false;
        }

        private void button_help_Click(object sender, EventArgs e)
        {
            switch (selectedForm)
            {
                case "Centerline":
                    Help_Form centerline_help_Form = new Help_Form(helpTitleCenterline, centerlineHelpMessage);
                    centerline_help_Form.Show();
                    break;

                case "PipeSummary":
                    Help_Form pipeSummary_help_Form = new Help_Form(helpTitlePipeSummary, pipesummaryHelpMessage);
                    pipeSummary_help_Form.Show();
                    break;

                case "CSFStationing":
                    Help_Form CSFStationing_help_Form = new Help_Form(helpTitleCSFStationing, csfstationingHelpMessage);
                    CSFStationing_help_Form.Show();
                    break;

                case "Utilities":
                    Help_Form Utilities_help_Form = new Help_Form(helpTitleUtilities, utilitiesHelpMessage);
                    Utilities_help_Form.Show();
                    break;

                default:
                    break;
            }
        }

        #endregion

        #region Radio Buttons

        private void radioButton_map_pipe_summary_auto_CheckedChanged(object sender, EventArgs e)
        {
            if(radioButton_map_pipe_summary_auto.Checked ==  true)
            {
                panel_pipe_summary_mapper.Visible = false;
            }
        }

        private void radioButton_map_pipe_summary_manual_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton_map_pipe_summary_manual.Checked == true)
            {
                panel_pipe_summary_mapper.Visible = true;
            }
        }

        #endregion

        #region DataTables



        #endregion

    }
}
