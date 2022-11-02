using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Alignment_mdi
{

    public partial class Lgen_mainform : Form
    {


        private bool clickdragdown;
        private Point lastLocation;

        public static Lgen_label_page tpage_label = new Lgen_label_page();
        public static Lgen_alias_page tpage_layer_alias = new Lgen_alias_page();

        public static System.Data.DataTable dt_alias = null;
        public static List<string> lista_layere = null;
        public static string Fisier_layer_alias = "";

        public static bool ExcelVisible = false;
        public static int Start_row_layer_alias = 8;

        public Lgen_mainform()
        {
            InitializeComponent();
            MenuStrip_lgen.BackColor = Color.FromArgb(37, 37, 38);
            MenuStrip_lgen.ForeColor = Color.FromArgb(0, 122, 204);
            MenuStrip_lgen.Renderer = new MyRenderer();

            foreach (Control ctrl in this.Controls)
            {
                if (ctrl is MdiClient)
                {
                    ctrl.BackColor = Color.FromArgb(62, 62, 66);
                }
            }
            tpage_layer_alias.Hide();

            tpage_label = new Lgen_label_page();
            tpage_label.MdiParent = this;
            tpage_label.Dock = DockStyle.Fill;
            tpage_label.Show();
            //Fisier_layer_alias = "G:\\PennEast\\353754_PennEast_Pipeline_EPCM\\DataProd\\_State_Permit\\_ProjectData\\Standards\\LGEN\\LGEN ALIAS TABLE.xlsx";



        }

        #region move and close
        private void clickmove_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            clickdragdown = true;
            lastLocation = e.Location;
        }

        private void clickmove_MouseMove(object sender, MouseEventArgs e)
        {
            if (clickdragdown == true)
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
            tpage_label.SET_VARIABLES_NULL();
            this.Close();
        }
        private void button_minimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }
        #endregion

        /// <summary>
        /// Hides the ugly MDI border around the child form
        /// </summary>
        /// <param name="e"></param>
        protected override void OnLoad(EventArgs e)
        {

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

        #region PageNavigation


        /// <summary>
        /// changes to the map layers child form and hides the mainpage
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void mapLayersToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tpage_label.Show();
            tpage_layer_alias.Hide();
        }

        private void setUpLabelStylesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tpage_label.Hide();

            tpage_layer_alias = new Lgen_alias_page();
            tpage_layer_alias.MdiParent = this;
            tpage_layer_alias.Dock = DockStyle.Fill;
            tpage_layer_alias.Show();

        }

        private void optionsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tpage_label.Hide();

            tpage_layer_alias = new Lgen_alias_page();
            tpage_layer_alias.MdiParent = this;
            tpage_layer_alias.Dock = DockStyle.Fill;
            tpage_layer_alias.Show();
        }

        #endregion

        private void Load_excel_alias_file_ToolStripMenuItem_Click(object sender, EventArgs e)
        {

            using (OpenFileDialog fbd = new OpenFileDialog())
            {
                fbd.Multiselect = false;
                fbd.Filter = "Excel files (*.xlsx)|*.xlsx";

                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    Fisier_layer_alias = fbd.FileName;
                    tpage_label.label_excel_file_green();
                    dt_alias = Load_existing_Lgen_layer_alias_from_excel(Fisier_layer_alias);
                }
                else
                {
                    tpage_label.label_excel_file_red();
                    Fisier_layer_alias = "";
                }
            }
        }
        public static System.Data.DataTable Load_existing_Lgen_layer_alias_from_excel(string File1)
        {

            if (System.IO.File.Exists(File1) == false)
            {
                MessageBox.Show("the layer alias data file does not exist");
                return null;
            }


            System.Data.DataTable dt1 = new System.Data.DataTable();

            try
            {
                Microsoft.Office.Interop.Excel.Application Excel1 = new Microsoft.Office.Interop.Excel.Application();
                if (Excel1 == null)
                {
                    MessageBox.Show("PROBLEM WITH EXCEL!");
                    return null;
                }


                Excel1.Visible = Lgen_mainform.ExcelVisible;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(File1);
                Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];

                try
                {
                    dt1 = Functions.Build_Lgen_Data_table_layer_alias_from_excel(W1, Lgen_mainform.Start_row_layer_alias + 1);
                    if (dt1.Rows.Count > 0)
                    {
                        lista_layere = new List<string>();
                        for (int i = 0; i < dt1.Rows.Count; ++i)
                        {
                            if (dt1.Rows[i][0] != DBNull.Value)
                            {
                                string layer1 = Convert.ToString(dt1.Rows[i][0]);
                                if (lista_layere.Contains(layer1) == false)
                                {
                                    lista_layere.Add(layer1);
                                }
                                else
                                {
                                    MessageBox.Show("the layer " + layer1 + "already exist in layer alias\r\nlayer not added to the layer alias");
                                }
                            }
                        }
                    }

                    Workbook1.Close();
                    Excel1.Quit();

                }
                catch (System.Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);

                }
                finally
                {
                    if (W1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W1);
                    if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                    if (Excel1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }
            return dt1;

        }
        private void open_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (System.IO.File.Exists(Fisier_layer_alias) == true)
            {
                Microsoft.Office.Interop.Excel.Application Excel1 = new Microsoft.Office.Interop.Excel.Application();
                if (Excel1 == null)
                {
                    MessageBox.Show("PROBLEM WITH EXCEL!");
                    return;
                }
                Excel1.Visible = true;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(Fisier_layer_alias);

            }
        }

        private void save_ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void saveAs_ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void Lgen_mainform_SizeChanged(object sender, EventArgs e)
        {
            button_Exit.Location = new Point(this.Width - 28, button_Exit.Location.Y);
            button_minimize.Location = new Point(this.Width - 59, button_minimize.Location.Y);
        }
    }
}
