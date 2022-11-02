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
    public partial class Wgen_extra_checks : Form
    {
        static public int X = 0;
        static public int Y = 0;
        public Wgen_extra_checks()
        {
            this.StartPosition = FormStartPosition.Manual;
            this.Location = new Point(X, Y);
            InitializeComponent();
            if (Wgen_main_form.check_dj_against_xray == true) checkBox_dj_vs_x_ray.Checked = true;

        }

        private void checkBox_dj_vs_x_ray_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_dj_vs_x_ray.Checked == true)
            {
                Wgen_main_form.check_dj_against_xray = true;
            }
            else
            {
                Wgen_main_form.check_dj_against_xray = false;
            }

            bool add_extra_check = true;
            for (int i = 0; i < Wgen_main_form.dt_feature_codes.Rows.Count; ++i)
            {
                if (Wgen_main_form.dt_feature_codes.Rows[i][0] != DBNull.Value)
                {
                    string client1 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[i][0]);

                    if (client1 == Wgen_main_form.lista_clienti[0])
                    {
                        if (add_extra_check == true)
                        {
                            if (checkBox_dj_vs_x_ray.Checked == true)
                            {
                                Wgen_main_form.dt_feature_codes.Rows[i]["CHECK DJ AGAINST XRAY"] = "YES";
                            }
                            else
                            {
                                Wgen_main_form.dt_feature_codes.Rows[i]["CHECK DJ AGAINST XRAY"] = "NO";

                            }
                            add_extra_check = false;
                        }

                    }

                }
            }

        }



        private void button_x_Click(object sender, EventArgs e)
        {
            checkBox_dj_vs_x_ray_CheckedChanged(sender, e);

            string file1 = Wgen_main_form.WGEN_folder + "wgen_feature_codes.xlsx";
            bool visible1 = false;
            if (Functions.is_dan_popescu() == true) visible1 = true;

            int nr_excel_open = Functions.Get_no_of_workbooks_from_Excel();

            if (System.IO.File.Exists(file1) == true)
            {
                Microsoft.Office.Interop.Excel.Application Excel1 = null;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;
                try
                {
                    try
                    {
                        Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    }
                    catch (System.Exception ex)
                    {
                        Excel1 = new Microsoft.Office.Interop.Excel.Application();
                    }
                    Excel1.Visible = visible1;
                    Workbook1 = Excel1.Workbooks.Open(file1);
                    Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];
                    if (checkBox_dj_vs_x_ray.Checked == true)
                    {
                        W1.Range["V2"].Value2 = "YES";
                    }
                    else
                    {
                        W1.Range["V2"].Value2 = "NO";
                    }

                    Workbook1.Save();
                    Workbook1.Close();
                    if (nr_excel_open == 0) Excel1.Quit();
                }
                catch (System.Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);
                }
                finally
                {
                    if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                    if (nr_excel_open == 0 && Excel1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }
            }
            this.Close();
        }
    }
}
