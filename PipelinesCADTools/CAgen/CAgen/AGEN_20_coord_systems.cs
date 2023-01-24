using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace Alignment_mdi
{
    public partial class cs_form : Form
    {


        public cs_form()
        {
            InitializeComponent();
        }





        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();


            lista_butoane.Add(button_convert);
            lista_butoane.Add(button_load_cs);




            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = false;
            }
        }
        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();


            lista_butoane.Add(button_convert);
            lista_butoane.Add(button_load_cs);



            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }



        private void button_convert_coordinates_Click(object sender, EventArgs e)
        {
            int start1 = 0;


            if (Functions.IsNumeric(textBox_start.Text) == true)
            {
                start1 = Convert.ToInt32(textBox_start.Text);
            }


            if (start1 <= 0)
            {
                MessageBox.Show("specify the start row!");
                return;
            }
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            ObjectId[] Empty_array = null;

            if (comboBox_xl.Text != "" && textBox_X.Text != "" && textBox_Y.Text != "")
            {
                string string1 = comboBox_xl.Text;
                if (string1.Contains("[") == true && string1.Contains("]") == true)
                {
                    string filename = string1.Substring(string1.IndexOf("]") + 4, string1.Length - (string1.IndexOf("]") + 4));

                    string sheet_name = string1.Substring(1, string1.IndexOf("]") - 1);
                    if (filename.Length > 0 && sheet_name.Length > 0)
                    {
                        set_enable_false();
                        Microsoft.Office.Interop.Excel.Worksheet W1 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, sheet_name);
                        if (W1 != null)
                        {
                            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                            try
                            {
                                set_enable_false();
                                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                                {
                                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                                    {
                                        System.Data.DataTable dt_coords = Build_Data_table_coordinates_from_excel(W1, start1);
                                        if (dt_coords.Rows.Count > 0)
                                        {
                                            for (int i = 0; i < dt_coords.Rows.Count; i++)
                                            {
                                                double x1 = Convert.ToDouble(dt_coords.Rows[i][0]);
                                                double y1 = Convert.ToDouble(dt_coords.Rows[i][1]);
                                                Point3d point_dest = Functions.Convert_coordinate_from_CS_to_new_CS(new Point3d(x1, y1, 0), comboBox_from.Text, comboBox_to.Text);
                                                dt_coords.Rows[i][3] = point_dest.X;
                                                dt_coords.Rows[i][4] = point_dest.Y;
                                                dt_coords.Rows[i][5] = comboBox_to.Text;
                                            }
                                            Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt_coords);
                                        }
                                        Trans1.Commit();
                                        MessageBox.Show("done");
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
            }




            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            set_enable_true();

        }

        public static System.Data.DataTable Creaza_dt_coordinates_datatable_structure()
        {

            System.Data.DataTable dt1 = new System.Data.DataTable();
            dt1.Columns.Add("Source\r\nX", typeof(double));
            dt1.Columns.Add("Source\r\nY", typeof(double));
            dt1.Columns.Add("Source\r\nCS", typeof(string));
            dt1.Columns.Add("Destination\r\nX", typeof(double));
            dt1.Columns.Add("Destination\r\nY", typeof(double));
            dt1.Columns.Add("Destination\r\nCS", typeof(string));

            return dt1;
        }

        private System.Data.DataTable Build_Data_table_coordinates_from_excel(Microsoft.Office.Interop.Excel.Worksheet W1, int Start_row)
        {

            System.Data.DataTable dt1 = Creaza_dt_coordinates_datatable_structure();

            string Col1 = textBox_X.Text.ToUpper();
            string Col2 = textBox_Y.Text.ToUpper();


            Range range1 = W1.Range[Col1 + Start_row.ToString() + ":" + Col1 + "1000000"];
            object[,] values1 = new object[1000000, 1];
            values1 = range1.Value2;

            Range range2 = W1.Range[Col2 + Start_row.ToString() + ":" + Col2 + "1000000"];
            object[,] values2 = new object[1000000, 1];
            values2 = range2.Value2;




            for (int i = 1; i <= values1.Length; ++i)
            {
                object Valoare1 = values1[i, 1];
                object Valoare2 = values2[i, 1];

                if (Valoare1 != null && Valoare2 != null)
                {
                    if (Functions.IsNumeric(Convert.ToString(Valoare1)) == true && Functions.IsNumeric(Convert.ToString(Valoare2)) == true)
                    {
                        dt1.Rows.Add();
                        dt1.Rows[dt1.Rows.Count - 1][0] = Valoare1;
                        dt1.Rows[dt1.Rows.Count - 1][1] = Valoare2;
                        dt1.Rows[dt1.Rows.Count - 1][2] = comboBox_from.Text;
                    }
                    else
                    {
                        i = values1.Length + 1;
                    }
                }
                else
                {
                    i = values1.Length + 1;
                }
            }
            return dt1;

        }

        private void comboBox_xl_DropDown(object sender, EventArgs e)
        {
            Functions.Load_opened_worksheets_to_combobox(comboBox_xl);
        }

        private void button_load_cs_Click(object sender, EventArgs e)
        {

            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {
                set_enable_false();
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        OSGeo.MapGuide.MgCoordinateSystemFactory Coord_factory1 = new OSGeo.MapGuide.MgCoordinateSystemFactory();
                        OSGeo.MapGuide.MgCoordinateSystemCatalog Catalog1 = Coord_factory1.GetCatalog();
                        OSGeo.MapGuide.MgCoordinateSystemDictionary Dictionary1 = Catalog1.GetCoordinateSystemDictionary();
                        OSGeo.MapGuide.MgCoordinateSystemEnum Enum1 = Dictionary1.GetEnum();
                        int count1 = Dictionary1.GetSize();
                        OSGeo.MapGuide.MgStringCollection Colectie_names_cs = Enum1.NextName(count1);

                        for (int k = 0; k < count1; k++)
                        {
                            string coord_sys_name = Colectie_names_cs.GetItem(k);

                            OSGeo.MapGuide.MgCoordinateSystem coord_sys = Dictionary1.GetCoordinateSystem(coord_sys_name);
                            string code1 = coord_sys.GetCsCode();

                            comboBox_from.Items.Add(code1);
                            comboBox_to.Items.Add(code1);
                        }

                        Autodesk.Gis.Map.Platform.AcMapMap Acmap = Autodesk.Gis.Map.Platform.AcMapMap.GetCurrentMap();
                        string string_current = Acmap.GetMapSRS();
                        if (string.IsNullOrEmpty(string_current) == false)
                        {
                            OSGeo.MapGuide.MgCoordinateSystem current_system = Coord_factory1.Create(string_current);
                            string code1 = current_system.GetCsCode();
                            if (comboBox_to.Items.Contains(code1) == true)
                            {
                                comboBox_from.SelectedIndex = comboBox_to.Items.IndexOf(code1);
                                comboBox_to.SelectedIndex = comboBox_from.Items.IndexOf(code1);
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            set_enable_true();

        }
    }
}
