using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System.Management;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.Civil.DatabaseServices;

namespace Alignment_mdi
{
    public class ZZCommand_class
    {
        public bool isSECURE()
        {

            string number_drive = GetHDDSerialNumber("C");

            switch (number_drive)
            {
                case "8CDA6CE3":
                    return true;
                case "36D79DE5":
                    return true;
                case "FEA3192C":
                    return true;
                case "B454BD5B":
                    return true;
                case "6E40460D":
                    return true;
                case "0892E01D":
                    return true;
                case "4ED21ABF":
                    return true;
                case "56766C69":
                    return true;
                case "DA214366":
                    return true;
                case "3CF68AF2":
                    return true;
                case "389A2249":
                    return true;
                case "AED6B68E":
                    return true;
                case "8C040338":
                    return true;
                case "8CD08F48":
                    return true;
                case "0E26E402":
                    return true;
                case "4A123A50":
                    return true;

                case "98D9B617":
                    return true;
                case "B838FEB4":
                    return true;
                case "1AE1721C":
                    return true;
                case "CA9E6FFE":
                    return true;
                case "DE281128":
                    return true;
                case "FC7C4F1":
                    return true;
                case "B67EC134":
                    return true;
                case "E64DBF0A":
                    return true;
                case "561F1509":
                    return true;

                case "120E4B54":
                    return true;
                case "F6633173":
                    return true;
                case "40D6BDCB":
                    return true;
                case "18399D24":
                    return true;


                case "B63AD3F6":
                    return true;
                default:
                    try
                    {
                        string UserDNS = Environment.GetEnvironmentVariable("USERDNSDOMAIN");
                        if (UserDNS.ToUpper() == "HMMG.CC" | UserDNS.ToLower() == "mottmac.group.int")
                        {
                            return true;
                        }
                        else
                        {
                            return false;
                        }
                    }
                    catch (System.Exception ex)
                    {
                        return false;
                    }
            }
        }


        public string GetHDDSerialNumber(string drive)
        {
            //check to see if the user provided a drive letter
            //if not default it to "C"
            if (drive == "" || drive == null)
            {
                drive = "C";
            }
            //create our ManagementObject, passing it the drive letter to the
            //DevideID using WQL
            ManagementObject disk = new ManagementObject("Win32_LogicalDisk.DeviceID=\"" + drive + ":\"");
            //bind our management object
            disk.Get();
            //return the serial number
            return disk["VolumeSerialNumber"].ToString();
        }





        [CommandMethod("wgen")]
        public void Show_wgen_form()
        {
            if (isSECURE() == true)
            {
                foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
                {
                    if (Forma1 is Alignment_mdi.Wgen_main_form)
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
                    Alignment_mdi.Wgen_main_form forma2 = new Alignment_mdi.Wgen_main_form();
                    Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma2);
                    forma2.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - forma2.Width) / 2,
                         (Screen.PrimaryScreen.WorkingArea.Height - forma2.Height) / 2);
                }
                catch (System.Exception EX)
                {
                    MessageBox.Show(EX.Message);
                }


            }
            else
            {
                return;
            }
        }

        [CommandMethod("aec1")]
        public void read_points()
        {
            if (isSECURE() == true)
            {

                string col_y = "Northing(N)";
                string col_x = "Easting (E)";
                string col_z = "Elevation (Z)";
                string col_pn = "Point Number";
                string col_slopeN = "N Slope %";
                string col_descr = "Description (D)";
                string col_NORTH_PN = "North Point number";


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
                            BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;


                            Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rez.MessageForAdding = "\nSelect points";
                            Prompt_rez.SingleOnly = false;
                            Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                            if (Rezultat1.Status != PromptStatus.OK)
                            {

                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }

                            System.Data.DataTable dt1 = new System.Data.DataTable();
                            dt1.Columns.Add(col_pn, typeof(string));
                            dt1.Columns.Add(col_y, typeof(double));
                            dt1.Columns.Add(col_x, typeof(double));
                            dt1.Columns.Add(col_z, typeof(double));
                            dt1.Columns.Add(col_descr, typeof(string));
                            dt1.Columns.Add(col_slopeN, typeof(double));
                            dt1.Columns.Add(col_NORTH_PN, typeof(string));


                            for (int i = 0; i < Rezultat1.Value.Count; ++i)
                            {
                                CogoPoint cg1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as CogoPoint;
                                if (cg1 != null)
                                {
                                    dt1.Rows.Add();
                                    dt1.Rows[dt1.Rows.Count - 1][col_pn] = cg1.PointNumber;
                                    dt1.Rows[dt1.Rows.Count - 1][col_descr] = cg1.FullDescription;
                                    dt1.Rows[dt1.Rows.Count - 1][col_x] = Math.Round(cg1.Location.X,4);
                                    dt1.Rows[dt1.Rows.Count - 1][col_y] =Math.Round( cg1.Location.Y,4);
                                    dt1.Rows[dt1.Rows.Count - 1][col_z] = Math.Round(cg1.Elevation,4);


                                }


                            }

                            System.Data.DataTable dt2 = Functions.Sort_data_table(dt1,col_y);

                            for(int i = 0; i<dt2.Rows.Count;++i)
                            {
                                double z1 =Convert.ToDouble( dt2.Rows[i][col_z]);
                                double x1 =Convert.ToDouble( dt2.Rows[i][col_x]);
                                double y1 =Convert.ToDouble( dt2.Rows[i][col_y]);

                                for (int j = i+1; j < dt2.Rows.Count; ++j)
                                {
                                    double z2 = Convert.ToDouble(dt2.Rows[j][col_z]);
                                    double x2 = Convert.ToDouble(dt2.Rows[j][col_x]);
                                    double y2 = Convert.ToDouble(dt2.Rows[j][col_y]);
                                    string pn2 = Convert.ToString(dt2.Rows[j][col_pn]);

                                    if (z2!=z1 && x1==x2)
                                    {

                                        double Rise = z2 - z1;
                                        double Run = y2 - y1;
                                        double Slope = Math.Round(100 * Rise / Run, 4);

                                        dt2.Rows[i][col_slopeN] = Slope;
                                        dt2.Rows[i][col_NORTH_PN] = pn2;
                                        j = dt2.Rows.Count;
                                    }
                                }

                            }



                            Worksheet W1 =    Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt2);
                            W1.Name = Environment.UserName + " " + DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + " at " + DateTime.Now.Hour + "h" + DateTime.Now.Minute + "m";

                            dt1.Dispose();
                            dt2.Dispose();


                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                Editor1.SetImpliedSelection(Empty_array);
                Editor1.WriteMessage("\nCommand:");




            }
            else
            {
                return;
            }
        }

    }
}
