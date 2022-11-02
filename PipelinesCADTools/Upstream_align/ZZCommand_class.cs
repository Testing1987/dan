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





    

        [CommandMethod("TCAL")]
        public void Show_tcal_form()
        {
            if (isSECURE() == true)
            {
                foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
                {
                    if (Forma1 is Alignment_mdi.Tcal)
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
                    Alignment_mdi.Tcal forma2 = new Alignment_mdi.Tcal();
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




        [CommandMethod("alignslice")]
        public void slice_of_tunnel()
        {
            if (isSECURE() == true)
            {
                foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
                {
                    if (Forma1 is slicer_form)
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
                    slicer_form forma2 = new slicer_form();
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


        [CommandMethod("pgen")]
        public void show_profiles_form()
        {
            if (isSECURE() == true)
            {
                foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
                {
                    if (Forma1 is Pgen_mainform)
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
                    Pgen_mainform forma2 = new Pgen_mainform();
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

        [CommandMethod("layer_sync")]
        public void show_layer_sync_form()
        {
            if (isSECURE() == true)
            {
                foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
                {
                    if (Forma1 is Layer_sync_mainform)
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
                    Layer_sync_mainform forma2 = new Layer_sync_mainform();
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

        [CommandMethod("monica_layers")]
        public void create_layers_for_monica()
        {


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
                        List<string> list1 = new List<string>();
                        list1.Add("E_FEA_Bollard");
                        list1.Add("E_FEA_CattleGuard");
                        list1.Add("E_FEA_Concrete");
                        list1.Add("E_FEA_Culvert");
                        list1.Add("E_FEA_DrainTile");
                        list1.Add("E_FEA_Fence");
                        list1.Add("E_FEA_Gate");
                        list1.Add("E_FEA_RipRap");
                        list1.Add("E_FEA_Sign");
                        list1.Add("E_FEA_StmDrain");
                        list1.Add("E_FEA_Tank");
                        list1.Add("E_FEA_WallRetain");
                        list1.Add("E_FEA_WallRock");
                        list1.Add("E_PIPE_AGM");
                        list1.Add("E_PIPE_LineMarker");
                        list1.Add("E_PIPE_Valve");
                        list1.Add("E_PIPE_Anomaly");
                        list1.Add("E_PIPE_Vent");
                        list1.Add("E_PRP_ATWS");
                        list1.Add("E_PRP_CL_Access");
                        list1.Add("E_PRP_ConstLmt");
                        list1.Add("E_PRP_EdgeAccess");
                        list1.Add("E_PRP_HDD_Pnt");
                        list1.Add("E_PRP_PermEase");
                        list1.Add("E_PRP_Route");
                        list1.Add("E_PRP_Site");
                        list1.Add("E_PRP_SoilSample");
                        list1.Add("E_PRP_TWS");
                        list1.Add("E_TER_AgField");
                        list1.Add("E_TER_BrushLine");
                        list1.Add("E_TER_RockArea");
                        list1.Add("E_TER_ToeBank");
                        list1.Add("E_TER_ToeBerm");
                        list1.Add("E_TER_ToeLevee");
                        list1.Add("E_TER_ToeSlope");
                        list1.Add("E_TER_TopBank");
                        list1.Add("E_TER_TopBerm");
                        list1.Add("E_TER_TopLevee");
                        list1.Add("E_TER_TopSlope");
                        list1.Add("E_TER_Tree");
                        list1.Add("E_TER_TreeLine");
                        list1.Add("E_TER_Pasture");
                        list1.Add("E_TER_Landscape");
                        list1.Add("E_TRA_BallastToe");
                        list1.Add("E_TRA_BallastTop");
                        list1.Add("E_TRA_Bridge");
                        list1.Add("E_TRA_CL_Rail");
                        list1.Add("E_TRA_CL_Road");
                        list1.Add("E_TRA_Curb");
                        list1.Add("E_TRA_Driveway");
                        list1.Add("E_TRA_Guardrail");
                        list1.Add("E_TRA_Median");
                        list1.Add("E_TRA_MilePost");
                        list1.Add("E_TRA_ParkingLot");
                        list1.Add("E_TRA_RD_EdgeAsph");
                        list1.Add("E_TRA_RD_EdgeCaliche");
                        list1.Add("E_TRA_RD_EdgeConc");
                        list1.Add("E_TRA_RD_EdgeDirt");
                        list1.Add("E_TRA_RD_EdgeGrav");
                        list1.Add("E_TRA_RD_EdgeRock");
                        list1.Add("E_TRA_Sidewalk");
                        list1.Add("E_TRA_TopRail");
                        list1.Add("E_UTL_Cable");
                        list1.Add("E_UTL_CableBox");
                        list1.Add("E_UTL_Cleanout");
                        list1.Add("E_UTL_ElecBox");
                        list1.Add("E_UTL_Fiber");
                        list1.Add("E_UTL_ForeignPL");
                        list1.Add("E_UTL_GuyAnchor");
                        list1.Add("E_UTL_GuyWire");
                        list1.Add("E_UTL_Hydrant");
                        list1.Add("E_UTL_LightPole");
                        list1.Add("E_UTL_MH_Elec");
                        list1.Add("E_UTL_MH_Fiber");
                        list1.Add("E_UTL_MH_Sewer");
                        list1.Add("E_UTL_MH_Water");
                        list1.Add("E_UTL_MH_Storm");
                        list1.Add("E_UTL_Marker");
                        list1.Add("E_UTL_Meter");
                        list1.Add("E_UTL_Pedestal");
                        list1.Add("E_UTL_PowerOH");
                        list1.Add("E_UTL_PowerPole");
                        list1.Add("E_UTL_PowerTrans");
                        list1.Add("E_UTL_PowerUG");
                        list1.Add("E_UTL_SanSewer");
                        list1.Add("E_UTL_Septic_S");
                        list1.Add("E_UTL_SepticFld");
                        list1.Add("E_UTL_StmSewer");
                        list1.Add("E_UTL_Telephone");
                        list1.Add("E_UTL_Tower");
                        list1.Add("E_UTL_Unknown");
                        list1.Add("E_UTL_Valve");
                        list1.Add("E_UTL_Water");
                        list1.Add("E_UTL_Well");
                        list1.Add("E_UTL_Text");
                        list1.Add("E_TRA_Text");
                        list1.Add("E_FEA_Text");
                        list1.Add("E_ENV_Text");
                        list1.Add("E_TER_Text");
                        list1.Add("E_PRP_Text");
                        list1.Add("E_PIPE_Text");
                        list1.Add("E_ASB_Text");

                        for (int i = 0; i < list1.Count; ++i)
                        {
                            Functions.Creaza_layer(list1[i], 7, true);
                        }

                        Trans1.Commit();

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

        [CommandMethod("zz")]
        public void XDATA1()
        {
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



                        Autodesk.AutoCAD.EditorInput.PromptEntityResult result_1;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_1;
                        Prompt_1 = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the polyline:");
                        Prompt_1.SetRejectMessage("\nSelect a polyline!");
                        Prompt_1.AllowNone = true;
                        Prompt_1.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        result_1 = ThisDrawing.Editor.GetEntity(Prompt_1);

                        if (result_1.Status != PromptStatus.OK)
                        {
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }
                        Entity ent1 = Trans1.GetObject(result_1.ObjectId, OpenMode.ForRead) as Entity;
                        ObjectId ext_dict_id = ent1.ExtensionDictionary;

                        System.Data.DataTable dt1 = new System.Data.DataTable();
                        dt1.Columns.Add("0", typeof(string));
                        dt1.Columns.Add("1", typeof(string));

                        if (ext_dict_id != ObjectId.Null)
                        {
                            DBDictionary Ext_dict = Trans1.GetObject(ext_dict_id, OpenMode.ForRead) as DBDictionary;
                            DBDictionary DBDict = Trans1.GetObject(Ext_dict.GetAt("ESRI_ATTRIBUTES"), OpenMode.ForRead) as DBDictionary;
                            foreach (DBDictionaryEntry entry1 in DBDict)
                            {
                                ObjectId id1 = entry1.Value;

                                Xrecord xrec1 = Trans1.GetObject(id1, OpenMode.ForRead) as Xrecord;
                                ResultBuffer rb1 = xrec1.Data;
                                if (rb1 != null)
                                {
                                    foreach (TypedValue tv in rb1)
                                    {
                                        string val1 = tv.Value.ToString();
                                        dt1.Rows.Add();
                                        dt1.Rows[dt1.Rows.Count - 1]["0"] = entry1.Key.ToString();
                                        dt1.Rows[dt1.Rows.Count - 1]["1"] = val1;
                                    }
                                }
                            }

                        }














                        try
                        {

                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }







                        Functions.Transfer_datatable_to_new_excel_spreadsheet(dt1);







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
    }
}
