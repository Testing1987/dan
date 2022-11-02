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

namespace Workspace_band_Csharp
{


    public class Workspace_command_class
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
                default:
                    try
                    {
                        string UserDNS = Environment.GetEnvironmentVariable("USERDNSDOMAIN");
                        if (UserDNS == "HMMG.CC")
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

        [CommandMethod("ZZZ")]
        public void Show_workspace_form()
        {
            if (isSECURE() == true)
            {
                foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
                {
                    if (Forma1 is Workspace_band_form)
                    {
                        Forma1.Focus();
                        Forma1.WindowState = System.Windows.Forms.FormWindowState.Normal;
                        return;
                    }
                }

                try
                {
                    Workspace_band_form forma2 = new Workspace_band_form();
                    Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma2);
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


        [CommandMethod("PAS")]
        public void Point_at_station()
        {
            if (isSECURE() == true)
            {
                try
                {
                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                    Polyline Poly1 = null;
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                    using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {

                            Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                            Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                            Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect reference polyline:");
                            Prompt_centerline.SetRejectMessage("\nSelect a polyline!");
                            Prompt_centerline.AllowNone = true;
                            Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                            Rezultat_centerline = ThisDrawing.Editor.GetEntity(Prompt_centerline);


                            if (Rezultat_centerline.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                            {

                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }

                            Poly1 = (Polyline)Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForRead);


                        }

                    station_pick:
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {

                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);



                            Autodesk.AutoCAD.EditorInput.PromptStringOptions String1 = new Autodesk.AutoCAD.EditorInput.PromptStringOptions("\nSpecify station:");
                            Autodesk.AutoCAD.EditorInput.PromptResult Descriptia = Editor1.GetString(String1);
                            if (Descriptia.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                            {

                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }

                            String Ch_result = Descriptia.StringResult;
                            Ch_result = Ch_result.Replace("+", "");
                            Ch_result = Ch_result.Replace(" ", "");
                            if (Workspace_Band.Functions.IsNumeric(Ch_result) == false)
                            {
                                ThisDrawing.Editor.WriteMessage("\nStation is not specified correctly");
                                return;
                            }

                            double Station1 = Math.Abs(Convert.ToDouble(Ch_result));

                            if (Poly1.Length >= Station1)
                            {
                                Point3d Point_on_poly = Poly1.GetPointAtDist(Station1);
                                Workspace_Band.Functions.Mleader_Create_without_UCS_transform(Point_on_poly, Workspace_Band.Functions.Get_chainage_feet_from_double(Station1, 0), 2, 1, 1, 1, 10);
                            }
                            else
                            {
                                ThisDrawing.Editor.WriteMessage("\nStation is not specified correctly");
                            }
                            Trans1.Commit();
                        }
                        goto station_pick;

                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

            else
            {
                return;
            }
        }

        [CommandMethod("SAP")]
        public void Station_at_point()
        {
            if (isSECURE() == true)
            {
                try
                {
                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                    Polyline Poly1 = null;
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                    using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {

                            Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                            Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                            Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect reference polyline:");
                            Prompt_centerline.SetRejectMessage("\nSelect a polyline!");
                            Prompt_centerline.AllowNone = true;
                            Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                            Rezultat_centerline = ThisDrawing.Editor.GetEntity(Prompt_centerline);


                            if (Rezultat_centerline.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                            {

                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }

                            Poly1 = (Polyline)Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForRead);


                        }

                    station_pick:
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {

                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                            Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                            PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify point:");
                            PP1.AllowNone = false;
                            Point_res1 = Editor1.GetPoint(PP1);

                            if (Point_res1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                            {

                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }

                            Point3d Point_on_poly = Poly1.GetClosestPointTo(Point_res1.Value, Vector3d.ZAxis, false);
                            Double Station1 = Poly1.GetDistAtPoint(Point_on_poly);
                            Workspace_Band.Functions.Mleader_Create_without_UCS_transform(Point_on_poly, Workspace_Band.Functions.Get_chainage_feet_from_double(Station1, 0), 2, 1, 1, 1, 10);

                            Trans1.Commit();
                        }
                        goto station_pick;

                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

            else
            {
                return;
            }
        }


    }
}
