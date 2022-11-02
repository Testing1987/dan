using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Alignment_mdi
{
    public class Command_class
    {


        public bool isSECURE()
        {

            string number_drive = GetHDDSerialNumber("C");

            switch (number_drive)
            {
                case "B63AD3F6":
                    return true;
                default:
                    try
                    {
                        string UserDNS = Environment.GetEnvironmentVariable("USERDNSDOMAIN");
                        if (UserDNS.ToUpper() == "HMMG.CC" || UserDNS.ToLower() == "mottmac.group.int")
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


        [CommandMethod("hdd1")]
        public void Show_HDD_mainForm()
        {
            if (isSECURE() == true)
            {



                foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
                {
                    if (Forma1 is Alignment_mdi.HDD_mainform)
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
                    Alignment_mdi.HDD_mainform forma2 = new Alignment_mdi.HDD_mainform();
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


        [CommandMethod("BH1")]
        public void Show_BH_mainForm()
        {
            if (isSECURE() == true)
            {



                foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
                {
                    if (Forma1 is Alignment_mdi.BH_mainform)
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
                    Alignment_mdi.BH_mainform forma2 = new Alignment_mdi.BH_mainform();
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








    }
}
