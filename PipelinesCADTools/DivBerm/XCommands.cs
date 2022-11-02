using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DivBerm
{
    public class Class1
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
        [CommandMethod("DivBerm")]
        public void Show_DivBerm_form()
        {
            if (isSECURE() == true)
            {
                foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
                {
                    if (Forma1 is Alignment_mdi.DivBerm)
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
                    Alignment_mdi.DivBerm forma2 = new Alignment_mdi.DivBerm();
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
