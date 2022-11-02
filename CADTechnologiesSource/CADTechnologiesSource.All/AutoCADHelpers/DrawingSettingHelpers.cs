using Autodesk.AutoCAD.ApplicationServices;
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;

namespace CADTechnologiesSource.All.AutoCADHelpers
{
    public class DrawingSettingHelpers
    {
        public bool IsVisretainOn()
        {
            try
            {
                object visretainMode = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("VISRETAIN");
                if(visretainMode.ToString() == "1")
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return false;
        }
    }
}
