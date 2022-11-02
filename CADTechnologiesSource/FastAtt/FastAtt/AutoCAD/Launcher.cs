using Autodesk.AutoCAD.Runtime;
using FastAtt.Views;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FastAtt.AutoCAD
{
    public class Launcher
    {
        [CommandMethod("FASTATT")]
        public void ShowFastAttMethod()
        {
            MainWindow_FastAtt mW = new MainWindow_FastAtt();
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessWindow(mW);
        }
    }
}
