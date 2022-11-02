using Autodesk.AutoCAD.Runtime;
using AttributeFinder.Views;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttributeFinder.Base
{
    public class Launcher
    {
        [CommandMethod("ATTRIBUTE_FINDER")]
        public void ShowFastAttMethod()
        {
            MainWindow_AttributeFinder mW = new MainWindow_AttributeFinder();
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessWindow(mW);
        }
    }
}
