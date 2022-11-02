using Autodesk.AutoCAD.Runtime;
using LayerManager.Views;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LayerManager.Base
{
    public class Launcher
    {
        [CommandMethod("QUICKLAYER")]
        public void ShowLayerManagerMethod()
        {
            MainWindow mW = new MainWindow();
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessWindow(mW);
        }
    }
}
