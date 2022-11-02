using Autodesk.AutoCAD.Runtime;
using BlockDeleter.Views;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlockDeleter.Base
{
    public class Launcher
    {
        [CommandMethod("BLOCK_DELETER")]
        public void ShowBlockDeleterMethod()
        {
            MainWindow mW = new MainWindow();
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessWindow(mW);
        }
    }
}
