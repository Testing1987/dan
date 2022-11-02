using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Text;
using TagDupeFix.UI.Views;

namespace TagDupeFix.UI.Commands
{
    public class Launcher
    {
        [CommandMethod("TDF")]
        public void ShowTDF()
        {
             MainWindow mwTDF = new MainWindow();
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessWindow(mwTDF);
        }
    }
}
