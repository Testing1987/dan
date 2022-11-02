using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Text;
using MLineBlockAttributeEditor.UI.Views;
using MLineBlockAttributeEditor.Core.IoCContainer;

namespace MLineBlockAttributeEditor.AutoCAD.Commands
{
    public class LaunchMLBED
    {
        [CommandMethod("QUICKATT")]
        public void ShowMLBEDMethod()
        {
            IoC.Setup();
            MainWindow_MLineBlockAttributeEditor mainWindowMLBED = new MainWindow_MLineBlockAttributeEditor();
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessWindow(mainWindowMLBED);
        }
    }
}
