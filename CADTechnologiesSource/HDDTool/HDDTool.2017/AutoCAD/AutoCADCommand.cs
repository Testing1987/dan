using Autodesk.AutoCAD.Runtime;
using HDDTool.UI.Views;

namespace HDDTool._2017.AutoCAD
{
    public class AutoCADCommand
    {
        [CommandMethod("HDDTool")]
        public void ShowHDDTool()
        {
            MainWindow_HDDTool mw_HDDTool = new MainWindow_HDDTool();
            Autodesk.AutoCAD.ApplicationServices.Core.Application.ShowModelessWindow(mw_HDDTool);
        }
    }
}
