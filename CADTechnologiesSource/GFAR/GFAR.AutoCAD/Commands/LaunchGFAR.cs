using Autodesk.AutoCAD.Runtime;
using GFAR.Core.IoCContainer;
using GFAR.UI.Views;

namespace GFAR.AutoCAD.Commands
{
    public class LaunchGFAR
    {
        [CommandMethod("GFAR")]
        public void ShowLCOMPMethod()
        {
            IoC.Setup();
            MainWindow_GFAR mainWindowGFAR = new MainWindow_GFAR();
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessWindow(mainWindowGFAR);
        }
    }
}
