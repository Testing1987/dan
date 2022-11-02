using Autodesk.AutoCAD.Runtime;

namespace LineTypeBuilderCAD.Commands
{
    public class Launcher
    {
        [CommandMethod("LTBUILDER")]
        public void ShowLTBuilderMethod()
        {
            MainWindow mainWindowr = new MainWindow();
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessWindow(mainWindowr);
        }
    }
}
