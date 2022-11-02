using Autodesk.AutoCAD.Runtime;

namespace BoreExcavationBuilder.Commands
{
    public class Launcher
    {
        [CommandMethod("BHBUILDER")]
        public void ShowBGBCommand()
        {
            MainWindow mainWindow = new MainWindow();
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessWindow(mainWindow);
        }
    }
}
