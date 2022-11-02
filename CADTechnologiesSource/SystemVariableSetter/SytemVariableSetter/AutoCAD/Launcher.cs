using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using SystemVariableSetter.Views;

namespace SystemVariableSetter.AutoCAD
{
    public class Launcher
    {
        [CommandMethod("QUICKSYS")]
        public void ShowSystemVariableSetterCommand()
        {
            MainWindow mW = new MainWindow();
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessWindow(mW);
        }
    }
}
