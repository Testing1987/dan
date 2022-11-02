using System.Windows;
using Autodesk.AutoCAD.Runtime;
using BlockManager.UI.Views;

namespace BlockManager.AutoCAD.Commands
{
    public class LaunchBM
    {
        [CommandMethod("BAM")]
        public void ShowBAMMethod()
        {
            bool showBETAmessage = false;
            if (showBETAmessage == true)
            {
                MessageBox.Show("Block Manager is still in the BETA testing phase. " +
                                                 "It is strongly recommended that you run this on " +
                                                 "'save-as' copies of your drawings then manually check each drawing " +
                                                 "with standard QAQC practices after submitting changes.",
                                                 "BETA Notice",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            }

            MainWindow_BM mainWindow_BM = new MainWindow_BM();
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessWindow(mainWindow_BM);
        }
    }
}
