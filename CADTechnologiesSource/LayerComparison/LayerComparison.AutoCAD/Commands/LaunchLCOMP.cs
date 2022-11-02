using Autodesk.AutoCAD.Runtime;
using LayerComparison.Core.Enums;
using LayerComparison.Core.IoCContainer;
using LayerComparison.Core.ViewModels;
using LayerComparison.UI.Views;
using System.Windows;

namespace LayerComparison.AutoCAD.Commands
{
    public class LaunchLCOMP
    {
        [CommandMethod("LCOMP")]
        public void ShowLCOMPMethod()
        {
            bool showBETAmessage = false;
            if (showBETAmessage == true)
            {
                MessageBox.Show("Layer Comparison is still in the BETA testing phase. " +
                                                 "It is strongly recommended that you run your comparison on " +
                                                 "'save-as' copies of your drawing set, and manually check each drawing " +
                                                 "with standard QAQC practices after applying fixes.",
                                                 "BETA Notice",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            }

            IoC.Setup();
            MainWindow_LayerComparison mainWindowLayerComparison = new MainWindow_LayerComparison();
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessWindow(mainWindowLayerComparison);
        }
    }
}
