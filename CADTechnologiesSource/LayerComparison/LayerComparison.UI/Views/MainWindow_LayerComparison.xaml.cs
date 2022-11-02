using System.Windows;

namespace LayerComparison.UI.Views
{
    /// <summary>
    /// Interaction logic for MainWindow_LayerComparison.xaml
    /// </summary>
    public partial class MainWindow_LayerComparison : Window
    {
        public Core.ViewModels.LayerComparisonApplicationViewModel ApplicationViewModel => new Core.ViewModels.LayerComparisonApplicationViewModel();
        public MainWindow_LayerComparison()
        {
            InitializeComponent();
        }
    }
}
