using CADTechnologiesSource.All.Base;
using LayerComparison.Core.ViewModels;
using System.Windows;

namespace LayerComparison.Core.DesignModels
{
    /// <summary>
    /// A design model used for binding inside the designer to show dummy data during design-time. Do not bind to this for run-time.
    /// </summary>
    public class RecentItemDesignModel : RecentItemViewModel
    {
        #region Singleton
        /// <summary>
        /// A single instance of the design model of the RecentItem control to bind to for design purpsoes only.
        /// </summary>
        public static RecentItemDesignModel DesignModel => new RecentItemDesignModel();
        #endregion

        #region Constructor
        /// <summary>
        /// A constructor containing dummy data to show during design-time.
        /// </summary>
        public RecentItemDesignModel()
        {
            FileName = "Alignment Sheets IFC PangburnLateral.lcomp";
            Path = "C\\:Users\\RickyRaccoon\\AppData\\LCOMP\\PangburnLateral\\";
            Command = new RelayCommand(() => BoundCommand());
        }

        #region Command Methods
        /// <summary>
        /// The Method bound to the Command of the RecentItem
        /// </summary>
        private void BoundCommand()
        {
            MessageBox.Show("You Pressed the Command Button");
        }
        #endregion

        #endregion
    }
}
