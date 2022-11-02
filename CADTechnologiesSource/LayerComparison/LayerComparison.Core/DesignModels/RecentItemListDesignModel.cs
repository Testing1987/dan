using CADTechnologiesSource.All.Base;
using LayerComparison.Core.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace LayerComparison.Core.DesignModels
{
    /// <summary>
    /// A design model used for binding inside the designer to show dummy data during design-time. Do not bind to this for run-time.
    /// </summary>
    public class RecentItemListDesignModel : RecentItemListViewModel
    {
        #region Singleton
        /// <summary>
        /// A single instance of the design model of the RecentItemList control to bind to for design purpsoes only.
        /// </summary>
        public static RecentItemListDesignModel DesignModel => new RecentItemListDesignModel();
        #endregion

        public ICommand TestCommand { get; set; }

        #region Constructor
        /// <summary>
        /// A constructor containing dummy data to show during design-time.
        /// </summary>
        public RecentItemListDesignModel()
        {
            TestCommand = new RelayCommand(() => ExecuteCommand());

            RecentItems = new List<RecentItemViewModel>
            {
                new RecentItemViewModel
                {
                    Path = "C:\\Users\\RickyRaccoon\\AppData\\LCOMP\\PangburnLateral",
                    FileName = "AlignmentSheets-IFC-20200925.lcomp",
                    Command = TestCommand
                },
                new RecentItemViewModel
                {
                    Path = "C:\\Users\\RickyRaccoon\\AppData\\LCOMP\\SmokeyTrail",
                    FileName = "RoadCrossings-testcom__19840115.lcomp",
                    Command = TestCommand
                },
                new RecentItemViewModel
                {
                    Path = "C:\\Users\\RickyRaccoon\\AppData\\LCOMP\\GrasslandsSouth",
                    FileName = "asdfThingbadfilename_nogood.lcomp",
                    Command = TestCommand
                },
                new RecentItemViewModel
                {
                    Path = "C:\\Users\\RickyRaccoon\\AppData\\LCOMP\\GrasslandsSouth",
                    FileName = "asdfThingbadfilename_nogood.lcomp",
                    Command = TestCommand
                },
                new RecentItemViewModel
                {
                    Path = "C:\\Users\\RickyRaccoon\\AppData\\LCOMP\\PangburnLateral",
                    FileName = "AlignmentSheets-IFC-20200925.lcomp",
                    Command = TestCommand
                },
                new RecentItemViewModel
                {
                    Path = "C:\\Users\\RickyRaccoon\\AppData\\LCOMP\\SmokeyTrail",
                    FileName = "RoadCrossings-testcom__19840115.lcomp",
                    Command = TestCommand
                },
                new RecentItemViewModel
                {
                    Path = "C:\\Users\\RickyRaccoon\\AppData\\LCOMP\\GrasslandsSouth",
                    FileName = "asdfThingbadfilename_nogood.lcomp",
                    Command = TestCommand
                },
                new RecentItemViewModel
                {
                    Path = "C:\\Users\\RickyRaccoon\\AppData\\LCOMP\\GrasslandsSouth",
                    FileName = "asdfThingbadfilename_nogood.lcomp",
                    Command = TestCommand
                },
                new RecentItemViewModel
                {
                    Path = "C:\\Users\\RickyRaccoon\\AppData\\LCOMP\\PangburnLateral",
                    FileName = "AlignmentSheets-IFC-20200925.lcomp",
                    Command = TestCommand
                },
                new RecentItemViewModel
                {
                    Path = "C:\\Users\\RickyRaccoon\\AppData\\LCOMP\\SmokeyTrail",
                    FileName = "RoadCrossings-testcom__19840115.lcomp",
                    Command = TestCommand
                },
                new RecentItemViewModel
                {
                    Path = "C:\\Users\\RickyRaccoon\\AppData\\LCOMP\\GrasslandsSouth",
                    FileName = "asdfThingbadfilename_nogood.lcomp",
                    Command = TestCommand
                },
                new RecentItemViewModel
                {
                    Path = "C:\\Users\\RickyRaccoon\\AppData\\LCOMP\\GrasslandsSouth",
                    FileName = "asdfThingbadfilename_nogood.lcomp",
                    Command = TestCommand
                },
            };
        }

        private void ExecuteCommand()
        {
            MessageBox.Show("You Pressed The Button.");
        }
        #endregion
    }
}
