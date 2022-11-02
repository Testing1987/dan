using CADTechnologiesSource.All.Base;
using System.Windows.Input;

namespace LayerComparison.Core.ViewModels
{
    /// <summary>
    /// A view model for each Recent Item shown on the Recent Items List.
    /// </summary>
    public class RecentItemViewModel : BaseViewModel
    {
        /// <summary>
        /// The filename of the .LCOMP file.
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// The saved path of the .LCOMP file.
        /// </summary>
        public string Path { get; set; }

        /// <summary>
        /// The command bound to the recent item button.
        /// </summary>
        public ICommand Command { get; set; }
    }
}
