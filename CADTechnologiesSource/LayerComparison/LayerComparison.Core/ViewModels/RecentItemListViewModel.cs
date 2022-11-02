using CADTechnologiesSource.All.Base;
using System.Collections.Generic;

namespace LayerComparison.Core.ViewModels
{
    /// <summary>
    /// A view model for the list of recent items.
    /// </summary>
    public class RecentItemListViewModel : BaseViewModel
    {
        /// <summary>
        /// The Recent Items for the list.
        /// </summary>
        public List<RecentItemViewModel> RecentItems { get; set; }
    }
}
