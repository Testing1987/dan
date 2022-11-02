using CADTechnologiesSource.All.Base;
using LayerComparison.Core.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LayerComparison.Core.ViewModels
{
    public class LayerItemListViewModel : BaseViewModel
    {
        /// <summary>
        /// The Recent Items for the list.
        /// </summary>
        public List<LayerItemViewModel> LayerItems { get; set; }
    }
}
