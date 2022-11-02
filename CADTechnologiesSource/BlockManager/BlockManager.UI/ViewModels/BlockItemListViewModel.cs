using CADTechnologiesSource.All.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlockManager.UI.ViewModels
{
    public class BlockItemListViewModel : BaseViewModel
    {
        public List<BlockItemViewModel> BlockItems { get; set; }
    }
}
