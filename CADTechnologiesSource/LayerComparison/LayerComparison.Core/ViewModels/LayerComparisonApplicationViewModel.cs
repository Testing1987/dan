using CADTechnologiesSource.All.Base;
using LayerComparison.Core.Enums;

namespace LayerComparison.Core.ViewModels
{
    public class LayerComparisonApplicationViewModel : BaseViewModel
    {
        public ApplicationPage CurrentPage { get; private set; } = ApplicationPage.StartPage;

        public void GoToPage(ApplicationPage page)
        {
            //set the current page
            CurrentPage = page;
        }
    }
}
