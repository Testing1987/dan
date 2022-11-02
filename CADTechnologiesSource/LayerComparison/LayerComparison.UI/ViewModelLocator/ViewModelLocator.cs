using LayerComparison.Core.IoCContainer;

namespace LayerComparison.UI.ViewModelLocator
{
    public class ViewModelLocator
    {
        /// <summary>
        /// A singleton instance of the locator
        /// </summary>
        public static ViewModelLocator Instance { get; private set; } = new ViewModelLocator();

        /// <summary>
        /// The application view model
        /// </summary>
        public static Core.ViewModels.LayerComparisonApplicationViewModel LayerComparisonApplicationViewModel => IoC.Get<Core.ViewModels.LayerComparisonApplicationViewModel>();
    }
}
