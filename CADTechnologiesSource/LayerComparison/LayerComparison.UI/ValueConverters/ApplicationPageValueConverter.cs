using CADTechnologiesSource.All.Base;
using System;
using System.Diagnostics;
using System.Globalization;

namespace LayerComparison.UI.ValueConverters
{
    /// <summary>
    /// converts the <see cref="ApplicationPage"/> from a property to a page.
    /// </summary>
    public class ApplicationPageValueConverter : BaseValueConverter<ApplicationPageValueConverter>
    {
        public override object Convert(object value, Type targetType = null, object parameter = null, CultureInfo culture = null)
        {
            // find the appropriate page
            switch ((Core.Enums.ApplicationPage)value)
            {
                case Core.Enums.ApplicationPage.StartPage:
                    return new Views.StartPage();

                case Core.Enums.ApplicationPage.MainPage:
                    return new Views.MainPage();

                case Core.Enums.ApplicationPage.AdvancedPage:
                    return new Views.AdvancedComparisonsPage();

                //case Core.LayerController.LayerControllerApplicationPage.ReviewPage:
                //    return new MottMacDonald.PipelineCADTools.LayerController.ReviewPage();

                //case ApplicationPage.WHATEVERENUM:
                //    return new WHATEVERPAGE();


                default:
                    Debugger.Break();
                    return null;
            }
        }

        public override object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
