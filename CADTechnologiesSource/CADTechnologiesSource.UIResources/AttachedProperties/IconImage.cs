using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace CADTechnologiesSource.UIResources.AttachedProperties
{
    public class IconImage : DependencyObject
    {
        public static readonly DependencyProperty IconProperty =
            DependencyProperty.RegisterAttached("FontAwesomeIcon", typeof(string), typeof(IconImage), new PropertyMetadata(default(string)));

        public static void SetIcon(UIElement element, string value)
        {
            element.SetValue(IconProperty, value);
        }

        public static string GetIcon(UIElement element)
        {
            return (string)element.GetValue(IconProperty);
        }
    }
}
