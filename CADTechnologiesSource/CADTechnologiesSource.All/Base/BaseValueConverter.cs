using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Markup;

namespace CADTechnologiesSource.All.Base
{
    /// <summary>
    /// A base value converter that allows direct XAML usage
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public abstract class BaseValueConverter<T> : MarkupExtension, IValueConverter
        where T : class, new()
    {       /// <summary>
            /// Method to convert one type to another
            /// </summary>
            /// <param name="value"></param>
            /// <param name="targetType"></param>
            /// <param name="parameter"></param>
            /// <param name="culture"></param>
            /// <returns></returns>
        public abstract object Convert(object value, Type targetType, object parameter, CultureInfo culture);


        /// <summary>
        /// The method to convert a value back to it's source type
        /// </summary>
        /// <param name="value"></param>
        /// <param name="targetType"></param>
        /// <param name="parameter"></param>
        /// <param name="culture"></param>
        /// <returns></returns>
        public abstract object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture);

        private static T mConverter = null;

        /// <summary>
        /// provides a static of the value converter
        /// </summary>
        /// <param name="serviceProvider">The Service Provider</param>
        /// <returns></returns>
        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            return mConverter ?? (mConverter = new T());
        }
    }
}
