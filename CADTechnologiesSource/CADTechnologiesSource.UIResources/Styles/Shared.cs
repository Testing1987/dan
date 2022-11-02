using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;

namespace CADTechnologiesSource.UIResources.Styles
{
    public class SharedResourceDictionary : ResourceDictionary
    {
        /// <summary>
        /// Internal cache of loaded dictionaries 
        /// </summary>
        public static Dictionary<Uri, ResourceDictionary> _sharedDictionaries =
            new Dictionary<Uri, ResourceDictionary>();

        /// <summary>
        /// Local member of the source uri
        /// </summary>
        private Uri _sourceUri;

        /// <summary>
        /// Gets or sets the uniform resource identifier (URI) to load resources from.
        /// </summary>
        public new Uri Source
        {
            get
            {
                if (IsInDesignMode)
                    return base.Source;
                return _sourceUri;
            }
            set
            {
                if (IsInDesignMode)
                {
                    try
                    {
                        _sourceUri = new Uri(value.OriginalString);
                    }
                    catch
                    {
                        // do nothing?
                    }

                    return;
                }

                try
                {
                    _sourceUri = new Uri(value.OriginalString);
                }
                catch
                {
                    // do nothing?
                }

                if (!_sharedDictionaries.ContainsKey(value))
                {
                    // If the dictionary is not yet loaded, load it by setting
                    // the source of the base class

                    base.Source = value;

                    // add it to the cache
                    _sharedDictionaries.Add(value, this);
                }
                else
                {
                    // If the dictionary is already loaded, get it from the cache
                    MergedDictionaries.Add(_sharedDictionaries[value]);
                }
            }
        }

        private static bool IsInDesignMode
        {
            get
            {
                return (bool)DependencyPropertyDescriptor.FromProperty(DesignerProperties.IsInDesignModeProperty,
                typeof(DependencyObject)).Metadata.DefaultValue;
            }
        }
    }
}
