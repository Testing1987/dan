using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using BlockManager.Core.Models;
using CADTechnologiesSource.All.Base;
using Microsoft.Win32.SafeHandles;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Runtime.InteropServices;
using System.Windows.Input;

namespace BlockManager.UI.ViewModels
{
    public class BlockItemViewModel : BaseViewModel
    {
        #region Singleton
        public static BlockItemViewModel Instance = new BlockItemViewModel();
        #endregion

        #region Public Properties
        public string BlockName { get; set; }
        public Handle BlockHandle { get; set; }
        public ObjectId BlockObjectId { get; set; }
        public string Drawing { get; set; }
        public string SavedDateTime { get; set; }
        public string BlockLocation { get; set; }
        public string LayoutName { get; set; }
        public string StringBlockLocation { get; set; }
        public string CurrentSpace { get; set; }
        public ObservableCollection<AttributeItemViewModel> AttributesList { get; set; }
        public ICommand RefreshCommand { get; set; }
        public ICommand UpdateCommand { get; set; }
        public ICommand OpenCommand { get; set; }
        public ICommand ZoomCommand { get; set; }
        public System.Windows.Visibility MyVisibility { get; set; }

        #endregion
    }
}
