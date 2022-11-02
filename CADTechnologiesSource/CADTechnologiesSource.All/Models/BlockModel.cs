using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Text;

namespace CADTechnologiesSource.All.Models
{
    /// <summary>
    /// A model that represents an AutoCAD block and it's attributes.
    /// </summary>
    public class BlockModel
    {
        public ObservableCollection<BlockAttributeModel> Attributes { get; set; }
        public string BlockName { get; set; }
    }
}
