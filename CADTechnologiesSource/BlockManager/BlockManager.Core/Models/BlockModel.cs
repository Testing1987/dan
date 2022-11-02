using Autodesk.AutoCAD.DatabaseServices;
using System.Collections.ObjectModel;

namespace BlockManager.Core.Models
{
    /// <summary>
    /// A model that represents an AutoCAD block and it's attributes.
    /// </summary>
    public class BlockModel
    {
        public ObservableCollection<AttributeModel> Attributes { get; set; }
        public string BlockName { get; set; }
        public Handle BlockHandle { get; set; }
        public ObjectId BlockObjectId { get; set; }
        public string HostDrawing { get; set; }
        public string StringBlockLocation { get; set; }
        public string BlockLocation { get; set; }
        public string CurrentSpace { get; set; }
        public string LayoutName { get; set; }

    }
}
