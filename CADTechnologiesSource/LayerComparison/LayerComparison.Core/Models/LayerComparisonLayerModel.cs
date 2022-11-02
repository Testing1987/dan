using CADTechnologiesSource.All.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LayerComparison.Core.Models
{
    /// <summary>
    /// An extention of a normal layer model that has additional properties for use in LCOMP.
    /// </summary>
    public class LayerComparisonLayerModel : LayerModel
    {
        public string SourceOnOff { get; set; }
        public string SourceFreeze { get; set; }
        public string SourceColor { get; set; }
        public string SourceLinetype { get; set; }
        public string SourceLineweight { get; set; }
        public string SourceTransparency { get; set; }
        public string SourcePlot { get; set; }
        public string MissingSourceLayer { get; set; }
    }
}
