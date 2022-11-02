using CADTechnologiesSource.All.Models;

namespace LayerComparison.Core.Models
{
    public class LayerComparisonViewportLayerModel : ViewportLayerModel
    {
        public string SourceViewportFreeze { get; set; }
        public string SourceViewportColor { get; set; }
        public string SourceViewportLinetype { get; set; }
        public string SourceViewportLineweight { get; set; }
        public string SourceViewportTransparency { get; set; }
    }
}
