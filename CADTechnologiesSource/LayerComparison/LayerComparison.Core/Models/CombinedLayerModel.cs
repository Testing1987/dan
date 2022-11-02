using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LayerComparison.Core.Models
{
    public class CombinedLayerModel 
    {
        public string SourceOnOff { get; set; }
        public string SourceFreeze { get; set; }
        public string SourceColor { get; set; }
        public string SourceLinetype { get; set; }
        public string SourceLineweight { get; set; }
        public string SourceTransparency { get; set; }
        public string SourcePlot { get; set; }
        public string MissingSourceLayer { get; set; }

        public string DrawingPath { get; set; }
        public string Name { get; set; }
        public string OnOff { get; set; }
        public string Freeze { get; set; }
        public string Color { get; set; }
        public string Linetype { get; set; }
        public string Lineweight { get; set; }
        public string Transparency { get; set; }
        public string Plot { get; set; }

        public string ViewportLayer { get; set; }
        public Point3d ViewportPosition { get; set; }
        public string ViewportFreeze { get; set; }
        public string ViewportColor { get; set; }
        public string ViewportLinetype { get; set; }
        public string ViewportLineweight { get; set; }
        public string ViewportTransparency { get; set; }

        public string SourceViewportFreeze { get; set; }
        public string SourceViewportColor { get; set; }
        public string SourceViewportLinetype { get; set; }
        public string SourceViewportLineweight { get; set; }
        public string SourceViewportTransparency { get; set; }

        public bool Equals(CombinedLayerModel other)
        {
            if (other == null)
                return false;
            return
            Name == other.Name
            && OnOff == other.OnOff
            && Freeze == other.Freeze
            && Color == other.Color
            && Linetype == other.Linetype
            && Lineweight == other.Lineweight
            && Linetype == other.Linetype
            && Transparency == other.Transparency
            && Plot == other.Plot
            && ViewportFreeze == other.ViewportFreeze
            && ViewportColor == other.ViewportColor
            && ViewportLinetype == other.ViewportLinetype
            && ViewportLineweight == other.ViewportLineweight
            && ViewportTransparency == other.ViewportTransparency;
        }
        public override int GetHashCode()
        {
            //If obj is null then return 0
            if (this == null)
            {
                return 0;
            }
            //Get the hash code values for each property
            int DrawingPathHashCode = DrawingPath == null ? 0 : DrawingPath.GetHashCode();
            int NameHashCode = Name == null ? 0 : Name.GetHashCode();
            int OnOffHashCode = OnOff == null ? 0 : OnOff.GetHashCode();
            int FreezeHashCode = Freeze == null ? 0 : Freeze.GetHashCode();
            int ColorHashCode = Color == null ? 0 : Color.GetHashCode();
            int LinetypeHashCode = Linetype == null ? 0 : Linetype.GetHashCode();
            int LineweightHashCode = Lineweight == null ? 0 : Lineweight.GetHashCode();
            int TransparencyHashCode = Transparency == null ? 0 : Transparency.GetHashCode();
            int PlotHashCode = Plot == null ? 0 : Plot.GetHashCode();
            int ViewportFreezeHashCode = ViewportFreeze == null ? 0 : ViewportFreeze.GetHashCode();
            int ViewportColorHashCode = ViewportColor == null ? 0 : ViewportColor.GetHashCode();
            int ViewportLinetypeHashCode = ViewportLinetype == null ? 0 : ViewportLinetype.GetHashCode();
            int ViewportLineweightHashCode = ViewportLineweight == null ? 0 : ViewportLineweight.GetHashCode();
            int ViewportTransparencyHashCode = ViewportTransparency == null ? 0 : ViewportTransparency.GetHashCode();

            return DrawingPathHashCode ^ NameHashCode ^ ViewportFreezeHashCode ^ ViewportColorHashCode ^ ViewportLinetypeHashCode ^ ViewportLineweightHashCode ^ ViewportTransparencyHashCode ^ OnOffHashCode ^ FreezeHashCode ^ ColorHashCode ^ LinetypeHashCode ^ LineweightHashCode ^ TransparencyHashCode ^ PlotHashCode;
        }
    }
}
