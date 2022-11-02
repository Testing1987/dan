using Autodesk.AutoCAD.Geometry;
using System;

namespace CADTechnologiesSource.All.Models
{
    public class ViewportLayerModel: IEquatable<ViewportLayerModel>
    {
        public string DrawingPath { get; set; }
        public string Name { get; set; }
        public string ViewportLayer { get; set; }
        public Point3d ViewportPosition { get; set; }
        public string ViewportFreeze { get; set; }
        public string ViewportColor { get; set; }
        public string ViewportLinetype { get; set; }
        public string ViewportLineweight { get; set; }
        public string ViewportTransparency { get; set; }

        public bool Equals(ViewportLayerModel other)
        {
            if (other == null)
                return false;
            return
            Name == other.Name
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
            int ViewportFreezeHashCode = ViewportFreeze == null ? 0 : ViewportFreeze.GetHashCode();
            int ViewportColorHashCode = ViewportColor == null ? 0 : ViewportColor.GetHashCode();
            int ViewportLinetypeHashCode = ViewportLinetype == null ? 0 : ViewportLinetype.GetHashCode();
            int ViewportLineweightHashCode = ViewportLineweight == null ? 0 : ViewportLineweight.GetHashCode();
            int ViewportTransparencyHashCode = ViewportTransparency == null ? 0 : ViewportTransparency.GetHashCode();

            return DrawingPathHashCode ^ NameHashCode ^ ViewportFreezeHashCode ^ ViewportColorHashCode ^ ViewportLinetypeHashCode ^ ViewportLineweightHashCode ^ ViewportTransparencyHashCode;
        }
    }
}
