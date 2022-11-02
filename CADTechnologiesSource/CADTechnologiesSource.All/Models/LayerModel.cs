using System;

namespace CADTechnologiesSource.All.Models
{
    /// <summary>
    /// An object model that represents an AutoCAD Layer
    /// </summary>
    public class LayerModel : IEquatable<LayerModel>
    {
        public string DrawingPath { get; set; }
        public string Name { get; set; }
        public string OnOff { get; set; }
        public string Freeze { get; set; }
        public string Color { get; set; }
        public string Linetype { get; set; }
        public string Lineweight { get; set; }
        public string Transparency { get; set; }
        public string Plot { get; set; }

        public bool Equals(LayerModel other)
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
            && Transparency == other.Transparency
            && Plot == other.Plot;
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

            return DrawingPathHashCode ^ NameHashCode  ^ OnOffHashCode ^ FreezeHashCode ^ ColorHashCode ^ LinetypeHashCode ^ LineweightHashCode ^ TransparencyHashCode ^ PlotHashCode;
        }
    }
}
