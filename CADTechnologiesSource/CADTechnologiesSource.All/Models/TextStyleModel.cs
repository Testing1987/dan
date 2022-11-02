using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.GraphicsInterface;
using System;
using System.Collections.Generic;
using System.Text;

namespace CADTechnologiesSource.All.Models
{
    public class TextStyleModel : IEquatable<TextStyleModel>
    {
        public string Name { get; set; }
        public  FontDescriptor Font { get; set; }
        public double TextSize { get; set; }
        public PaperOrientationStates PaperOrientation { get; set; }
        public double ObliquingAngle { get; set; }
        public bool IsVertical { get; set; }



        public bool Equals(TextStyleModel other)
        {
            if (other == null)
                return false;
            return
            Name == other.Name
            && Font == other.Font
            && TextSize == other.TextSize
            && PaperOrientation == other.PaperOrientation
            && ObliquingAngle == other.ObliquingAngle
            && IsVertical == other.IsVertical;
        }

        public override int GetHashCode()
        {
            //If obj is null then return 0
            if (this == null)
            {
                return 0;
            }
            //Get the hash code values for each property
            int NameHashCode = Name == null ? 0 : Name.GetHashCode();
            int FontHashCode = Font == null ? 0 : Font.GetHashCode();
            int TextSizeHashCode = TextSize.GetHashCode();
            int PaperOrientationHashCode = PaperOrientation.GetHashCode();
            int ObliquingAngleHashCode =  ObliquingAngle.GetHashCode();
            int IsVerticalHashCode = IsVertical.GetHashCode();

            return NameHashCode ^ FontHashCode ^ TextSizeHashCode ^ PaperOrientationHashCode ^ ObliquingAngleHashCode ^ IsVerticalHashCode;
        }
    }
}
