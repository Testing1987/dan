using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Security.Permissions;
using System.Text;

namespace CADTechnologiesSource.All.Models
{
    /// <summary>
    /// A model representing an AutoCAD block attribute.
    /// </summary>
    public class BlockAttributeModel
    {
        public string BlockName { get; set; }
        public string Tag { get; set; }
        public string TextString { get; set; }
        public string MtextAttributeContent { get; set; }
        public string Layer { get; set; }
        public Color Color { get; set; }
        public string Linetype { get; set; }
        public LineWeight Lineweight { get; set; }
        public string LineWeightString { get; set; }
        public AttachmentPoint Justification { get; set; }
        public string JustificationString { get; set; }
        public double Height { get; set; }
        public double Rotation { get; set; }
        public double WidthFactor { get; set; }
        public bool BackgroundFill { get; set; }
        public bool UseBackgroundColor { get; set; }
        public double BackgroundMaskScaleFactor { get; set; }
    }
}
