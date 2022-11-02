using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LayerComparison.Core.Models
{
    public class MissingLayerModel : IEquatable<MissingLayerModel>
    {
        public string SourceLayerName { get; set; }
        public string TargetDrawingPath { get; set; }

        public bool Equals(MissingLayerModel other)
        {
            if (other == null)
                return false;
            return
            SourceLayerName == other.SourceLayerName
            && TargetDrawingPath == other.TargetDrawingPath;
        }

        public override int GetHashCode()
        {
            //If obj is null then return 0
            if (this == null)
            {
                return 0;
            }
            //Get the hash code values for each property
            int SourceLayerNameHashCode = SourceLayerName == null ? 0 : SourceLayerName.GetHashCode();
            int TargetDrawingPathHashCode = TargetDrawingPath == null ? 0 : TargetDrawingPath.GetHashCode();

            return SourceLayerNameHashCode ^ TargetDrawingPathHashCode;
        }
    }
}
