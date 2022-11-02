using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CADTechnologiesSource.All.Models
{
    public class TargetDrawingModel
    {
        public TargetDrawingModel()
        {
            //needed for serialization.
        }
        public string DrawingPath { get; set; }
        public bool Selected { get; set; }
        public string Visretain { get; set; }

        /// <summary>
        /// Overrides <see cref="ToString"/> to provide the actual value of the string to the UI.
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return this.DrawingPath;
        }
    }
}
