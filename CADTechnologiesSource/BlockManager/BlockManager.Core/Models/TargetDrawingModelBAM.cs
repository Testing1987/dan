using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlockManager.Core.Models
{
    public class TargetDrawingModelBAM
    {
        public TargetDrawingModelBAM()
        {
            //needed for serialization.
        }
        public string DrawingPath { get; set; }
        public string TrimmedPath { get; set; }
        public bool Selected { get; set; }
        public string Visretain { get; set; }

        /// <summary>
        /// Overrides <see cref="ToString"/> to provide the actual value of the string to the UI.
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return DrawingPath;
        }
    }
}
