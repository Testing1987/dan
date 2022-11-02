using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LayerComparison.Core.Models
{
    public class DrawingModel
    {
        public DrawingModel()
        {
            Layers = new List<string>();
        }

        public string DrawingPath { get; set; }

        public List<string> Layers { get; set; }
    }
}
