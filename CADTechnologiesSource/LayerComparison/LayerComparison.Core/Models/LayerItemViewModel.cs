using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LayerComparison.Core.Models
{
    public class LayerItemViewModel
    {
        public string Drawing { get; set; }
        public string Name { get; set; }
        public string Color { get; set; }
        public bool On { get; set; }
        public bool Freeze { get; set; }
        public string Linetype { get; set; }
        public string Lineweight { get; set; }
        public string Transparency { get; set; }
        public bool Plot { get; set; }
    }
}
