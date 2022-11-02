using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LayerComparison.Core.Models
{
    public class ConflictModel
    {
        public string Drawing
        {
            get; set;
        }
        public string IssueFound { get; set; }

        public int IssueQuantity { get; set; }
    }
}
