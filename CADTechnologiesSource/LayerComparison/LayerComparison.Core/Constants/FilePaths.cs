using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LayerComparison.Core.Constants
{
    public class FilePaths
    {
        public static string RecentComparisonPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "LayerComparison");

        public static string RecentComparisonFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "LayerComparison", "RecentComparisons.txt");
    }
}
