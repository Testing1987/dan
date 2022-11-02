using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Alignment_mdi
{
    public class MyColors : ProfessionalColorTable
    {
        public override Color MenuItemSelected
        {
            get
            {
                return Color.FromArgb(0, 122, 204);
            }
        }
        public override Color MenuItemBorder
        {
            get
            {
                return Color.White;
            }
        }
    }
}
