using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Alignment_mdi
{
    public partial class Blank_Form : Form
    {
        public static Blank_Form tpage_blank = null;
        public Blank_Form()
        {
            InitializeComponent();
            tpage_blank = this;
        }

    }
}
