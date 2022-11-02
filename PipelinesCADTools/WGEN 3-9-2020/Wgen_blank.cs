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
    public partial class Wgen_Blank_form : Form
    {
        public Wgen_Blank_form()
        {
            InitializeComponent();
        }

        public void get_label_wait_visible (bool visible)
        {
            label_wait.Visible = visible;
        }
    }
}
