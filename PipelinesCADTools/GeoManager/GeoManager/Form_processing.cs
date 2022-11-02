using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MMGeoTools
{
    public partial class Form_processing : Form
    {
        public Form_processing()
        {
            InitializeComponent();
        }

        private void Form_processing_Load(object sender, EventArgs e)
        {
            this.Top = this.Owner.Top + this.Owner.Height / 2 - this.Height / 2;
            this.Left = this.Owner.Left + this.Owner.Width / 2 - this.Width/2;
        }
    }
}
