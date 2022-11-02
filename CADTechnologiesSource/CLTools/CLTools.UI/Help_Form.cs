using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CLTools.UI
{
    public partial class Help_Form : Form
    {
        public Help_Form(string helpTitle, string helpmessage)
        {
            InitializeComponent();
            label_help_message.Text = helpmessage;
            label_help_title.Text = helpTitle;
        }
    }
}
