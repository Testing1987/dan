using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace Alignment_mdi
{
    public partial class Lgen_alias_page : Form
    {
        #region Public Members

        #endregion

        #region Private Members

        private string ObjectCBOXGenericItem = "Objects with no object data";

        #endregion

        #region Constructor

        public Lgen_alias_page()
        {
            InitializeComponent();
        }

        #endregion

        #region ButtonClicks

        private void Button_LGEN_01_Load_Layers_Click(object sender, EventArgs e)
        {
            comboBox_layers.Items.Clear();
           // Func.load_OD_Tables_to_Combobox(comboBox_layeralias);
            comboBox_layers.Items.Add(ObjectCBOXGenericItem);
        }

        private void button_minus1_Click(object sender, EventArgs e)
        {
            int nrmax = comboBox_layers.Items.Count;
            if (nrmax > 0)
            {
                int index1 = comboBox_layers.SelectedIndex;

                if (index1 == 0)
                {
                    index1 = nrmax - 1;
                }
                else if (index1 > 0)
                {
                    index1 = index1 - 1;
                }
                comboBox_layers.SelectedIndex = index1;
            }
        }

        private void button_plus1_Click(object sender, EventArgs e)
        {
            int nrmax = comboBox_layers.Items.Count;
            if (nrmax > 0)
            {
                int index1 = comboBox_layers.SelectedIndex;

                if (index1 == nrmax - 1)
                {
                    index1 = 0;
                }
                else if (index1 < nrmax)
                {
                    index1 = index1 + 1;
                }
                comboBox_layers.SelectedIndex = index1;
            }
        }

        #endregion

        #region TestButtons

        private void button1_Click(object sender, EventArgs e)
        {
           // Lgen_mainform.Data_table_layer_alias = Func.Create_layer_alias_datatable_structure();
            //Func.Transfer_datatable_to_new_excel_spreadsheet(Lgen_mainform.Data_table_layer_alias);
        }

        #endregion
    }
}
