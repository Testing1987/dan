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
    public partial class Igen__Start_Page_form : Form
    {
        public Igen__Start_Page_form()
        {
            InitializeComponent();
        }

        private void button_agen_Click(object sender, EventArgs e)
        {

            int i = 0;
            do
            {
                System.Windows.Forms.Form Forma1 = System.Windows.Forms.Application.OpenForms[i] as System.Windows.Forms.Form;
                if (Forma1 is Alignment_mdi._AGEN_mainform)
                {
                    Forma1.Focus();
                    Forma1.WindowState = System.Windows.Forms.FormWindowState.Normal;
                    Forma1.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - Forma1.Width) / 2,
                      (Screen.PrimaryScreen.WorkingArea.Height - Forma1.Height) / 2);

                    this.Close();
                    return;
                }

                i = i + 1;
            } while (i < System.Windows.Forms.Application.OpenForms.Count);



            try
            {
                Alignment_mdi._AGEN_mainform forma2 = new Alignment_mdi._AGEN_mainform();
                Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma2);
                forma2.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - forma2.Width) / 2,
                     (Screen.PrimaryScreen.WorkingArea.Height - forma2.Height) / 2);
                this.Close();
            }
            catch (System.Exception EX)
            {
                MessageBox.Show(EX.Message);
            }
        }

        private void button_pt_inq_Click(object sender, EventArgs e)
        {

            int i = 0;
            do
            {
                System.Windows.Forms.Form Forma1 = System.Windows.Forms.Application.OpenForms[i] as System.Windows.Forms.Form;
                if (Forma1 is Alignment_mdi.Igen_main_form)
                {
                    Forma1.Focus();
                    Forma1.WindowState = System.Windows.Forms.FormWindowState.Normal;
                    Forma1.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - Forma1.Width) / 2,
                      (Screen.PrimaryScreen.WorkingArea.Height - Forma1.Height) / 2);

                    this.Close();
                    return;
                }

                i = i + 1;
            } while (i < System.Windows.Forms.Application.OpenForms.Count);


            try
            {
                Alignment_mdi.Igen_main_form forma2 = new Alignment_mdi.Igen_main_form();
                Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma2);
                forma2.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - forma2.Width) / 2,
                     (Screen.PrimaryScreen.WorkingArea.Height - forma2.Height) / 2);
                this.Close();
            }
            catch (System.Exception EX)
            {
                MessageBox.Show(EX.Message);
            }

        }
    }
}
