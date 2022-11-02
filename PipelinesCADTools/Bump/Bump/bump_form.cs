using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace Bump
{
    public partial class bump_form : Form
    {
        //Global Variables
        private bool clickdragdown;
        private System.Drawing.Point lastLocation;




        #region set enable true or false    
        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(btn_draw);
            lista_butoane.Add(button_minimize);
            lista_butoane.Add(tbox_Bump_Value);
            lista_butoane.Add(comboBox_precision);
            lista_butoane.Add(rb_Imperial);
            lista_butoane.Add(rb_Metric);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = false;
            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(btn_draw);
            lista_butoane.Add(button_minimize);
            lista_butoane.Add(button_Exit);
            lista_butoane.Add(tbox_Bump_Value);
            lista_butoane.Add(comboBox_precision);
            lista_butoane.Add(rb_Imperial);
            lista_butoane.Add(rb_Metric);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }
        #endregion


        public bump_form()
        {
            InitializeComponent();
            this.comboBox_precision.Items.AddRange(new object[] { "0", "0.0", "0.00", "0.000" });
            comboBox_precision.SelectedIndex = 0;
        }

        private void button_minimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void button_Exit_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void clickmove_MouseDown(object sender, MouseEventArgs e)
        {
            clickdragdown = true;
            lastLocation = e.Location;
        }

        private void clickmove_MouseMove(object sender, MouseEventArgs e)
        {
            if (clickdragdown)
            {
                this.Location = new System.Drawing.Point(
                  (this.Location.X - lastLocation.X) + e.X, (this.Location.Y - lastLocation.Y) + e.Y);

                this.Update();
            }
        }

        private void clickmove_MouseUp(object sender, MouseEventArgs e)
        {
            clickdragdown = false;
        }

        public bool IsPresent(System.Windows.Forms.TextBox textBox1, string name1, string header)
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show(name1 + " is Missing Value", header, MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox1.Focus();
                return false;
            }

            return true;
        }

        private bool IsNumeric(string s)
        // checks to see if value selected is a number.
        {
            bool result1 = false;
            double myNum;

            if (Double.TryParse(s, out myNum))
            {
                result1 = true;
                if (s.Contains(",") == true)
                {
                    result1 = false;
                }
            }
            else
            {
                result1 = false;
            }
            return result1;
        }

        private void btn_bump_Click(object sender, EventArgs e)
        {
            //string string1 = "is a chainage -12+345.6 how are you";
            //string1 = STA_From_String(string1);

            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {
                set_enable_false();
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;


                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();

                        string message1 = "";
                        if (rb_Imperial.Checked == true)
                        {
                            message1 = "\nSelect Stationing You Want to Bump";
                        }

                        if (rb_Metric.Checked == true)
                        {
                            message1 = "\nSelect Chainage You Want to Bump";
                        }

                        Prompt_rez.MessageForAdding = message1;
                        Prompt_rez.SingleOnly = false;
                        Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            set_enable_true();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        int precision = 0;
                        if (comboBox_precision.Text == "0")
                        {
                            precision = 0;
                        }
                        if (comboBox_precision.Text == "0.0")
                        {
                            precision = 1;
                        }
                        if (comboBox_precision.Text == "0.00")
                        {
                            precision = 2;
                        }
                        if (comboBox_precision.Text == "0.000")
                        {
                            precision = 3;
                        }


                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {
                            string units = "";
                            BlockReference block1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForWrite) as BlockReference;
                            if (block1 != null)
                            {
                                if (block1.AttributeCollection.Count > 0)
                                {
                                    foreach (ObjectId atid in block1.AttributeCollection)
                                    {
                                        AttributeReference Atr1 = (AttributeReference)Trans1.GetObject(atid, OpenMode.ForWrite);

                                        if (Atr1 != null)
                                        {

                                            string oldname = "";
                                            if (Atr1.IsMTextAttribute == false)
                                            {
                                                oldname = Atr1.TextString;

                                            }
                                            else
                                            {
                                                oldname = Atr1.MTextAttribute.Contents;
                                            }

                                            string old_blk_value = Extract_STA_From_String_with_precision(oldname, precision);
                                            if (IsNumeric(old_blk_value.Replace("+", "")) == true)
                                            {
                                                double number1 = Convert.ToDouble(old_blk_value.Replace("+", ""));
                                                double new_number = number1 + Convert.ToDouble(tbox_Bump_Value.Text);

                                                if (rb_Imperial.Checked == true)
                                                {
                                                    units = "f";
                                                }
                                                if (rb_Metric.Checked == true)
                                                {
                                                    units = "m";
                                                }

                                                string new_blk_value = Get_chainage_from_double(new_number, units, precision);
                                                Atr1.TextString = Atr1.TextString.Replace(old_blk_value, new_blk_value);
                                            }

                                        }
                                    }

                                }
                            }

                            DBText text1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForWrite) as DBText;
                            if (text1 != null)
                            {
                                string textstring1 = text1.TextString;
                                string old_value = Extract_STA_From_String_with_precision(textstring1, precision);
                                if (IsNumeric(old_value.Replace("+", "")) == true)
                                {
                                    double number1 = Convert.ToDouble(old_value.Replace("+", ""));
                                    double new_number = number1 + Convert.ToDouble(tbox_Bump_Value.Text);
                                    if (rb_Imperial.Checked == true)
                                    {
                                        units = "f";
                                    }
                                    if (rb_Metric.Checked == true)
                                    {
                                        units = "m";
                                    }
                                    string new_value = Get_chainage_from_double(new_number, units, precision);
                                    text1.TextString = text1.TextString.Replace(old_value, new_value);
                                }

                            }

                            MText Mtext1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForWrite) as MText;
                            if (Mtext1 != null)
                            {
                                string textstring1 = Mtext1.Text;
                                string old_value = Extract_STA_From_String_with_precision(textstring1, precision);
                                if (IsNumeric(old_value.Replace("+", "")) == true)
                                {
                                    double number1 = Convert.ToDouble(old_value.Replace("+", ""));
                                    double new_number = number1 + Convert.ToDouble(tbox_Bump_Value.Text);
                                    if (rb_Imperial.Checked == true)
                                    {
                                        units = "f";
                                    }
                                    if (rb_Metric.Checked == true)
                                    {
                                        units = "m";
                                    }
                                    string new_value = Get_chainage_from_double(new_number, units, precision);
                                    Mtext1.Contents = Mtext1.Contents.Replace(old_value, new_value);
                                }

                            }

                            MLeader mleader1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForWrite) as MLeader;
                            if (mleader1 != null)
                            {



                                string textstring1 = mleader1.MText.Contents;
                                string old_value = Extract_STA_From_String_with_precision(textstring1, precision);
                                if (IsNumeric(old_value.Replace("+", "")) == true)
                                {
                                    double number1 = Convert.ToDouble(old_value.Replace("+", ""));
                                    double new_number = number1 + Convert.ToDouble(tbox_Bump_Value.Text);
                                    if (rb_Imperial.Checked == true)
                                    {
                                        units = "f";
                                    }
                                    if (rb_Metric.Checked == true)
                                    {
                                        units = "m";
                                    }
                                    string new_value = Get_chainage_from_double(new_number, units, precision);
                                    MText mtext2 = new MText();
                                    mtext2.Contents = textstring1.Replace(old_value, new_value);
                                    mtext2.TextHeight = mleader1.MText.TextHeight;
                                    mtext2.TextStyleId = mleader1.MText.TextStyleId;
                                    mleader1.MText = mtext2;


                                }

                            }

                        }


                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            set_enable_true();




        }

        public string Extract_STA_From_String(string text1)
        {
            string new_string = "";

            if (text1.Contains("+") == true && text1.Contains(".") == true)
            {

                List<string> list_STA = new List<string>();
                int pos_plus = text1.IndexOf("+");
                list_STA.Insert(0, "+");
                bool run1 = true;

                for (int i = pos_plus; i >= 1; --i)
                {
                    string letter1 = text1.Substring(i - 1, 1);
                    if (run1 == true)
                    {
                        if (letter1 == "0" || letter1 == "1" || letter1 == "2" || letter1 == "3" || letter1 == "4" || letter1 == "5" || letter1 == "6" || letter1 == "7" || letter1 == "8" || letter1 == "9")
                        {
                            list_STA.Insert(0, letter1);
                        }
                        else
                        {
                            if (letter1 == "-")
                            {
                                list_STA.Insert(0, letter1);

                            }
                            run1 = false;
                        }
                    }
                }

                run1 = true;
                bool dec_pt = false;
                for (int i = pos_plus + 1; i < text1.Length; ++i)
                {
                    string letter1 = text1.Substring(i, 1);
                    if (run1 == true)
                    {
                        if (letter1 == "0" || letter1 == "1" || letter1 == "2" || letter1 == "3" || letter1 == "4" || letter1 == "5" || letter1 == "6" || letter1 == "7" || letter1 == "8" || letter1 == "9")
                        {
                            list_STA.Add(letter1);
                            if (dec_pt == true) run1 = false;
                        }
                        else
                        {
                            if (letter1 == ".")
                            {
                                list_STA.Add(letter1);
                                dec_pt = true;
                            }
                            else
                            {
                                if (dec_pt == true)
                                {
                                    run1 = false;
                                }
                            }
                        }
                    }
                }

                for (int i = 0; i < list_STA.Count; ++i)
                {
                    new_string = new_string + list_STA[i];
                }

            }

            else
            {
                new_string = "no station";
            }


            return new_string;

        }

        public string Extract_STA_From_String_with_precision(string text1, int round1)
        {
            string new_string = "";

            if (text1.Contains("+") == true)
            {

                List<string> list_STA = new List<string>();
                int pos_plus = text1.IndexOf("+");
                list_STA.Insert(0, "+");
                bool run1 = true;

                for (int i = pos_plus; i >= 1; --i)
                {
                    string letter1 = text1.Substring(i - 1, 1);
                    if (run1 == true)
                    {
                        if (letter1 == "0" || letter1 == "1" || letter1 == "2" || letter1 == "3" || letter1 == "4" || letter1 == "5" || letter1 == "6" || letter1 == "7" || letter1 == "8" || letter1 == "9")
                        {
                            list_STA.Insert(0, letter1);
                        }
                        else
                        {
                            if (letter1 == "-")
                            {
                                list_STA.Insert(0, letter1);

                            }
                            run1 = false;
                        }
                    }
                }

                run1 = true;
                bool dec_pt = false;


                for (int i = pos_plus + 1; i < text1.Length; ++i)
                {
                    string letter1 = text1.Substring(i, 1);
                    if (run1 == true)
                    {
                        if (letter1 == "0" || letter1 == "1" || letter1 == "2" || letter1 == "3" || letter1 == "4" || letter1 == "5" || letter1 == "6" || letter1 == "7" || letter1 == "8" || letter1 == "9")
                        {
                            list_STA.Add(letter1);

                        }
                        else
                        {
                            if (letter1 == ".")
                            {
                                list_STA.Add(letter1);
                                dec_pt = true;
                            }
                            else
                            {
                                if (dec_pt == true)
                                {
                                    string l1 = list_STA[list_STA.Count - 1];
                                    if (l1 == ".")
                                    {
                                        list_STA.RemoveAt(list_STA.Count - 1);
                                    }
                                }
                                run1 = false;

                            }
                        }
                    }
                }

                for (int i = 0; i < list_STA.Count; ++i)
                {
                    new_string = new_string + list_STA[i];
                }

            }

            else
            {
                new_string = "no station";
            }

            return new_string;

        }

        static public string Get_String_Rounded(double Numar, int Nr_dec)
        {

            String String1, String2, Zero, zero1;
            Zero = "";
            zero1 = "";

            String String_punct = "";

            if (Nr_dec > 0)
            {
                String_punct = ".";
                for (int i = 1; i <= Nr_dec; i = i + 1)
                {
                    Zero = Zero + "0";
                }
            }

            string String_minus = "";

            if (Numar < 0)
            {
                String_minus = "-";
                Numar = -Numar;
            }

            String1 = Math.Round(Numar, Nr_dec, MidpointRounding.AwayFromZero).ToString();

            String2 = String1;

            if (String1.Contains(".") == false)
            {
                String2 = String1 + String_punct + Zero;
                goto end;
            }

            if (String1.Length - String1.IndexOf(".") - 1 - Nr_dec != 0)
            {
                for (int i = 1; i <= String1.IndexOf(".") + 1 + Nr_dec - String1.Length; i = i + 1)
                {
                    zero1 = zero1 + "0";
                }

                String2 = String1 + zero1;
            }

        end:
            return String_minus + String2;

        }

        static public string Get_chainage_from_double(double Numar, string units, int Nr_dec)
        {

            string String2, String3;
            String3 = "";
            String String_minus = "";

            if (Numar < 0)
            {
                String_minus = "-";
                Numar = -Numar;
            }




            String2 = Get_String_Rounded(Numar, Nr_dec);




            int Punct;
            if (String2.Contains(".") == false)
            {
                Punct = 0;
            }
            else
            {
                Punct = 1;
            }


            if (String2.Length - Nr_dec - Punct >= 4)
            {
                if (units == "f") String3 = String2.Substring(0, String2.Length - 2 - Nr_dec - Punct) + "+" + String2.Substring(String2.Length - (2 + Nr_dec + Punct));
                if (units == "m") String3 = String2.Substring(0, String2.Length - 3 - Nr_dec - Punct) + "+" + String2.Substring(String2.Length - (3 + Nr_dec + Punct));
            }
            else
            {
                if (units == "f")
                {
                    if (String2.Length - Nr_dec - Punct == 1) String3 = "0+0" + String2;
                    if (String2.Length - Nr_dec - Punct == 2) String3 = "0+" + String2;
                    if (String2.Length - Nr_dec - Punct == 3) String3 = String2.Substring(0, 1) + "+" + String2.Substring(String2.Length - (2 + Nr_dec + Punct));
                }
                if (units == "m")
                {
                    if (String2.Length - Nr_dec - Punct == 1) String3 = "0+00" + String2;
                    if (String2.Length - Nr_dec - Punct == 2) String3 = "0+0" + String2;
                    if (String2.Length - Nr_dec - Punct == 3) String3 = "0+" + String2;
                }
            }


            return String_minus + String3;

        }

        bool isNumeric2(string letter1, int pos1, bool dot)
        {
            if (letter1.Length == 1)
            {
                if (letter1 == "0" || letter1 == "1" || letter1 == "2" || letter1 == "3" || letter1 == "4" || letter1 == "5" || letter1 == "6" || letter1 == "7" || letter1 == "8" || letter1 == "9")
                {
                    return true;
                }
                else
                {
                    if (pos1 == 0 && letter1 == "-")
                    {
                        return true;
                    }
                    else
                    {
                        if (dot == true && letter1 == ".")
                        {
                            return true;
                        }
                    }
                }
            }

            return false;
        }

        private void tbox_Bump_Value_TextChanged_1(object sender, EventArgs e)
        {
            int wrd_length = tbox_Bump_Value.Text.Length;
            string contents = tbox_Bump_Value.Text;
            List<string> banned_list = new List<string>();

            if (wrd_length > 0)
            {
                for (int i = 0; i < wrd_length; ++i)
                {
                    string oops = contents.Substring(i, 1);
                    banned_list.Add(oops);
                }

                string first_val = banned_list[0];
                bool isdot = false;
                if (first_val == ".")
                {
                    isdot = true;
                }
                if (isNumeric2(first_val, 0, isdot) == false)
                {
                    banned_list.RemoveAt(0);
                }

                int rep_dot = 0;
                for (int i = banned_list.Count - 1; i >= 0; --i)
                {
                    bool isdot1 = false;
                    string value1 = banned_list[i];
                    if (value1 == "." && isdot1 == false && rep_dot == 0)
                    {
                        isdot1 = true;

                    }

                    if (isNumeric2(value1, i, isdot1) == false)
                    {
                        banned_list.RemoveAt(i);
                    }

                    if (isdot1 == true)
                    {
                        ++rep_dot;
                    }
                }

                string newstring = "";

                if (banned_list.Count > 0)
                {
                    for (int i = 0; i < banned_list.Count; ++i)
                    {
                        string value1 = banned_list[i];
                        newstring = newstring + value1;
                    }
                }

                tbox_Bump_Value.Text = newstring;


            }


        }


    }




}
