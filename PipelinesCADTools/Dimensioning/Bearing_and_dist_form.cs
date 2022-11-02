
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

namespace Dimensioning
{
    public partial class Bearing_and_dist_form : Form
    {

        bool Freeze_operations = false;

        public Bearing_and_dist_form()
        {
            InitializeComponent();

            Functions.Incarca_existing_Blocks_with_attributes_to_combobox(comboBox_blocks);
        }

        private void Bearing_and_dist_form_Load(object sender, EventArgs e)
        {
            comboBox_Precision.SelectedIndex = 0;
            comboBox_Scale.SelectedIndex = 0;
            comboBox_Label_Position.SelectedIndex = 0;
            textBox_CT_Index.Visible = false;
            textBox_LT_Index.Visible = false;
            label_Index.Visible = false;

        }

        private void radioButton_LT_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton_LT.Checked == true)
            {
                textBox_CT_Index.Visible = false;
                textBox_LT_Index.Visible = true;
                label_Index.Visible = true;

            }


            if (radioButton_NE.Checked == true)
            {
                comboBox_Label_Position.Visible = false;

            }
            else
            {

                comboBox_Label_Position.Visible = true;
            }
            comboBox_Label_Position.SelectedIndex = 0;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

        }

        private void radioButton_CT_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton_CT.Checked == true)
            {
                textBox_CT_Index.Visible = true;
                textBox_LT_Index.Visible = false;
                label_Index.Visible = true;

            }


            if (radioButton_NE.Checked == true)
            {
                comboBox_Label_Position.Visible = false;

            }
            else
            {

                comboBox_Label_Position.Visible = true;
            }
            comboBox_Label_Position.SelectedIndex = 0;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
        }

        private void radioButton_others(object sender, EventArgs e)
        {
            if (radioButton_CT.Checked == false & radioButton_LT.Checked == false)
            {
                textBox_CT_Index.Visible = false;
                textBox_LT_Index.Visible = false;
                label_Index.Visible = false;
            }

            comboBox_Label_Position.SelectedIndex = 0;
            if (radioButton_PI.Checked == true)
            {
                comboBox_Precision.Visible = false;
                comboBox_Label_Position.Visible = false;
                textBox_textheight.Visible = false;
                label_precision.Visible = false;
                label_textheight.Visible = false;
            }
            else
            {
                comboBox_Precision.Visible = true;
                comboBox_Label_Position.Visible = true;
                textBox_textheight.Visible = true;
                label_precision.Visible = true;
                label_textheight.Visible = true;
            }




            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
        }

        private void radioButton_ne(object sender, EventArgs e)
        {

            if (radioButton_NE.Checked == true)
            {
                comboBox_Label_Position.Items.Remove("Top");
                comboBox_Label_Position.Items.Remove("Bottom");
                comboBox_Label_Position.Items.Insert(0, "Leader");


            }
            else
            {
                comboBox_Label_Position.Items.Remove("Leader");
                comboBox_Label_Position.Items.Insert(0, "Bottom");
                comboBox_Label_Position.Items.Insert(0, "Top");


            }
            comboBox_Label_Position.Visible = true;
            comboBox_Label_Position.SelectedIndex = 0;





            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
        }

        private void comboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
        }

        public double Calculate_text_height()
        {

            string Scale_string = comboBox_Scale.Text;
            Double Text_height = 0;
            double f1 = 1;
            switch (Scale_string)
            {
                case "1 = 10":
                    f1 = 10;
                    break;
                case "1 = 20":
                    f1 = 20;
                    break;
                case "1 = 30":
                    f1 = 30;
                    break;
                case "1 = 40":
                    f1 = 40;
                    break;
                case "1 = 50":
                    f1 = 50;
                    break;
                case "1 = 60":
                    f1 = 60;
                    break;
                case "1 = 100":
                    f1 = 100;
                    break;
                case "1 = 200":
                    f1 = 200;
                    break;
                case "1 = 300":
                    f1 = 300;
                    break;
                case "1 = 400":
                    f1 = 400;
                    break;
                case "1 = 500":
                    f1 = 500;
                    break;
                case "1 = 600":
                    f1 = 600;
                    break;
                case "1 = 1000":
                    f1 = 1000;
                    break;
                case "1 = 2000":
                    f1 = 2000;
                    break;
                default:
                    f1 = 1;
                    break;

            }
            if (Functions.IsNumeric(textBox_textheight.Text) == true)
            {
                Text_height = f1 * Convert.ToDouble(textBox_textheight.Text);
            }

            return Text_height;
        }

        public int Rounding()
        {
            string Scale_string = comboBox_Precision.Text;
            int Round1 = 0;
            switch (Scale_string)
            {
                case "0.0":
                    Round1 = 1;
                    break;
                case "0.00":
                    Round1 = 2;
                    break;
                case "0.000":
                    Round1 = 3;
                    break;
                case "0.0000":
                    Round1 = 4;
                    break;
                default:
                    Round1 = 0;
                    break;
            }

            return Round1;
        }

        public double Calculate_BLOCK_SCALE()
        {
            string Scale_string = comboBox_Scale.Text;

            double f1 = 1;
            switch (Scale_string)
            {
                case "1 = 10":
                    f1 = 10;
                    break;
                case "1 = 20":
                    f1 = 20;
                    break;
                case "1 = 30":
                    f1 = 30;
                    break;
                case "1 = 40":
                    f1 = 40;
                    break;
                case "1 = 50":
                    f1 = 50;
                    break;
                case "1 = 60":
                    f1 = 60;
                    break;
                case "1 = 100":
                    f1 = 100;
                    break;
                case "1 = 200":
                    f1 = 200;
                    break;
                case "1 = 300":
                    f1 = 300;
                    break;
                case "1 = 400":
                    f1 = 400;
                    break;
                case "1 = 500":
                    f1 = 500;
                    break;
                case "1 = 600":
                    f1 = 600;
                    break;
                case "1 = 1000":
                    f1 = 1000;
                    break;
                case "1 = 2000":
                    f1 = 2000;
                    break;
                default:
                    f1 = 1;
                    break;

            }


            return f1;
        }

        public Point3d get_insertion_pt(Point3d PointM, Point3d Point2, double Bearing1, double TextH)
        {
            Point3d Point_ins;
            Line Linie1 = new Line(PointM, Point2);
            if (comboBox_Label_Position.Text == "Top")
            {
                if (Bearing1 <= Math.PI / 2)
                {
                    Linie1.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, PointM));
                }
                if (Bearing1 > Math.PI / 2 & Bearing1 <= Math.PI)
                {
                    Linie1.TransformBy(Matrix3d.Rotation(-Math.PI / 2, Vector3d.ZAxis, PointM));
                }
                if (Bearing1 > Math.PI & Bearing1 <= 3 * Math.PI / 2)
                {
                    Linie1.TransformBy(Matrix3d.Rotation(-Math.PI / 2, Vector3d.ZAxis, PointM));
                }
                if (Bearing1 > 3 * Math.PI / 2)
                {
                    Linie1.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, PointM));
                }
            }
            else
            {
                if (Bearing1 <= Math.PI / 2)
                {
                    Linie1.TransformBy(Matrix3d.Rotation(-Math.PI / 2, Vector3d.ZAxis, PointM));
                }
                if (Bearing1 > Math.PI / 2 & Bearing1 <= Math.PI)
                {
                    Linie1.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, PointM));
                }
                if (Bearing1 > Math.PI & Bearing1 <= 3 * Math.PI / 2)
                {
                    Linie1.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, PointM));
                }
                if (Bearing1 > 3 * Math.PI / 2)
                {
                    Linie1.TransformBy(Matrix3d.Rotation(-Math.PI / 2, Vector3d.ZAxis, PointM));
                }
            }

            if (Linie1.Length < TextH)
            {
                Linie1.TransformBy(Matrix3d.Scaling(1.1 * TextH / Linie1.Length, PointM));
            }
            Point_ins = Linie1.GetPointAtDist(TextH / 2);

            return Point_ins;
        }

        public string Increase_index(string Old_string, int factor)
        {
            string new_string = "";
            int old_number = Extrage_nr_de_la_finalul_textului(Old_string);


            if (old_number == -1000000000)
            {
                new_string = Old_string;
            }
            else
            {
                new_string = Old_string.Replace(old_number.ToString(), "") + (old_number + factor).ToString();
            }

            return new_string;

        }

        public int Extrage_nr_de_la_finalul_textului(string Text1)
        {
            int Numar1 = -1000000000;
            String Numar_string = "";
            for (int i = Text1.Length - 1; i >= 0; i = i - 1)
            {
                string Letter = Text1.Substring(i, 1);
                if (Functions.IsNumeric(Letter) == true)
                {
                    Numar_string = Letter + Numar_string;
                }
                else
                {
                    i = -1;
                }
            }

            if (Numar_string != "")
            {
                Numar1 = Convert.ToInt32(Numar_string);
            }

            return Numar1;
        }

        private void Insert_blocks(System.Data.DataTable Table1, string Block_Name)
        {
            double Hb = 0.152;
            if(Functions.IsNumeric(textBox_block_height.Text)==true)
            {
                Hb = Math.Abs(Convert.ToDouble(textBox_block_height.Text));
            }

            if (Table1 != null)
            {
                if (Table1.Rows.Count > 0)
                {
                    try
                    {
                        Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                        Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                        double SCALE1 = Calculate_BLOCK_SCALE();

                        using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                        {

                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                            {


                                Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify the table insertion point:");
                                PP1.AllowNone = false;
                                Point_res1 = Editor1.GetPoint(PP1);

                                if (Point_res1.Status != PromptStatus.OK)
                                {
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    return;
                                }

                                for (int i = 0; i < Table1.Rows.Count; ++i)
                                {
                                    System.Collections.Specialized.StringCollection Col1 = new System.Collections.Specialized.StringCollection();
                                    System.Collections.Specialized.StringCollection Col2 = new System.Collections.Specialized.StringCollection();

                                    if (Table1.Columns.Contains("CURVE") == true)
                                    {
                                        if (Table1.Rows[i]["CURVE"] != DBNull.Value)
                                        {
                                            Col1.Add("CURVE");
                                            Col2.Add(Table1.Rows[i]["CURVE"].ToString());
                                        }
                                    }
                                    if (Table1.Columns.Contains("RADIUS") == true)
                                    {
                                        if (Table1.Rows[i]["RADIUS"] != DBNull.Value)
                                        {
                                            Col1.Add("RADIUS");
                                            Col2.Add(Table1.Rows[i]["RADIUS"].ToString());
                                        }
                                    }
                                    if (Table1.Columns.Contains("ARC_DIST") == true)
                                    {
                                        if (Table1.Rows[i]["ARC_DIST"] != DBNull.Value)
                                        {
                                            Col1.Add("ARC_DIST");
                                            Col2.Add(Table1.Rows[i]["ARC_DIST"].ToString());
                                        }
                                    }
                                    if (Table1.Columns.Contains("DELTA") == true)
                                    {
                                        if (Table1.Rows[i]["DELTA"] != DBNull.Value)
                                        {
                                            Col1.Add("DELTA");
                                            Col2.Add(Table1.Rows[i]["DELTA"].ToString());
                                        }
                                    }
                                    if (Table1.Columns.Contains("TANGENT") == true)
                                    {
                                        if (Table1.Rows[i]["TANGENT"] != DBNull.Value)
                                        {
                                            Col1.Add("TANGENT");
                                            Col2.Add(Table1.Rows[i]["TANGENT"].ToString());
                                        }
                                    }

                                    if (Table1.Columns.Contains("DIRECTION") == true)
                                    {
                                        if (Table1.Rows[i]["DIRECTION"] != DBNull.Value)
                                        {
                                            Col1.Add("DIRECTION");
                                            Col2.Add(Table1.Rows[i]["DIRECTION"].ToString());
                                        }
                                    }

                                    if (Table1.Columns.Contains("CHORD") == true)
                                    {
                                        if (Table1.Rows[i]["CHORD"] != DBNull.Value)
                                        {
                                            Col1.Add("CHORD");
                                            Col2.Add(Table1.Rows[i]["CHORD"].ToString());
                                        }
                                    }

                                    if (Table1.Columns.Contains("LINE") == true)
                                    {
                                        if (Table1.Rows[i]["LINE"] != DBNull.Value)
                                        {
                                            Col1.Add("LINE");
                                            Col2.Add(Table1.Rows[i]["LINE"].ToString());
                                        }
                                    }
                                    if (Table1.Columns.Contains("BEARING") == true)
                                    {
                                        if (Table1.Rows[i]["BEARING"] != DBNull.Value)
                                        {
                                            Col1.Add("BEARING");
                                            Col2.Add(Table1.Rows[i]["BEARING"].ToString());
                                        }
                                    }
                                    if (Table1.Columns.Contains("DISTANCE") == true)
                                    {
                                        if (Table1.Rows[i]["DISTANCE"] != DBNull.Value)
                                        {
                                            Col1.Add("DISTANCE");
                                            Col2.Add(Table1.Rows[i]["DISTANCE"].ToString());
                                        }
                                    }

                                    if (Table1.Columns.Contains("PT_START") == true)
                                    {
                                        if (Table1.Rows[i]["PT_START"] != DBNull.Value)
                                        {
                                            Col1.Add("PT_START");
                                            Col2.Add(Table1.Rows[i]["PT_START"].ToString());
                                        }
                                    }

                                    if (Table1.Columns.Contains("PT_END") == true)
                                    {
                                        if (Table1.Rows[i]["PT_END"] != DBNull.Value)
                                        {
                                            Col1.Add("PT_END");
                                            Col2.Add(Table1.Rows[i]["PT_END"].ToString());
                                        }
                                    }

                                    Point3d Ptins = new Point3d(Point_res1.Value.X, Point_res1.Value.Y - i * Hb * SCALE1, 0);

                                    Functions.InsertBlock_with_multiple_atributes("", Block_Name, Ptins, SCALE1, "0", Col1, Col2);

                                }

                                Trans1.Commit();
                            }

                        }
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }

        }

        AttachmentPoint Get_mTEXT_attachment()
        {
            AttachmentPoint Atch = AttachmentPoint.BottomLeft;

            if (comboBox_Label_Position.Text == "Top")
            {
                Atch = AttachmentPoint.BottomCenter;
            }
            if (comboBox_Label_Position.Text == "Bottom")
            {
                Atch = AttachmentPoint.TopCenter;
            }

            return Atch;
        }

        private bool background_fill()
        {

            if (checkBox_background_mask.Checked == false)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        private string suffix1()
        {

            if (checkBox_tolerance.Checked == true)
            {
                return "\u00B1";
            }

            return "";
        }

        private void button_Label_Click(object sender, EventArgs e)
        {
            bool space_in_bear = true;
            if (checkBox_no_space_in_bearing.Checked==true)
            {
                space_in_bear = false;
            }

            if (Freeze_operations == false)
            {
                Freeze_operations = true;


                try
                {
                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                    System.Data.DataTable TableL = new System.Data.DataTable();
                    System.Data.DataTable TableR = new System.Data.DataTable();

                    TableL.Columns.Add("LINE", typeof(string));
                    TableL.Columns.Add("BEARING", typeof(string));
                    TableL.Columns.Add("DISTANCE", typeof(string));

                    TableR.Columns.Add("CURVE", typeof(string));
                    TableR.Columns.Add("RADIUS", typeof(string));
                    TableR.Columns.Add("ARC_DIST", typeof(string));
                    TableR.Columns.Add("DELTA", typeof(string));
                    TableR.Columns.Add("TANGENT", typeof(string));
                    TableR.Columns.Add("DIRECTION", typeof(string));
                    TableR.Columns.Add("CHORD", typeof(string));

                    using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                    {


                        Double vp_SCALE = 1;
                        Double vp_TWIST = 0;
                        ObjectId Ent_vp_id = ObjectId.Null;
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            if (comboBox_Scale.Text == "PSpace")
                            {
                                int Tilemode1 = Convert.ToInt32(Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("TILEMODE"));
                                int CVport1 = Convert.ToInt32(Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("CVPORT"));

                                if (Tilemode1 == 0)
                                {
                                    if (CVport1 != 1)
                                    {
                                        Editor1.SwitchToPaperSpace();
                                    }

                                }
                                else
                                {
                                    MessageBox.Show("this option is meant to work in paper space \r\nPlease contact HECTOR MORALES to teach you how to switch in paper space!");
                                    Freeze_operations = false;
                                    return;
                                }


                                Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_Viewport;
                                Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_viewport;
                                Prompt_viewport = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the viewport:");
                                Prompt_viewport.SetRejectMessage("\nSelect a viewport!");
                                Prompt_viewport.AllowNone = true;
                                Prompt_viewport.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Viewport), false);
                                Prompt_viewport.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);

                                Rezultat_Viewport = ThisDrawing.Editor.GetEntity(Prompt_viewport);

                                if (Rezultat_Viewport.Status != PromptStatus.OK)
                                {
                                    MessageBox.Show("this requires you to select a viewport \r\nPlease contact HECTOR MORALES to teach you how to select a viewport!");
                                    Freeze_operations = false;
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    return;
                                }





                                Viewport Ent_vp = Trans1.GetObject(Rezultat_Viewport.ObjectId, OpenMode.ForRead) as Viewport;
                                if (Ent_vp == null)
                                {
                                    Polyline Ent_poly = Trans1.GetObject(Rezultat_Viewport.ObjectId, OpenMode.ForRead) as Polyline;
                                    if (Ent_poly != null)
                                    {
                                        ObjectId vpId = LayoutManager.Current.GetNonRectangularViewportIdFromClipId(Rezultat_Viewport.ObjectId);
                                        if (Trans1.GetObject(vpId, OpenMode.ForRead) is Viewport)
                                        {
                                            Ent_vp = Trans1.GetObject(vpId, OpenMode.ForRead) as Viewport;
                                        }

                                    }
                                }

                                if (Ent_vp != null)
                                {

                                    vp_SCALE = Ent_vp.CustomScale;
                                    vp_TWIST = Ent_vp.TwistAngle;
                                    Ent_vp_id = Ent_vp.ObjectId;

                                }



                            }
                            Trans1.Commit();
                        }


                    l123:


                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                            BlockTable BlockTable1 = (BlockTable)Trans1.GetObject(ThisDrawing.Database.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);


                            string NumeBlock = "";
                            if (radioButton_CT.Checked == true)
                            {
                                if (BlockTable1.Has("Curve_Table") == true)
                                {
                                    NumeBlock = "Curve_Table";
                                }
                                else
                                {
                                    radioButton_BD.Checked = true;
                                    MessageBox.Show("Curve_table block is not inserted on this drawing \r\nPlease contact HECTOR MORALES!");
                                }
                            }
                            if (radioButton_LT.Checked == true)
                            {
                                if (BlockTable1.Has("Line_Table") == true)
                                {
                                    NumeBlock = "Line_Table";
                                }
                                else
                                {
                                    radioButton_BD.Checked = true;
                                    MessageBox.Show("Line_Table block is not inserted on this drawing \r\nPlease contact HECTOR MORALES!");
                                }
                            }





                            if (radioButton_BD.Checked == true | radioButton_B.Checked == true | radioButton_D.Checked == true | radioButton_LT.Checked == true | radioButton_CT.Checked == true)
                            {


                                Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_poly1 = null;

                                if (radioButton_CT.Checked == true)
                                {
                                    Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_Poly1;
                                    Prompt_Poly1 = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect a polyline or arc:");
                                    Prompt_Poly1.SetRejectMessage("\nSelect a line, polyline or arc!");
                                    Prompt_Poly1.AllowNone = true;
                                    Prompt_Poly1.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                                    Prompt_Poly1.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Arc), false);
                                    Rezultat_poly1 = ThisDrawing.Editor.GetEntity(Prompt_Poly1);

                                    if (Rezultat_poly1.Status != PromptStatus.OK)
                                    {
                                        if (NumeBlock != "")
                                        {
                                            if (radioButton_CT.Checked == true)
                                            {
                                                Insert_blocks(TableR, NumeBlock);
                                            }

                                            if (radioButton_LT.Checked == true)
                                            {
                                                Insert_blocks(TableL, NumeBlock);
                                            }
                                        }

                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                        Freeze_operations = false;
                                        Trans1.Commit();
                                        return;
                                    }
                                }

                                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the first point");
                                PP1.AllowNone = false;
                                Point_res1 = Editor1.GetPoint(PP1);

                                if (Point_res1.Status != PromptStatus.OK)
                                {


                                    if (NumeBlock != "")
                                    {
                                        if (radioButton_CT.Checked == true)
                                        {
                                            Insert_blocks(TableR, NumeBlock);
                                        }

                                        if (radioButton_LT.Checked == true)
                                        {
                                            Insert_blocks(TableL, NumeBlock);
                                        }
                                    }


                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    Freeze_operations = false;
                                    Trans1.Commit();
                                    return;
                                }


                                Point3d Point1 = new Point3d();
                                Point1 = Point_res1.Value;

                                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                                Autodesk.AutoCAD.EditorInput.PromptPointOptions PP2;
                                PP2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the second point");
                                PP2.AllowNone = false;
                                PP2.BasePoint = Point1;
                                PP2.UseBasePoint = true;
                                Point_res2 = Editor1.GetPoint(PP2);

                                if (Point_res2.Status != PromptStatus.OK)
                                {
                                    if (NumeBlock != "")
                                    {
                                        if (radioButton_CT.Checked == true)
                                        {
                                            Insert_blocks(TableR, NumeBlock);
                                        }

                                        if (radioButton_LT.Checked == true)
                                        {
                                            Insert_blocks(TableL, NumeBlock);
                                        }
                                    }

                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    Freeze_operations = false;
                                    Trans1.Commit();
                                    return;
                                }
                                Point3d Point2 = new Point3d();
                                Point2 = Point_res2.Value;




                                double TextH = Calculate_text_height();


                                Point3d Pt1 = new Point3d();
                                Point3d Pt2 = new Point3d();

                                Point3d Pt1_nj = new Point3d();
                                Point3d Pt2_nj = new Point3d();

                                double x1 = Pt1.X;
                                double y1 = Pt1.Y;
                                double x2 = Pt2.X;
                                double y2 = Pt2.Y;
                                double Bulge1 = 0;

                                string Content1 = "";
                                double Dist1 = 0;
                                MText Mtext1 = new MText();

                                if (radioButton_CT.Checked == true)
                                {
                                    Entity Ent1 = (Entity)Trans1.GetObject(Rezultat_poly1.ObjectId, OpenMode.ForRead);
                                    if (Ent1 is Polyline | Ent1 is Arc)
                                    {
                                        Curve Curve1 = (Curve)Ent1;
                                        Pt1 = Curve1.GetClosestPointTo(Point1, Vector3d.ZAxis, false);
                                        Pt2 = Curve1.GetClosestPointTo(Point2, Vector3d.ZAxis, false);

                                        double Radius1 = 0;

                                        string Unghi_la_centru = "";
                                        x1 = Pt1.X;
                                        y1 = Pt1.Y;
                                        x2 = Pt2.X;
                                        y2 = Pt2.Y;

                                        Dist1 = Math.Pow(Math.Pow(x1 - x2, 2) + Math.Pow(y1 - y2, 2), 0.5);

                                        if (Ent1 is Polyline)
                                        {
                                            Polyline Poly_With_arc = (Polyline)Ent1;
                                            double Param1 = Poly_With_arc.GetParameterAtPoint(Pt1);
                                            double Param2 = Poly_With_arc.GetParameterAtPoint(Pt2);

                                            bool Execute1 = false;

                                            if (Math.Round(Param1, 4) == 0 & Math.Round(Param2, 4) >= Poly_With_arc.NumberOfVertices - 1) // THIS IS IN CASE THE LAST SEGMENT OF A CLOSED POLY IS AN ARC
                                            {
                                                Param1 = Param2;
                                                Param2 = 0;
                                                Execute1 = true;
                                            }
                                            else if (Math.Round(Param2, 4) == 0 & Math.Round(Param1, 4) == Poly_With_arc.NumberOfVertices - 1)// THIS IS IN CASE THE LAST SEGMENT OF A CLOSED POLY IS AN ARC
                                            {
                                                Execute1 = true;
                                            }
                                            else
                                            {
                                                if (Math.Round(Param1, 4) >= Math.Round(Param2, 4))
                                                {
                                                    double T = Param1;
                                                    Param1 = Param2;
                                                    Param2 = T;
                                                }
                                            }


                                            if (Math.Round(Param2, 4) - Math.Round(Param1, 4) <= 1)
                                            {
                                                Execute1 = true;

                                            }


                                            if (Poly_With_arc.GetSegmentType(Convert.ToInt32(Math.Floor(Param1))) != SegmentType.Arc)
                                            {
                                                Freeze_operations = false;
                                                MessageBox.Show("what you picked is not an arc segment \\r\\nyou may use osnap.....");
                                                return;
                                            }

                                            if (Execute1 == true)
                                            {



                                                CircularArc2d Arc2d = Poly_With_arc.GetArcSegment2dAt(Convert.ToInt32(Math.Floor(Param1)));




                                                Point3d pt3 = Poly_With_arc.GetPointAtParameter(Math.Floor(Param1) + 0.5);



                                                Arc Arc1 = new Arc(new Point3d(Arc2d.Center.X, Arc2d.Center.Y, 0), Vector3d.ZAxis, Arc2d.Radius,
                                                                    Functions.GET_Bearing_rad(Arc2d.Center.X, Arc2d.Center.Y, Arc2d.StartPoint.X, Arc2d.StartPoint.Y),
                                                                    Functions.GET_Bearing_rad(Arc2d.Center.X, Arc2d.Center.Y, Arc2d.EndPoint.X, Arc2d.EndPoint.Y));

                                                if (Arc1.GetPointAtDist(Arc1.Length / 2).DistanceTo(pt3) > Arc1.Radius)
                                                {
                                                    Arc1 = new Arc(new Point3d(Arc2d.Center.X, Arc2d.Center.Y, 0), Vector3d.ZAxis, Arc2d.Radius,
                                                                    Functions.GET_Bearing_rad(Arc2d.Center.X, Arc2d.Center.Y, Arc2d.EndPoint.X, Arc2d.EndPoint.Y),
                                                                    Functions.GET_Bearing_rad(Arc2d.Center.X, Arc2d.Center.Y, Arc2d.StartPoint.X, Arc2d.StartPoint.Y));

                                                }

                                                Point3d pt1_0 = new Point3d();
                                                pt1_0 = Arc1.GetClosestPointTo(Pt1, Vector3d.ZAxis, false);

                                                Point3d pt2_0 = new Point3d();
                                                pt2_0 = Arc1.GetClosestPointTo(Pt2, Vector3d.ZAxis, false);


                                                Bulge1 = 1;
                                                double Arc_dist = Math.Abs(Arc1.GetDistAtPoint(pt1_0) - Arc1.GetDistAtPoint(pt2_0));

                                                double Chord = new Point3d(pt1_0.X, pt1_0.Y, 0).DistanceTo(new Point3d(pt2_0.X, pt2_0.Y, 0));
                                                double Bear_rad = Functions.GET_Bearing_rad(pt1_0.X, pt1_0.Y, pt2_0.X, pt2_0.Y);
                                                string Bearing_quadrant = Functions.Get_Quadrant_bearing(Bear_rad, space_in_bear);

                                                Radius1 = Arc1.Radius;
                                                Unghi_la_centru = Functions.Get_DMS((Arc1.TotalAngle) * 180 / Math.PI);
                                                Content1 = textBox_CT_Index.Text;
                                                double d1 = Arc1.GetDistAtPoint(pt1_0);
                                                double d2 = Arc1.GetDistAtPoint(pt2_0);
                                                double d = (d1 + d2) / 2;
                                                Point3d Point_on_arc = Arc1.GetPointAtDist(d);
                                                double tang1 = Math.Abs(Radius1 * Math.Tan(Arc1.TotalAngle / 2));

                                                if (comboBox_coord_system.Text != "")
                                                {



                                                    Pt1_nj = Convert_coordinate_to_new_CS(Pt1);
                                                    Pt2_nj = Convert_coordinate_to_new_CS(Pt2);
                                                    Point3d Pt_cen_nj = Convert_coordinate_to_new_CS(Arc1.Center);
                                                    Point3d Point_on_arc_nj1 = Convert_coordinate_to_new_CS(Point_on_arc);


                                                    double Start1 = Functions.GET_Bearing_rad(Pt_cen_nj.X, Pt_cen_nj.Y, Pt1_nj.X, Pt1_nj.Y);
                                                    double End1 = Functions.GET_Bearing_rad(Pt_cen_nj.X, Pt_cen_nj.Y, Pt2_nj.X, Pt2_nj.Y);

                                                    Chord = new Point3d(Pt1_nj.X, Pt1_nj.Y, 0).DistanceTo(new Point3d(Pt2_nj.X, Pt2_nj.Y, 0));
                                                    Bear_rad = Functions.GET_Bearing_rad(Pt1_nj.X, Pt1_nj.Y, Pt2_nj.X, Pt2_nj.Y);
                                                    Bearing_quadrant = Functions.Get_Quadrant_bearing(Bear_rad, space_in_bear);

                                                    Double R1 = Pt1_nj.GetVectorTo(Pt_cen_nj).Length;


                                                    Arc Arc2;
                                                    Arc2 = new Arc(Pt_cen_nj, Vector3d.ZAxis, R1, Start1, End1);
                                                    Point3d Point_on_arc_nj2 = Arc2.GetPointAtDist(Arc2.Length / 2);

                                                    if (Point_on_arc_nj1.DistanceTo(Point_on_arc_nj2) > R1)
                                                    {
                                                        Arc2 = new Arc(Pt_cen_nj, Vector3d.ZAxis, R1, End1, Start1);
                                                    }

                                                    Arc_dist = Arc2.Length;
                                                    Unghi_la_centru = Functions.Get_DMS((Arc2.TotalAngle) * 180 / Math.PI);

                                                    tang1 = Math.Abs(R1 * Math.Tan(Arc2.TotalAngle / 2));


                                                    TableR.Rows.Add();
                                                    TableR.Rows[TableR.Rows.Count - 1]["CURVE"] = textBox_CT_Index.Text;
                                                    TableR.Rows[TableR.Rows.Count - 1]["RADIUS"] = Functions.Get_String_Rounded(R1, Rounding()) + "'";
                                                    TableR.Rows[TableR.Rows.Count - 1]["ARC_DIST"] = Functions.Get_String_Rounded(Arc_dist, Rounding()) + "'";
                                                    TableR.Rows[TableR.Rows.Count - 1]["DELTA"] = Unghi_la_centru;
                                                    TableR.Rows[TableR.Rows.Count - 1]["DIRECTION"] = Bearing_quadrant;
                                                    TableR.Rows[TableR.Rows.Count - 1]["CHORD"] = Functions.Get_String_Rounded(Chord, Rounding()) + "'";
                                                    TableR.Rows[TableR.Rows.Count - 1]["TANGENT"] = Functions.Get_String_Rounded(tang1, Rounding()) + "'";




                                                }
                                                else
                                                {
                                                    TableR.Rows.Add();
                                                    TableR.Rows[TableR.Rows.Count - 1]["CURVE"] = textBox_CT_Index.Text;
                                                    TableR.Rows[TableR.Rows.Count - 1]["RADIUS"] = Functions.Get_String_Rounded(Radius1, Rounding()) + "'";
                                                    TableR.Rows[TableR.Rows.Count - 1]["ARC_DIST"] = Functions.Get_String_Rounded(Arc_dist, Rounding()) + "'";
                                                    TableR.Rows[TableR.Rows.Count - 1]["DELTA"] = Unghi_la_centru;
                                                    TableR.Rows[TableR.Rows.Count - 1]["DIRECTION"] = Bearing_quadrant;
                                                    TableR.Rows[TableR.Rows.Count - 1]["CHORD"] = Functions.Get_String_Rounded(Chord, Rounding()) + "'";
                                                    TableR.Rows[TableR.Rows.Count - 1]["TANGENT"] = Functions.Get_String_Rounded(tang1, Rounding()) + "'";
                                                }

                                                Bulge1 = 1;

                                            }
                                            else
                                            {
                                                Freeze_operations = false;
                                                MessageBox.Show("between the picked points there is another vertex");
                                                return;
                                            }
                                        }
                                        if (Ent1 is Arc)
                                        {
                                            Arc Arc1 = (Arc)Ent1;

                                            Bulge1 = 1;
                                            double Arc_dist = Math.Abs(Arc1.GetDistAtPoint(Pt1) - Arc1.GetDistAtPoint(Pt2));

                                            double Chord = new Point3d(Pt1.X, Pt1.Y, 0).DistanceTo(new Point3d(Pt2.X, Pt2.Y, 0));
                                            double Bear_rad = Functions.GET_Bearing_rad(Pt1.X, Pt1.Y, Pt2.X, Pt2.Y);
                                            string Bearing_quadrant = Functions.Get_Quadrant_bearing(Bear_rad, space_in_bear);

                                            Radius1 = Arc1.Radius;
                                            Unghi_la_centru = Functions.Get_DMS((Arc1.TotalAngle) * 180 / Math.PI);
                                            Content1 = textBox_CT_Index.Text;
                                            double d1 = Arc1.GetDistAtPoint(Pt1);
                                            double d2 = Arc1.GetDistAtPoint(Pt2);
                                            double d = (d1 + d2) / 2;
                                            Point3d Point_on_arc = Arc1.GetPointAtDist(d);
                                            double tang1 = Math.Abs(Radius1 * Math.Tan(Arc1.TotalAngle / 2));


                                            if (comboBox_coord_system.Text != "")
                                            {
                                                Pt1_nj = Convert_coordinate_to_new_CS(Pt1);
                                                Pt2_nj = Convert_coordinate_to_new_CS(Pt2);
                                                Point3d Pt_cen_nj = Convert_coordinate_to_new_CS(Arc1.Center);
                                                Point3d Point_on_arc_nj1 = Convert_coordinate_to_new_CS(Point_on_arc);


                                                double Start1 = Functions.GET_Bearing_rad(Pt_cen_nj.X, Pt_cen_nj.Y, Pt1_nj.X, Pt1_nj.Y);
                                                double End1 = Functions.GET_Bearing_rad(Pt_cen_nj.X, Pt_cen_nj.Y, Pt2_nj.X, Pt2_nj.Y);

                                                Chord = new Point3d(Pt1_nj.X, Pt1_nj.Y, 0).DistanceTo(new Point3d(Pt2_nj.X, Pt2_nj.Y, 0));
                                                Bear_rad = Functions.GET_Bearing_rad(Pt1_nj.X, Pt1_nj.Y, Pt2_nj.X, Pt2_nj.Y);
                                                Bearing_quadrant = Functions.Get_Quadrant_bearing(Bear_rad,space_in_bear);

                                                Double R1 = Pt1_nj.GetVectorTo(Pt_cen_nj).Length;


                                                Arc Arc2;
                                                Arc2 = new Arc(Pt_cen_nj, Vector3d.ZAxis, R1, Start1, End1);
                                                Point3d Point_on_arc_nj2 = Arc2.GetPointAtDist(Arc2.Length / 2);

                                                if (Point_on_arc_nj1.DistanceTo(Point_on_arc_nj2) > R1)
                                                {
                                                    Arc2 = new Arc(Pt_cen_nj, Vector3d.ZAxis, R1, End1, Start1);
                                                }

                                                Arc_dist = Arc2.Length;
                                                Unghi_la_centru = Functions.Get_DMS((Arc2.TotalAngle) * 180 / Math.PI);

                                                tang1 = Math.Abs(R1 * Math.Tan(Arc2.TotalAngle / 2));

                                                TableR.Rows.Add();
                                                TableR.Rows[TableR.Rows.Count - 1]["CURVE"] = textBox_CT_Index.Text;
                                                TableR.Rows[TableR.Rows.Count - 1]["RADIUS"] = Functions.Get_String_Rounded(R1, Rounding()) + "'";
                                                TableR.Rows[TableR.Rows.Count - 1]["ARC_DIST"] = Functions.Get_String_Rounded(Arc_dist, Rounding()) + "'";
                                                TableR.Rows[TableR.Rows.Count - 1]["DELTA"] = Unghi_la_centru;
                                                TableR.Rows[TableR.Rows.Count - 1]["DIRECTION"] = Bearing_quadrant;
                                                TableR.Rows[TableR.Rows.Count - 1]["CHORD"] = Functions.Get_String_Rounded(Chord, Rounding()) + "'";
                                                TableR.Rows[TableR.Rows.Count - 1]["TANGENT"] = Functions.Get_String_Rounded(tang1, Rounding()) + "'";
                                            }
                                            else
                                            {
                                                TableR.Rows.Add();
                                                TableR.Rows[TableR.Rows.Count - 1]["CURVE"] = textBox_CT_Index.Text;
                                                TableR.Rows[TableR.Rows.Count - 1]["RADIUS"] = Functions.Get_String_Rounded(Radius1, Rounding()) + "'";
                                                TableR.Rows[TableR.Rows.Count - 1]["ARC_DIST"] = Functions.Get_String_Rounded(Arc_dist, Rounding()) + "'";
                                                TableR.Rows[TableR.Rows.Count - 1]["DELTA"] = Unghi_la_centru;
                                                TableR.Rows[TableR.Rows.Count - 1]["DIRECTION"] = Bearing_quadrant;
                                                TableR.Rows[TableR.Rows.Count - 1]["CHORD"] = Functions.Get_String_Rounded(Chord, Rounding()) + "'";
                                                TableR.Rows[TableR.Rows.Count - 1]["TANGENT"] = Functions.Get_String_Rounded(tang1, Rounding()) + "'";
                                            }

                                        }


                                    }
                                }


                                double x11 = x1;
                                double x22 = x2;
                                double y11 = y1;
                                double y22 = y2;

                                if (Bulge1 == 0)
                                {

                                    Pt1 = Point1;
                                    Pt2 = Point2;
                                    x1 = Pt1.X;
                                    y1 = Pt1.Y;
                                    x2 = Pt2.X;
                                    y2 = Pt2.Y;

                                    Point3d Pt11 = Point1;
                                    Point3d Pt22 = Point2;
                                    x11 = x1;
                                    x22 = x2;
                                    y11 = y1;
                                    y22 = y2;

                                    if (comboBox_Scale.Text == "PSpace")
                                    {

                                        Viewport Vp1 = Trans1.GetObject(Ent_vp_id, OpenMode.ForRead) as Viewport;
                                        if (Vp1 != null)
                                        {
                                            Matrix3d TransforMatrix = Functions.PaperToModel(Vp1);
                                            Pt11 = Point1.TransformBy(TransforMatrix);
                                            Pt22 = Point2.TransformBy(TransforMatrix);
                                            x11 = Pt11.X;
                                            y11 = Pt11.Y;
                                            x22 = Pt22.X;
                                            y22 = Pt22.Y;
                                        }

                                    }




                                    Dist1 = Math.Pow(Math.Pow(x11 - x22, 2) + Math.Pow(y11 - y22, 2), 0.5);


                                    if (comboBox_coord_system.Text != "")
                                    {
                                        Pt1_nj = Convert_coordinate_to_new_CS(Pt11);
                                        Pt2_nj = Convert_coordinate_to_new_CS(Pt22);

                                        Dist1 = Math.Pow(Math.Pow(Pt1_nj.X - Pt2_nj.X, 2) + Math.Pow(Pt1_nj.Y - Pt2_nj.Y, 2), 0.5);
                                    }


                                }

                                double Bearing1 = Functions.GET_Bearing_rad(x11, y11, x22, y22);
                                double Bearing2 = Functions.GET_Bearing_rad(x1, y1, x2, y2);

                                double Rot_t = Functions.calculate_rotatie_text(Bearing2);



                                string Quadrant1 = Functions.Get_Quadrant_bearing(Bearing1, space_in_bear);

                                Point3d PointM = new Point3d((x1 + x2) / 2, (y1 + y2) / 2, 0);

                                Point3d Point_ins = get_insertion_pt(PointM, new Point3d(x2, y2, 0), Bearing2, TextH);

                                double BearingNJ = 0;
                                if (comboBox_coord_system.Text != "")
                                {


                                    BearingNJ = Functions.GET_Bearing_rad(Pt1_nj.X, Pt1_nj.Y, Pt2_nj.X, Pt2_nj.Y);

                                    Quadrant1 = Functions.Get_Quadrant_bearing(BearingNJ, space_in_bear);

                                }


                                if (Bulge1 == 0)
                                {
                                    if (comboBox_Label_Position.Text == "Top")
                                    {
                                        if (radioButton_BD.Checked == true)
                                        {
                                            if (checkBox_thousand_sep.Checked == false) Content1 = Functions.Get_String_Rounded(Dist1, Rounding()) + "'" + suffix1() + "\\P" + Quadrant1;

                                            if (checkBox_thousand_sep.Checked == true) Content1 = Functions.Get_String_Rounded_with_thousand_sep(Dist1, Rounding()) + "'" + suffix1() + "\\P" + Quadrant1;


                                        }
                                        if (radioButton_B.Checked == true)
                                        {
                                            Content1 = Quadrant1;
                                        }
                                        if (radioButton_D.Checked == true)
                                        {
                                            if (checkBox_thousand_sep.Checked == false) Content1 = Functions.Get_String_Rounded(Dist1, Rounding()) + "'" + suffix1();

                                            if (checkBox_thousand_sep.Checked == true) Content1 = Functions.Get_String_Rounded_with_thousand_sep(Dist1, Rounding()) + "'" + suffix1();

                                        }

                                    }
                                    else if (comboBox_Label_Position.Text == "Bottom")
                                    {
                                        if (radioButton_BD.Checked == true)
                                        {
                                            if (checkBox_thousand_sep.Checked == false) Content1 = Quadrant1 + "\\P" + Functions.Get_String_Rounded(Dist1, Rounding()) + "'" + suffix1();

                                            if (checkBox_thousand_sep.Checked == true) Content1 = Quadrant1 + "\\P" + Functions.Get_String_Rounded_with_thousand_sep(Dist1, Rounding()) + "'" + suffix1();

                                        }
                                        if (radioButton_B.Checked == true)
                                        {
                                            Content1 = Quadrant1;
                                        }
                                        if (radioButton_D.Checked == true)
                                        {
                                            if (checkBox_thousand_sep.Checked == false) Content1 = Functions.Get_String_Rounded(Dist1, Rounding()) + "'" + suffix1();
                                            if (checkBox_thousand_sep.Checked == true) Content1 = Functions.Get_String_Rounded_with_thousand_sep(Dist1, Rounding()) + "'" + suffix1();
                                        }

                                    }
                                    else
                                    {
                                        if (radioButton_BD.Checked == true)
                                        {
                                            if (checkBox_thousand_sep.Checked == false) Content1 = Quadrant1 + "\\P" + Functions.Get_String_Rounded(Dist1, Rounding()) + "'" + suffix1();
                                            if (checkBox_thousand_sep.Checked == true) Content1 = Quadrant1 + "\\P" + Functions.Get_String_Rounded_with_thousand_sep(Dist1, Rounding()) + "'" + suffix1();
                                        }
                                        if (radioButton_B.Checked == true)
                                        {
                                            Content1 = Quadrant1;
                                        }
                                        if (radioButton_D.Checked == true)
                                        {
                                            if (checkBox_thousand_sep.Checked == false) Content1 = Functions.Get_String_Rounded(Dist1, Rounding()) + "'" + suffix1();
                                            if (checkBox_thousand_sep.Checked == true) Content1 = Functions.Get_String_Rounded_with_thousand_sep(Dist1, Rounding()) + "'" + suffix1();
                                        }

                                    }
                                }

                                if (radioButton_LT.Checked == true & textBox_LT_Index.Text != "")
                                {
                                    TableL.Rows.Add();
                                    TableL.Rows[TableL.Rows.Count - 1]["LINE"] = textBox_LT_Index.Text;
                                    TableL.Rows[TableL.Rows.Count - 1]["BEARING"] = Quadrant1;
                                    if (checkBox_thousand_sep.Checked == false) TableL.Rows[TableL.Rows.Count - 1]["DISTANCE"] = Functions.Get_String_Rounded(Dist1, Rounding()) + "'";
                                    if (checkBox_thousand_sep.Checked == true) TableL.Rows[TableL.Rows.Count - 1]["DISTANCE"] = Functions.Get_String_Rounded_with_thousand_sep(Dist1, Rounding()) + "'";
                                    Content1 = textBox_LT_Index.Text;
                                }

                                Mtext1.Rotation = Rot_t;



                                if (comboBox_Label_Position.Text == "Curved leader")
                                {
                                    Point3d P1 = Point1;
                                    Point3d P2 = Point1;
                                    Point3d P3 = Point1;

                                    String Continut = "";
                                    if (radioButton_LT.Checked == true)
                                    {
                                        Continut = textBox_LT_Index.Text;
                                    }
                                    if (radioButton_CT.Checked == true)
                                    {
                                        Continut = textBox_CT_Index.Text;
                                    }
                                    if (radioButton_BD.Checked == true | radioButton_B.Checked == true | radioButton_D.Checked == true)
                                    {
                                        Continut = Content1;
                                    }

                                    if (Continut != "")
                                    {
                                        Jig_Class1 Jig1 = new Jig_Class1();
                                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PA1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPick Curved Leader first Point : ");
                                        Autodesk.AutoCAD.EditorInput.PromptPointResult Pt_l1;

                                        PA1.AllowNone = false;
                                        Pt_l1 = Editor1.GetPoint(PA1);

                                        double Arrow1 = 62.5 * Calculate_BLOCK_SCALE();

                                        if (Pt_l1.Status == PromptStatus.OK)
                                        {
                                            P1 = Pt_l1.Value;
                                            Autodesk.AutoCAD.EditorInput.PromptPointResult Pt_l2;
                                            Pt_l2 = Jig1.StartJig(Pt_l1.Value, Arrow1);
                                            if (Pt_l2.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                            {
                                                P2 = Pt_l2.Value;
                                                Jig_Class2 Jig2 = new Jig_Class2();
                                                Autodesk.AutoCAD.EditorInput.PromptPointResult Pt_l3;
                                                Pt_l3 = Jig2.StartJig(Pt_l1.Value, Pt_l2.Value, Arrow1);
                                                if (Pt_l3.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                                {
                                                    P3 = Pt_l3.Value;
                                                }
                                            }
                                        }

                                        if (P3 != Point1)
                                        {
                                            Double Wdth1 = Arrow1 / 1250;
                                            Jig_Class2 xx = new Jig_Class2();
                                            Double XX1 = xx.X_Y_for_arc_leader(P1.X, P1.Y, P2.X, P2.Y, Arrow1, "X");
                                            Double YY1 = xx.X_Y_for_arc_leader(P1.X, P1.Y, P2.X, P2.Y, Arrow1, "Y");
                                            Double BulgeX = xx.Bulge_for_arc_leader(P1.X, P1.Y, P2.X, P2.Y, P3.X, P3.Y, Arrow1);
                                            Polyline ArcL = new Autodesk.AutoCAD.DatabaseServices.Polyline();

                                            ArcL.AddVertexAt(0, new Point2d(P1.X, P1.Y), 0, 0, Wdth1);
                                            ArcL.AddVertexAt(1, new Point2d(XX1, YY1), BulgeX, 0, 0);
                                            ArcL.AddVertexAt(2, new Point2d(P3.X, P3.Y), 0, 0, 0);
                                            BTrecord.AppendEntity(ArcL);
                                            Trans1.AddNewlyCreatedDBObject(ArcL, true);
                                            Trans1.TransactionManager.QueueForGraphicsFlush();

                                            Mtext1.Rotation = 0;

                                            Jig_Class3 Jig3 = new Jig_Class3();
                                            Autodesk.AutoCAD.EditorInput.PromptPointResult Pt_txt;
                                            Pt_txt = Jig3.StartJig(TextH, Continut);

                                            Point_ins = Pt_txt.Value;

                                        }
                                    }

                                }

                                Mtext1.Attachment = Get_mTEXT_attachment();
                                Mtext1.Contents = Content1;
                                Mtext1.TextHeight = TextH;
                                Mtext1.Location = Point_ins;

                                if (background_fill() == true)
                                {
                                    Mtext1.BackgroundFill = true;
                                    Mtext1.UseBackgroundColor = true;
                                    Mtext1.BackgroundScaleFactor = 1.2;
                                }
                                BTrecord.AppendEntity(Mtext1);
                                Trans1.AddNewlyCreatedDBObject(Mtext1, true);
                            }

                            if (radioButton_NE.Checked == true)
                            {

                                object OLD_OSnap = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("OSMODE");

                                object new_OSnap = 33;

                                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", new_OSnap);

                                if (comboBox_Label_Position.Text == "Leader")
                                {


                                    //End = 1,
                                    //Middle = 2,
                                    //Center = 4,
                                    //Node = 8,
                                    //Quadrant = 16,
                                    //Intersection = 32,
                                    //Insertion = 64,
                                    //Perpendicular = 128,
                                    //Tangent = 256,
                                    //Near = 512,
                                    // Quick = 1024,
                                    //ApparentIntersection = 2048,
                                    //Immediate = 65536,
                                    //AllowTangent = 131072,
                                    // DisablePerpendicular = 262144,
                                    //RelativeCartesian = 524288,
                                    //RelativePolar = 1048576,
                                    //NoneOverride = 2097152,    


                                    ObjectId Arrowid = Get_Arrow_dimension_ID("DIMBLK2", "_DotSmall");


                                    double Landinggap = Calculate_BLOCK_SCALE() / 20;
                                    double Doglength1 = Calculate_BLOCK_SCALE() / 10;
                                    double Texth = Calculate_text_height();
                                    double Arrowsize1 = Calculate_BLOCK_SCALE() / 5;

                                    Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                    Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                    PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the first point");
                                    PP1.AllowNone = false;
                                    Point_res1 = Editor1.GetPoint(PP1);

                                    if (Point_res1.Status != PromptStatus.OK)
                                    {
                                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                        Freeze_operations = false;
                                        Trans1.Commit();
                                        return;
                                    }


                                    Point3d Point1 = new Point3d();
                                    Point1 = Point_res1.Value;

                                    Point3d PointT1 = new Point3d();
                                    PointT1 = Point_res1.Value;


                                    Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                                    Autodesk.AutoCAD.EditorInput.PromptPointOptions PP2;
                                    PP2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the second point");
                                    PP2.AllowNone = false;
                                    PP2.BasePoint = Point1;
                                    PP2.UseBasePoint = true;
                                    Point_res2 = Editor1.GetPoint(PP2);

                                    if (Point_res2.Status != PromptStatus.OK)
                                    {

                                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                        Freeze_operations = false;
                                        Trans1.Commit();
                                        return;
                                    }
                                    Point3d Point2 = new Point3d();
                                    Point2 = Point_res2.Value;


                                    MLeader Mleader1 = new MLeader();
                                    int Nr1 = Mleader1.AddLeader();
                                    int Nr2 = Mleader1.AddLeaderLine(Nr1);
                                    Mleader1.AddFirstVertex(Nr2, Point1);
                                    Mleader1.AddLastVertex(Nr2, Point2);
                                    Mleader1.LeaderLineType = LeaderType.StraightLeader;

                                    Mleader1.ContentType = ContentType.MTextContent;




                                    Mleader1.SetTextAttachmentType(TextAttachmentType.AttachmentMiddleOfTop, LeaderDirectionType.LeftLeader);
                                    Mleader1.SetTextAttachmentType(TextAttachmentType.AttachmentMiddleOfTop, LeaderDirectionType.RightLeader);
                                    Mleader1.Annotative = AnnotativeStates.False;


                                    if (comboBox_Scale.Text == "PSpace")
                                    {

                                        Viewport Vp1 = Trans1.GetObject(Ent_vp_id, OpenMode.ForRead) as Viewport;
                                        if (Vp1 != null)
                                        {
                                            Matrix3d TransforMatrix = Functions.PaperToModel(Vp1);

                                            PointT1 = Point1.TransformBy(TransforMatrix);


                                        }

                                    }


                                    MText Mtext1 = new MText();

                                    if (checkBox_thousand_sep.Checked == false) Mtext1.Contents = "N:" + Functions.Get_String_Rounded(PointT1.Y, Rounding()) + "'\r\nE:" + Functions.Get_String_Rounded(PointT1.X, Rounding()) + "'";
                                    if (checkBox_thousand_sep.Checked == true) Mtext1.Contents = "N:" + Functions.Get_String_Rounded_with_thousand_sep(PointT1.Y, Rounding()) + "'\r\nE:" + Functions.Get_String_Rounded_with_thousand_sep(PointT1.X, Rounding()) + "'";
                                    Mtext1.ColorIndex = 0;

                                    if (background_fill() == true)
                                    {
                                        Mtext1.BackgroundFill = true;
                                        Mtext1.UseBackgroundColor = true;
                                        Mtext1.BackgroundScaleFactor = 1.2;
                                    }


                                    Mleader1.MText = Mtext1;


                                    Mleader1.TextHeight = Texth;
                                    Mleader1.ArrowSymbolId = Arrowid;
                                    Mleader1.LandingGap = Landinggap;
                                    Mleader1.ArrowSize = Arrowsize1;
                                    Mleader1.DoglegLength = Doglength1;




                                    BTrecord.AppendEntity(Mleader1);
                                    Trans1.AddNewlyCreatedDBObject(Mleader1, true);

                                }

                                if (comboBox_Label_Position.Text == "Curved leader")
                                {

                                    Point3d Point1 = new Point3d();

                                    Jig_Class1 Jig1 = new Jig_Class1();
                                    Autodesk.AutoCAD.EditorInput.PromptPointOptions PA1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPick Curved Leader first Point : ");
                                    Autodesk.AutoCAD.EditorInput.PromptPointResult Pt_l1;

                                    PA1.AllowNone = false;
                                    Pt_l1 = Editor1.GetPoint(PA1);

                                    if (Pt_l1.Status != PromptStatus.OK)
                                    {
                                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                        Freeze_operations = false;
                                        Trans1.Commit();
                                        return;
                                    }


                                    double Arrow1 = 62.5 * Calculate_BLOCK_SCALE();

                                    if (Pt_l1.Status == PromptStatus.OK)
                                    {
                                        Point1 = Pt_l1.Value;

                                        Point3d PointT1 = new Point3d();
                                        PointT1 = Pt_l1.Value;

                                        Autodesk.AutoCAD.EditorInput.PromptPointResult Pt_l2;
                                        Pt_l2 = Jig1.StartJig(Pt_l1.Value, Arrow1);

                                        if (Pt_l2.Status != PromptStatus.OK)
                                        {
                                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                            Freeze_operations = false;
                                            Trans1.Commit();
                                            return;
                                        }
                                        if (Pt_l2.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                        {
                                            ;
                                            Jig_Class2 Jig2 = new Jig_Class2();
                                            Autodesk.AutoCAD.EditorInput.PromptPointResult Pt_l3;
                                            Pt_l3 = Jig2.StartJig(Pt_l1.Value, Pt_l2.Value, Arrow1);

                                            if (Pt_l3.Status != PromptStatus.OK)
                                            {
                                                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                Freeze_operations = false;
                                                Trans1.Commit();
                                                return;
                                            }

                                            if (Pt_l3.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                            {

                                                if (comboBox_Scale.Text == "PSpace")
                                                {

                                                    Viewport Vp1 = Trans1.GetObject(Ent_vp_id, OpenMode.ForRead) as Viewport;
                                                    if (Vp1 != null)
                                                    {
                                                        Matrix3d TransforMatrix = Functions.PaperToModel(Vp1);

                                                        PointT1 = Point1.TransformBy(TransforMatrix);


                                                    }

                                                }

                                                string Continut = "N:" + Functions.Get_String_Rounded(PointT1.Y, Rounding()) + "'\r\nE:" + Functions.Get_String_Rounded(PointT1.X, Rounding()) + "'";
                                                if (checkBox_thousand_sep.Checked == true) Continut = "N:" + Functions.Get_String_Rounded_with_thousand_sep(PointT1.Y, Rounding()) + "'\r\nE:" + Functions.Get_String_Rounded_with_thousand_sep(PointT1.X, Rounding()) + "'";

                                                Double Wdth1 = Arrow1 / 1250;
                                                Jig_Class2 xx = new Jig_Class2();
                                                Double XX1 = xx.X_Y_for_arc_leader(Point1.X, Point1.Y, Pt_l2.Value.X, Pt_l2.Value.Y, Arrow1, "X");
                                                Double YY1 = xx.X_Y_for_arc_leader(Point1.X, Point1.Y, Pt_l2.Value.X, Pt_l2.Value.Y, Arrow1, "Y");
                                                Double BulgeX = xx.Bulge_for_arc_leader(Point1.X, Point1.Y, Pt_l2.Value.X, Pt_l2.Value.Y, Pt_l3.Value.X, Pt_l3.Value.Y, Arrow1);
                                                Polyline ArcL = new Autodesk.AutoCAD.DatabaseServices.Polyline();

                                                ArcL.AddVertexAt(0, new Point2d(Point1.X, Point1.Y), 0, 0, Wdth1);
                                                ArcL.AddVertexAt(1, new Point2d(XX1, YY1), BulgeX, 0, 0);
                                                ArcL.AddVertexAt(2, new Point2d(Pt_l3.Value.X, Pt_l3.Value.Y), 0, 0, 0);
                                                BTrecord.AppendEntity(ArcL);
                                                Trans1.AddNewlyCreatedDBObject(ArcL, true);
                                                Trans1.TransactionManager.QueueForGraphicsFlush();



                                                Jig_Class3 Jig3 = new Jig_Class3();
                                                Autodesk.AutoCAD.EditorInput.PromptPointResult Pt_txt;
                                                Pt_txt = Jig3.StartJig(Calculate_text_height(), Continut);

                                                if (Pt_txt.Status != PromptStatus.OK)
                                                {
                                                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                    Freeze_operations = false;
                                                    Trans1.Commit();
                                                    return;
                                                }

                                                if (Pt_txt.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                                {
                                                    MText Mtext1 = new MText();
                                                    Mtext1.Attachment = Get_mTEXT_attachment();
                                                    Mtext1.Contents = Continut;
                                                    Mtext1.TextHeight = Calculate_text_height();
                                                    Mtext1.Location = Pt_txt.Value;
                                                    Mtext1.Rotation = 0;

                                                    if (background_fill() == true)
                                                    {
                                                        Mtext1.BackgroundFill = true;
                                                        Mtext1.UseBackgroundColor = true;
                                                        Mtext1.BackgroundScaleFactor = 1.2;
                                                    }

                                                    BTrecord.AppendEntity(Mtext1);
                                                    Trans1.AddNewlyCreatedDBObject(Mtext1, true);

                                                }


                                            }
                                        }
                                    }

                                }


                            }

                            if (radioButton_PI.Checked == true)
                            {
                                if (BlockTable1.Has("PI_BLOCK") == false)
                                {
                                    MessageBox.Show("PI_BLOCK is not inserted on this drawing \r\nPlease contact HECTOR MORALES!");
                                    Freeze_operations = false;
                                    return;

                                }


                                Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_polyline;
                                Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_polyline;
                                Prompt_polyline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the polyline:");
                                Prompt_polyline.SetRejectMessage("\nSelect a polyline!");
                                Prompt_polyline.AllowNone = true;
                                Prompt_polyline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                                Rezultat_polyline = ThisDrawing.Editor.GetEntity(Prompt_polyline);

                                if (Rezultat_polyline.Status != PromptStatus.OK)
                                {
                                    Freeze_operations = false;
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    return;
                                }

                                Polyline Poly2 = (Polyline)Trans1.GetObject(Rezultat_polyline.ObjectId, OpenMode.ForRead);
                                for (int i = 0; i < Poly2.NumberOfVertices; ++i)
                                {
                                    Functions.InsertBlock_with_multiple_atributes("", "PI_BLOCK", Poly2.GetPointAtParameter(i), Calculate_BLOCK_SCALE(), "0", new System.Collections.Specialized.StringCollection(), new System.Collections.Specialized.StringCollection());
                                }
                            }

                            if (radioButton_draw_arcL.Checked == true)
                            {

                                Point3d Point1 = new Point3d();

                                Jig_Class1 Jig1 = new Jig_Class1();
                                Autodesk.AutoCAD.EditorInput.PromptPointOptions PA1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPick Curved Leader first Point : ");
                                Autodesk.AutoCAD.EditorInput.PromptPointResult Pt_l1;

                                PA1.AllowNone = false;
                                Pt_l1 = Editor1.GetPoint(PA1);

                                if (Pt_l1.Status != PromptStatus.OK)
                                {

                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    Freeze_operations = false;
                                    Trans1.Commit();
                                    return;
                                }


                                double Arrow1 = 62.5 * Calculate_BLOCK_SCALE();

                                if (Pt_l1.Status == PromptStatus.OK)
                                {
                                    Point1 = Pt_l1.Value;
                                    Autodesk.AutoCAD.EditorInput.PromptPointResult Pt_l2;
                                    Pt_l2 = Jig1.StartJig(Pt_l1.Value, Arrow1);

                                    if (Pt_l2.Status != PromptStatus.OK)
                                    {

                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                        Freeze_operations = false;
                                        Trans1.Commit();
                                        return;
                                    }
                                    if (Pt_l2.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                    {
                                        ;
                                        Jig_Class2 Jig2 = new Jig_Class2();
                                        Autodesk.AutoCAD.EditorInput.PromptPointResult Pt_l3;
                                        Pt_l3 = Jig2.StartJig(Pt_l1.Value, Pt_l2.Value, Arrow1);

                                        if (Pt_l3.Status != PromptStatus.OK)
                                        {
                                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                            Freeze_operations = false;
                                            Trans1.Commit();
                                            return;
                                        }

                                        if (Pt_l3.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                        {
                                            Double Wdth1 = Arrow1 / 1250;
                                            Jig_Class2 xx = new Jig_Class2();
                                            Double XX1 = xx.X_Y_for_arc_leader(Point1.X, Point1.Y, Pt_l2.Value.X, Pt_l2.Value.Y, Arrow1, "X");
                                            Double YY1 = xx.X_Y_for_arc_leader(Point1.X, Point1.Y, Pt_l2.Value.X, Pt_l2.Value.Y, Arrow1, "Y");
                                            Double BulgeX = xx.Bulge_for_arc_leader(Point1.X, Point1.Y, Pt_l2.Value.X, Pt_l2.Value.Y, Pt_l3.Value.X, Pt_l3.Value.Y, Arrow1);
                                            Polyline ArcL = new Autodesk.AutoCAD.DatabaseServices.Polyline();

                                            ArcL.AddVertexAt(0, new Point2d(Point1.X, Point1.Y), 0, 0, Wdth1);
                                            ArcL.AddVertexAt(1, new Point2d(XX1, YY1), BulgeX, 0, 0);
                                            ArcL.AddVertexAt(2, new Point2d(Pt_l3.Value.X, Pt_l3.Value.Y), 0, 0, 0);
                                            BTrecord.AppendEntity(ArcL);
                                            Trans1.AddNewlyCreatedDBObject(ArcL, true);
                                            Trans1.Commit();
                                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                            Freeze_operations = false;
                                            return;

                                        }
                                    }
                                }

                            }

                            if (radioButton_tie_acholade.Checked == true)
                            {


                                double Arrow1 = Calculate_BLOCK_SCALE() * 1000 * 0.0625;
                                double TextH = Calculate_text_height();
                                double Text_rotation = 0;
                                double Arr_len = Arrow1 / 500;
                                double Arr_width = Arrow1 / 1250;
                                double rad_small = Arrow1 / 250;
                                object OLD_OSnap = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("OSMODE");

                                object new_OSnap = 33;

                                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", new_OSnap);
                                Matrix3d Current_UCS_matrix = Editor1.CurrentUserCoordinateSystem;



                                //End = 1,
                                //Middle = 2,
                                //Center = 4,
                                //Node = 8,
                                //Quadrant = 16,
                                //Intersection = 32,
                                //Insertion = 64,
                                //Perpendicular = 128,
                                //Tangent = 256,
                                //Near = 512,
                                // Quick = 1024,
                                //ApparentIntersection = 2048,
                                //Immediate = 65536,
                                //AllowTangent = 131072,
                                // DisablePerpendicular = 262144,
                                //RelativeCartesian = 524288,
                                //RelativePolar = 1048576,
                                //NoneOverride = 2097152,    





                                PromptPointOptions PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPick 1st point:");
                                PromptPointResult Point1;
                                PP1.AllowNone = true;
                                Point1 = Editor1.GetPoint(PP1);
                                if (Point1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                {
                                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                                    Freeze_operations = false;
                                    return;
                                }

                                PromptPointOptions PP2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPick 2nd point:");
                                PromptPointResult Point2;
                                PP2.AllowNone = true;
                                PP2.BasePoint = Point1.Value;
                                PP2.UseBasePoint = true;
                                Point2 = Editor1.GetPoint(PP2);
                                if (Point2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                {
                                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                                    Freeze_operations = false;
                                    return;
                                }


                                if (Point1.Value.GetVectorTo(Point2.Value).Length < 2 * Arr_len)
                                {
                                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                                    Freeze_operations = false;
                                    return;
                                }

                                Acholade_Jig_Class JigAcolade = new Acholade_Jig_Class();


                                PromptPointResult Point3 = JigAcolade.StartJig(Point1.Value.TransformBy(Current_UCS_matrix), Point2.Value.TransformBy(Current_UCS_matrix), Arr_len, Arr_width, rad_small);
                                if (Point3 == null)
                                {
                                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                                    Freeze_operations = false;
                                    return;
                                }
                                if (Point3.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                {
                                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                                    Freeze_operations = false;
                                    return;
                                }


                                CircularArc3d Arc3d = new CircularArc3d(Point1.Value.TransformBy(Current_UCS_matrix), Point2.Value.TransformBy(Current_UCS_matrix), Point3.Value);
                                Circle Circle1 = new Circle(Arc3d.Center, Vector3d.ZAxis, Point1.Value.TransformBy(Current_UCS_matrix).GetVectorTo(Arc3d.Center).Length);


                                Line Line1 = new Line(Circle1.Center, Point3.Value);
                                double Scale_f = (Math.Pow(((Circle1.Radius + rad_small) * (Circle1.Radius + rad_small) - rad_small * rad_small), 0.5)) / Circle1.Radius;
                                Line1.TransformBy(Matrix3d.Scaling(Scale_f, Circle1.Center));




                                Point3d PointT = Line1.EndPoint;

                                Point3d PointA = Line1.GetPointAtDist(Line1.Length - rad_small);
                                Line LinieR = new Line(PointT, PointA);


                                Line LinieL = (Line)LinieR.Clone();
                                LinieL.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, PointT));
                                LinieR.TransformBy(Matrix3d.Rotation(-Math.PI / 2, Vector3d.ZAxis, PointT));


                                Circle Circle2 = new Circle(LinieR.EndPoint, Vector3d.ZAxis, rad_small);

                                Circle Circle3 = new Circle(LinieL.EndPoint, Vector3d.ZAxis, rad_small);

                                Point3d PtI1 = new Point3d();
                                Point3d PtI2 = new Point3d();

                                Point3dCollection Colint1 = new Point3dCollection();
                                Circle1.IntersectWith(Circle2, Intersect.OnBothOperands, Colint1, IntPtr.Zero, IntPtr.Zero);
                                if (Colint1.Count > 0)
                                {
                                    PtI1 = Colint1[0];
                                }

                                Point3dCollection Colint2 = new Point3dCollection();
                                Circle1.IntersectWith(Circle3, Intersect.OnBothOperands, Colint2, IntPtr.Zero, IntPtr.Zero);
                                if (Colint2.Count > 0)
                                {
                                    PtI2 = Colint2[0];
                                }

                                if (Colint1.Count > 0 && Colint2.Count > 0)
                                {
                                    if (PtI1 != null && PtI2 != null)
                                    {
                                        Double AngleStart = Functions.GET_Bearing_rad(Circle2.Center.X, Circle2.Center.Y, PtI1.X, PtI1.Y);
                                        Double AngleEnd = Functions.GET_Bearing_rad(Circle2.Center.X, Circle2.Center.Y, PointT.X, PointT.Y);
                                        Arc Arc1 = new Arc(Circle2.Center, rad_small, AngleStart, AngleEnd);

                                        AngleStart = Functions.GET_Bearing_rad(Circle3.Center.X, Circle3.Center.Y, PtI2.X, PtI2.Y);
                                        AngleEnd = Functions.GET_Bearing_rad(Circle3.Center.X, Circle3.Center.Y, PointT.X, PointT.Y);
                                        Arc Arc2 = new Arc(Circle3.Center, rad_small, AngleEnd, AngleStart);

                                        Point3d PointB3 = Point1.Value.TransformBy(Current_UCS_matrix);
                                        Point3d PointB4 = Point2.Value.TransformBy(Current_UCS_matrix);

                                        if (PointB3.GetVectorTo(Circle2.Center).Length < PointB3.GetVectorTo(Circle3.Center).Length)
                                        {
                                            Point3d T = new Point3d();
                                            T = PointB3;
                                            PointB3 = PointB4;
                                            PointB4 = T;
                                        }

                                        AngleStart = Functions.GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PtI2.X, PtI2.Y);
                                        AngleEnd = Functions.GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PointB3.X, PointB3.Y);
                                        Arc Arc3 = new Arc(Circle1.Center, Circle1.Radius, AngleEnd, AngleStart);


                                        AngleStart = Functions.GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PtI1.X, PtI1.Y);
                                        AngleEnd = Functions.GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PointB4.X, PointB4.Y);
                                        Arc Arc4 = new Arc(Circle1.Center, Circle1.Radius, AngleStart, AngleEnd);

                                        Polyline Poly123 = new Polyline();


                                        Double b0 = Math.Tan(Arc3.TotalAngle / 4);
                                        Double b1 = -Math.Tan(Arc2.TotalAngle / 4);
                                        Double b2 = -Math.Tan(Arc1.TotalAngle / 4);
                                        Double b3 = Math.Tan(Arc4.TotalAngle / 4);

                                        if (Arc3.Length > Arr_len)
                                        {

                                            Point3d PtArr3 = Arc3.GetPointAtDist(Arr_len);


                                            AngleStart = Functions.GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PtArr3.X, PtArr3.Y);
                                            AngleEnd = Functions.GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PointB3.X, PointB3.Y);
                                            Arc Arc31 = new Arc(Circle1.Center, Circle1.Radius, AngleEnd, AngleStart);
                                            double b01 = Math.Tan(Arc31.TotalAngle / 4);


                                            Poly123.AddVertexAt(0, new Point2d(PointB3.X, PointB3.Y), b01, 0, Arr_width);

                                            AngleStart = Functions.GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PtI2.X, PtI2.Y);
                                            AngleEnd = Functions.GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PtArr3.X, PtArr3.Y);
                                            Arc Arc32 = new Arc(Circle1.Center, Circle1.Radius, AngleEnd, AngleStart);
                                            double b02 = Math.Tan(Arc32.TotalAngle / 4);


                                            Poly123.AddVertexAt(1, new Point2d(PtArr3.X, PtArr3.Y), b02, 0, 0);
                                            Poly123.AddVertexAt(2, new Point2d(PtI2.X, PtI2.Y), b1, 0, 0);
                                            Poly123.AddVertexAt(3, new Point2d(PointT.X, PointT.Y), b2, 0, 0);

                                            if (Arc4.Length > Arr_len)
                                            {

                                                Point3d PtArr4 = Arc4.GetPointAtDist(Arc4.Length - Arr_len);

                                                AngleStart = Functions.GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PtArr4.X, PtArr4.Y);
                                                AngleEnd = Functions.GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PtI1.X, PtI1.Y);
                                                Arc Arc41 = new Arc(Circle1.Center, Circle1.Radius, AngleEnd, AngleStart);
                                                double b41 = Math.Tan(Arc41.TotalAngle / 4);
                                                Poly123.AddVertexAt(4, new Point2d(PtI1.X, PtI1.Y), b41, 0, 0);

                                                AngleStart = Functions.GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PointB4.X, PointB4.Y);
                                                AngleEnd = Functions.GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PtArr4.X, PtArr4.Y);
                                                Arc Arc42 = new Arc(Circle1.Center, Circle1.Radius, AngleEnd, AngleStart);
                                                double b42 = Math.Tan(Arc42.TotalAngle / 4);

                                                Poly123.AddVertexAt(5, new Point2d(PtArr4.X, PtArr4.Y), b42, Arr_width, 0);
                                                Poly123.AddVertexAt(6, new Point2d(PointB4.X, PointB4.Y), 0, 0, 0);

                                                BTrecord.AppendEntity(Poly123);
                                                Trans1.AddNewlyCreatedDBObject(Poly123, true);
                                                Trans1.TransactionManager.QueueForGraphicsFlush();


                                                Double Dist = Math.Pow((Point1.Value.X - Point2.Value.X) * (Point1.Value.X - Point2.Value.X) + (Point1.Value.Y - Point2.Value.Y) * (Point1.Value.Y - Point2.Value.Y), 0.5);

                                                if (comboBox_Scale.Text == "PSpace")
                                                {

                                                    Viewport Vp1 = Trans1.GetObject(Ent_vp_id, OpenMode.ForRead) as Viewport;
                                                    if (Vp1 != null)
                                                    {
                                                        Matrix3d TransforMatrix = Functions.PaperToModel(Vp1);

                                                        Point3d PointT1 = Point1.Value.TransformBy(TransforMatrix);
                                                        Point3d PointT2 = Point2.Value.TransformBy(TransforMatrix);
                                                        Dist = Math.Pow((PointT1.X - PointT2.X) * (PointT1.X - PointT2.X) + (PointT1.Y - PointT2.Y) * (PointT1.Y - PointT2.Y), 0.5);
                                                    }

                                                }

                                                String Distance_string = Functions.Get_String_Rounded(Dist, Rounding()) + "'" + suffix1();

                                                if (checkBox_thousand_sep.Checked == true)
                                                {
                                                    Distance_string = Functions.Get_String_Rounded_with_thousand_sep(Dist, Rounding()) + "'";
                                                }

                                                PromptPointResult PointMtext;

                                                jig_Mtext_at_acholade_class Jig_mt = new jig_Mtext_at_acholade_class(new MText(), TextH, Text_rotation, Distance_string);
                                                PointMtext = Jig_mt.BeginJig();


                                                if (PointMtext == null)
                                                {
                                                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                                                    Freeze_operations = false;
                                                    return;
                                                }

                                                if (PointMtext.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel)
                                                {
                                                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                                                    Freeze_operations = false;
                                                    return;
                                                }

                                                MText Mtext1 = new MText();
                                                Mtext1.Location = PointMtext.Value;
                                                Mtext1.TextHeight = TextH;
                                                Mtext1.Contents = Distance_string;
                                                Mtext1.Attachment = AttachmentPoint.MiddleCenter;
                                                if (background_fill() == true)
                                                {
                                                    Mtext1.BackgroundFill = true;
                                                    Mtext1.UseBackgroundColor = true;
                                                    Mtext1.BackgroundScaleFactor = 1.2;
                                                }
                                                Mtext1.Rotation = Text_rotation;
                                                BTrecord.AppendEntity(Mtext1);
                                                Trans1.AddNewlyCreatedDBObject(Mtext1, true);


                                                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                                                Trans1.Commit();
                                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                Freeze_operations = false;
                                                return;



                                            }
                                        }
                                    }
                                }





                            }


                            if (radioButton_add_0_ang_dim.Checked == true)
                            {


                                Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                                Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                                Prompt_rez.MessageForAdding = "\nSelect the angular dimensions:";
                                Prompt_rez.SingleOnly = false;
                                Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                                if (Rezultat1.Status != PromptStatus.OK)
                                {

                                    Trans1.Commit();
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    Freeze_operations = false;
                                    return;
                                }

                                for (int i = 0; i < Rezultat1.Value.Count; ++i)
                                {
                                    LineAngularDimension2 dimang = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForWrite) as LineAngularDimension2;
                                    if (dimang != null)
                                    {

                                        Double val1 = dimang.Measurement * 180 / Math.PI;
                                        String DMS = Functions.Get_0DMS(val1);

                                        dimang.DimensionText = DMS;




                                    }

                                }


                                Trans1.Commit();
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                Freeze_operations = false;
                                return;








                            }


                            Trans1.Commit();

                            if (radioButton_LT.Checked == true & textBox_LT_Index.Text != "")
                            {
                                textBox_LT_Index.Text = Increase_index(textBox_LT_Index.Text, 1);
                            }

                            if (radioButton_CT.Checked == true & textBox_CT_Index.Text != "")
                            {
                                textBox_CT_Index.Text = Increase_index(textBox_CT_Index.Text, 1);
                            }

                        }
                        goto l123;

                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                Freeze_operations = false;
            }

        }

        Point3d Convert_coordinate_to_new_CS(Point3d Point1)
        {
            Autodesk.Gis.Map.Platform.AcMapMap Acmap = Autodesk.Gis.Map.Platform.AcMapMap.GetCurrentMap();
            Point3d Point2 = new Point3d();
            string Curent_system = Acmap.GetMapSRS();
            if (string.IsNullOrEmpty(Curent_system) == true)
            {
                MessageBox.Show("Please set your coordinate system");
                return Point2;
            }

            OSGeo.MapGuide.MgCoordinateSystemFactory Coord_factory1 = new OSGeo.MapGuide.MgCoordinateSystemFactory();
            OSGeo.MapGuide.MgCoordinateSystemCatalog Catalog1 = Coord_factory1.GetCatalog();
            OSGeo.MapGuide.MgCoordinateSystemDictionary Dictionary1 = Catalog1.GetCoordinateSystemDictionary();
            OSGeo.MapGuide.MgCoordinateSystemEnum Enum1 = Dictionary1.GetEnum();

            OSGeo.MapGuide.MgCoordinateSystem CoordSys1 = Coord_factory1.Create(Curent_system);

            OSGeo.MapGuide.MgCoordinateSystem CoordSys2 = Dictionary1.GetCoordinateSystem(comboBox_coord_system.Text);

            OSGeo.MapGuide.MgCoordinateSystemTransform Transform1 = Coord_factory1.GetTransform(CoordSys1, CoordSys2);
            OSGeo.MapGuide.MgCoordinate Coord1 = Transform1.Transform(Point1.X, Point1.Y);

            Point2 = new Point3d(Coord1.X, Coord1.Y, 0);
            return Point2;
        }


        Point3d Convert_coordinate_to_ll84(double n, double e, string coord_sys_code)
        {
            OSGeo.MapGuide.MgCoordinateSystemFactory Coord_factory1 = new OSGeo.MapGuide.MgCoordinateSystemFactory();
            OSGeo.MapGuide.MgCoordinateSystemCatalog Catalog1 = Coord_factory1.GetCatalog();
            OSGeo.MapGuide.MgCoordinateSystemDictionary Dictionary1 = Catalog1.GetCoordinateSystemDictionary();
            OSGeo.MapGuide.MgCoordinateSystemEnum Enum1 = Dictionary1.GetEnum();

            int cscount1 = Dictionary1.GetSize();
            OSGeo.MapGuide.MgStringCollection col_names = Enum1.NextName(cscount1);

            string good_cs = "LL84";

            for (int i = 0; i < cscount1; ++i)
            {
                string nume1 = col_names.GetItem(i);
                if(nume1.ToLower()==coord_sys_code.ToLower())
                {
                    good_cs = nume1;
                }
            }

            OSGeo.MapGuide.MgCoordinateSystem CoordSys1 = Dictionary1.GetCoordinateSystem(good_cs);
            OSGeo.MapGuide.MgCoordinateSystem CoordSys2 = Dictionary1.GetCoordinateSystem("LL84");
            OSGeo.MapGuide.MgCoordinateSystemTransform Transform1 = Coord_factory1.GetTransform(CoordSys1, CoordSys2);
            OSGeo.MapGuide.MgCoordinate Coord1 = Transform1.Transform(e, n);
            return new Point3d (Coord1.X, Coord1.Y,0);
        }


        //int csCount = csDict.GetSize();

        //ed.WriteMessage("\nCoordinate System Count : " + csCount.ToString());

        //ed.WriteMessage("\n-------------------------------------------------");

        //MgStringCollection csNames = csDictEnum.NextName(csCount);

        //string csName = null;            

        //MgCoordinateSystem cs = null; 



        //for (int i = 0; i < csCount; i++)

        //{

        // csName = csNames.GetItem(i);

        // cs = csDict.GetCoordinateSystem(csName);                         

        //  ed.WriteMessage("\nCoordinate System Name : " + csName.ToString() + "  " + "CS Code :  " + cs.CsCode.ToString());

        // ed.WriteMessage("\n-------------------------------------------------");

        // }



        //Dim String_UTM83_12 As String = "PROJCS[" & Chr(34) & "UTM83-12" & Chr(34) & ",GEOGCS[" & Chr(34) & "LL83" & Chr(34) & ",DATUM[" & Chr(34) & "NAD83" & Chr(34) & ",SPHEROID[" & Chr(34) & "GRS1980" & Chr(34) & ",6378137.000,298.25722210]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.017453292519943295]],PROJECTION[" & Chr(34) & "Transverse_Mercator" & Chr(34) & "],PARAMETER[" & Chr(34) & "false_easting" & Chr(34) & ",500000.000],PARAMETER[" & Chr(34) & "false_northing" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "central_meridian" & Chr(34) & ",-111.00000000000000],PARAMETER[" & Chr(34) & "scale_factor" & Chr(34) & ",0.9996],PARAMETER[" & Chr(34) & "latitude_of_origin" & Chr(34) & ",0.000],UNIT[" & Chr(34) & "Meter" & Chr(34) & ",1.00000000000000]]"
        //Dim String_UTM83_11 As String = "PROJCS[" & Chr(34) & "UTM83-11" & Chr(34) & ",GEOGCS[" & Chr(34) & "LL83" & Chr(34) & ",DATUM[" & Chr(34) & "NAD83" & Chr(34) & ",SPHEROID[" & Chr(34) & "GRS1980" & Chr(34) & ",6378137.000,298.25722210]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.017453292519943295]],PROJECTION[" & Chr(34) & "Transverse_Mercator" & Chr(34) & "],PARAMETER[" & Chr(34) & "false_easting" & Chr(34) & ",500000.000],PARAMETER[" & Chr(34) & "false_northing" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "central_meridian" & Chr(34) & ",-117.00000000000000],PARAMETER[" & Chr(34) & "scale_factor" & Chr(34) & ",0.9996],PARAMETER[" & Chr(34) & "latitude_of_origin" & Chr(34) & ",0.000],UNIT[" & Chr(34) & "Meter" & Chr(34) & ",1.00000000000000]]"
        //Dim String_CANA83_10TM115 As String = "PROJCS[" & Chr(34) & "CANA83-10TM115" & Chr(34) & ",GEOGCS[" & Chr(34) & "LL83" & Chr(34) & ",DATUM[" & Chr(34) & "NAD83" & Chr(34) & ",SPHEROID[" & Chr(34) & "GRS1980" & Chr(34) & ",6378137.000,298.25722210]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.017453292519943295]],PROJECTION[" & Chr(34) & "Transverse_Mercator" & Chr(34) & "],PARAMETER[" & Chr(34) & "false_easting" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "false_northing" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "scale_factor" & Chr(34) & ",0.999200000000],PARAMETER[" & Chr(34) & "central_meridian" & Chr(34) & ",-115.00000000000000],PARAMETER[" & Chr(34) & "latitude_of_origin" & Chr(34) & ",0.00000000000000],UNIT[" & Chr(34) & "Meter" & Chr(34) & ",1.00000000000000]]"
        //Dim String_CANA83_3TM114 As String = "PROJCS[" & Chr(34) & "CANA83-3TM114" & Chr(34) & ",GEOGCS[" & Chr(34) & "LL83" & Chr(34) & ",DATUM[" & Chr(34) & "NAD83" & Chr(34) & ",SPHEROID[" & Chr(34) & "GRS1980" & Chr(34) & ",6378137.000,298.25722210]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.017453292519943295]],PROJECTION[" & Chr(34) & "Transverse_Mercator" & Chr(34) & "],PARAMETER[" & Chr(34) & "false_easting" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "false_northing" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "scale_factor" & Chr(34) & ",0.999900000000],PARAMETER[" & Chr(34) & "central_meridian" & Chr(34) & ",-114.00000000000000],PARAMETER[" & Chr(34) & "latitude_of_origin" & Chr(34) & ",0.00000000000000],UNIT[" & Chr(34) & "Meter" & Chr(34) & ",1.00000000000000]]"
        //Dim String_CANA83_3TM111 As String = "PROJCS[" & Chr(34) & "CANA83-3TM111" & Chr(34) & ",GEOGCS[" & Chr(34) & "LL83" & Chr(34) & ",DATUM[" & Chr(34) & "NAD83" & Chr(34) & ",SPHEROID[" & Chr(34) & "GRS1980" & Chr(34) & ",6378137.000,298.25722210]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.017453292519943295]],PROJECTION[" & Chr(34) & "Transverse_Mercator" & Chr(34) & "],PARAMETER[" & Chr(34) & "false_easting" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "false_northing" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "scale_factor" & Chr(34) & ",0.999900000000],PARAMETER[" & Chr(34) & "central_meridian" & Chr(34) & ",-111.00000000000000],PARAMETER[" & Chr(34) & "latitude_of_origin" & Chr(34) & ",0.00000000000000],UNIT[" & Chr(34) & "Meter" & Chr(34) & ",1.00000000000000]]"
        //Dim String_CANA83_3TM117 As String = "PROJCS[" & Chr(34) & "CANA83-3TM117" & Chr(34) & ",GEOGCS[" & Chr(34) & "LL83" & Chr(34) & ",DATUM[" & Chr(34) & "NAD83" & Chr(34) & ",SPHEROID[" & Chr(34) & "GRS1980" & Chr(34) & ",6378137.000,298.25722210]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.017453292519943295]],PROJECTION[" & Chr(34) & "Transverse_Mercator" & Chr(34) & "],PARAMETER[" & Chr(34) & "false_easting" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "false_northing" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "scale_factor" & Chr(34) & ",0.999900000000],PARAMETER[" & Chr(34) & "central_meridian" & Chr(34) & ",-117.00000000000000],PARAMETER[" & Chr(34) & "latitude_of_origin" & Chr(34) & ",0.00000000000000],UNIT[" & Chr(34) & "Meter" & Chr(34) & ",1.00000000000000]]"
        //Dim String_CANA83_3TM120 As String = "PROJCS[" & Chr(34) & "CANA83-3TM120" & Chr(34) & ",GEOGCS[" & Chr(34) & "LL83" & Chr(34) & ",DATUM[" & Chr(34) & "NAD83" & Chr(34) & ",SPHEROID[" & Chr(34) & "GRS1980" & Chr(34) & ",6378137.000,298.25722210]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.017453292519943295]],PROJECTION[" & Chr(34) & "Transverse_Mercator" & Chr(34) & "],PARAMETER[" & Chr(34) & "false_easting" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "false_northing" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "scale_factor" & Chr(34) & ",0.999900000000],PARAMETER[" & Chr(34) & "central_meridian" & Chr(34) & ",-120.00000000000000],PARAMETER[" & Chr(34) & "latitude_of_origin" & Chr(34) & ",0.00000000000000],UNIT[" & Chr(34) & "Meter" & Chr(34) & ",1.00000000000000]]"
        //Dim String_UTM27_11 As String = "PROJCS[" & Chr(34) & "UTM27-11" & Chr(34) & ",GEOGCS[" & Chr(34) & "LL27" & Chr(34) & ",DATUM[" & Chr(34) & "NAD27" & Chr(34) & ",SPHEROID[" & Chr(34) & "CLRK66" & Chr(34) & ",6378206.400,294.97869821]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.017453292519943295]],PROJECTION[" & Chr(34) & "Transverse_Mercator" & Chr(34) & "],PARAMETER[" & Chr(34) & "false_easting" & Chr(34) & ",500000.000],PARAMETER[" & Chr(34) & "false_northing" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "central_meridian" & Chr(34) & ",-117.00000000000000],PARAMETER[" & Chr(34) & "scale_factor" & Chr(34) & ",0.9996],PARAMETER[" & Chr(34) & "latitude_of_origin" & Chr(34) & ",0.000],UNIT[" & Chr(34) & "Meter" & Chr(34) & ",1.00000000000000]]"
        //Dim String_UTM27_12 As String = "PROJCS[" & Chr(34) & "UTM27-12" & Chr(34) & ",GEOGCS[" & Chr(34) & "LL27" & Chr(34) & ",DATUM[" & Chr(34) & "NAD27" & Chr(34) & ",SPHEROID[" & Chr(34) & "CLRK66" & Chr(34) & ",6378206.400,294.97869821]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.017453292519943295]],PROJECTION[" & Chr(34) & "Transverse_Mercator" & Chr(34) & "],PARAMETER[" & Chr(34) & "false_easting" & Chr(34) & ",500000.000],PARAMETER[" & Chr(34) & "false_northing" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "central_meridian" & Chr(34) & ",-111.00000000000000],PARAMETER[" & Chr(34) & "scale_factor" & Chr(34) & ",0.9996],PARAMETER[" & Chr(34) & "latitude_of_origin" & Chr(34) & ",0.000],UNIT[" & Chr(34) & "Meter" & Chr(34) & ",1.00000000000000]]"
        //Dim String_LL84 As String = "GEOGCS[" & Chr(34) & "LL84" & Chr(34) & ",DATUM[" & Chr(34) & "WGS84" & Chr(34) & ",SPHEROID[" & Chr(34) & "WGS84" & Chr(34) & ",6378137.000,298.25722293]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.01745329251994]]"
        //Dim String_LL83 As String = "GEOGCS[" & Chr(34) & "LL83" & Chr(34) & ",DATUM[" & Chr(34) & "NAD83" & Chr(34) & ",SPHEROID[" & Chr(34) & "GRS1980" & Chr(34) & ",6378137.000,298.25722210]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.01745329251994]]"
        //Dim String_UTM83_10 As String = "PROJCS[" & Chr(34) & "UTM83-10" & Chr(34) & ",GEOGCS[" & Chr(34) & "LL83" & Chr(34) & ",DATUM[" & Chr(34) & "NAD83" & Chr(34) & ",SPHEROID[" & Chr(34) & "GRS1980" & Chr(34) & ",6378137.000,298.25722210]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.017453292519943295]],PROJECTION[" & Chr(34) & "Transverse_Mercator" & Chr(34) & "],PARAMETER[" & Chr(34) & "false_easting" & Chr(34) & ",500000.000],PARAMETER[" & Chr(34) & "false_northing" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "central_meridian" & Chr(34) & ",-123.00000000000000],PARAMETER[" & Chr(34) & "scale_factor" & Chr(34) & ",0.9996],PARAMETER[" & Chr(34) & "latitude_of_origin" & Chr(34) & ",0.000],UNIT[" & Chr(34) & "Meter" & Chr(34) & ",1.00000000000000]]"
        //Dim String_UTM27_10 As String = "PROJCS[" & Chr(34) & "UTM27-10" & Chr(34) & ",GEOGCS[" & Chr(34) & "LL27" & Chr(34) & ",DATUM[" & Chr(34) & "NAD27" & Chr(34) & ",SPHEROID[" & Chr(34) & "CLRK66" & Chr(34) & ",6378206.400,294.97869821]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.017453292519943295]],PROJECTION[" & Chr(34) & "Transverse_Mercator" & Chr(34) & "],PARAMETER[" & Chr(34) & "false_easting" & Chr(34) & ",500000.000],PARAMETER[" & Chr(34) & "false_northing" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "central_meridian" & Chr(34) & ",-123.00000000000000],PARAMETER[" & Chr(34) & "scale_factor" & Chr(34) & ",0.9996],PARAMETER[" & Chr(34) & "latitude_of_origin" & Chr(34) & ",0.000],UNIT[" & Chr(34) & "Meter" & Chr(34) & ",1.00000000000000]]"


        // Dim Coord_factory1 As New OSGeo.MapGuide.MgCoordinateSystemFactory
        //Dim CoordSys1 As OSGeo.MapGuide.MgCoordinateSystem '= Coord_factory1.Create(Curent_system)
        //Dim CoordSys2 As OSGeo.MapGuide.MgCoordinateSystem ' = Coord_factory1.Create(String_LL84)

        //Coord1 = Transform1.Transform(Val(W1.Cells(i, Column_x_from).value), Val(W1.Cells(i, Column_y_from).value))

        // Dim Transform1 As OSGeo.MapGuide.MgCoordinateSystemTransform = Coord_factory1.GetTransform(CoordSys1, CoordSys2)
        //Dim Coord1 As OSGeo.MapGuide.MgCoordinate '= Transform1.Transform(X5, Y5)

        public ObjectId Get_Arrow_dimension_ID(string NUME_variabila, string NUME_ARROW)
        {

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            string OLD_VALUE = (string)Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable(NUME_variabila);
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable(NUME_variabila, NUME_ARROW);
            if (OLD_VALUE.Length != 0) Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable(NUME_variabila, OLD_VALUE);
            ObjectId ID1;
            using (Transaction Trans1 = ThisDrawing.Database.TransactionManager.StartTransaction())
            {
                BlockTable BlockTable1 = (BlockTable)Trans1.GetObject(ThisDrawing.Database.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                ID1 = BlockTable1[NUME_ARROW];
                Trans1.Commit();


            }

            return ID1;

        }

        private void button_label_on_poly_Click(object sender, EventArgs e)
        {
            bool space_in_bear = true;
            if (checkBox_no_space_in_bearing.Checked == true) space_in_bear = false;

            if (Freeze_operations == false)
            {
                Freeze_operations = true;

                try
                {
                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                    Matrix3d CurentUCSmatrix = Editor1.CurrentUserCoordinateSystem;

                    string nume_layer_pt = "__PT";

                    System.Data.DataTable TableL = new System.Data.DataTable();

                    TableL.Columns.Add("PT_START", typeof(string));
                    TableL.Columns.Add("PT_END", typeof(string));
                    TableL.Columns.Add("BEARING", typeof(string));
                    TableL.Columns.Add("DISTANCE", typeof(string));

                    Functions.Creaza_layer(nume_layer_pt, 5, false);


                    using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                    {
                    l123:


                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                            BlockTable BlockTable1 = (BlockTable)Trans1.GetObject(ThisDrawing.Database.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                            Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

                            Create_point_object_data();

                            string NumeBlock = "";


                            if (BlockTable1.Has(comboBox_blocks.Text) == true)
                            {
                                NumeBlock = comboBox_blocks.Text;
                            }
                            else
                            {
                                radioButton_BD.Checked = true;
                                MessageBox.Show("Line_Table block is not inserted on this drawing");
                            }




                            Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rez.MessageForAdding = "\nSelect the polylines:";
                            Prompt_rez.SingleOnly = false;
                            Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                            if (Rezultat1.Status != PromptStatus.OK)
                            {
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                Freeze_operations = false;
                                return;
                            }

                            string current_number = textBox_current_number.Text;

                            for (int i = 0; i < Rezultat1.Value.Count; ++i)
                            {
                                Entity Ent1 = (Entity)Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead);
                                if (Ent1 is Polyline)
                                {
                                    Polyline Poly1 = (Polyline)Ent1;
                                    if (Poly1.NumberOfVertices > 1)
                                    {
                                        for (int k = 1; k < Poly1.NumberOfVertices; ++k)
                                        {
                                            Point3d Pt1 = Poly1.GetPoint3dAt(k - 1);
                                            Point3d Pt2 = Poly1.GetPoint3dAt(k);


                                            double x1 = Pt1.X;
                                            double y1 = Pt1.Y;
                                            double x2 = Pt2.X;
                                            double y2 = Pt2.Y;
                                            double dist1 = Math.Pow(Math.Pow(x1 - x2, 2) + Math.Pow(y1 - y2, 2), 0.5);
                                            double bearing1 = Functions.GET_Bearing_rad(x1, y1, x2, y2);
                                            double rot_t = Functions.calculate_rotatie_text(bearing1);
                                            string quadrant1 = Functions.Get_Quadrant_bearing(bearing1, space_in_bear);
                                            string next_number = Get_next_number(current_number, 1, 2);


                                            DBPoint Point1 = new DBPoint(Pt1);
                                            Point1.Layer = nume_layer_pt;
                                            BTrecord.AppendEntity(Point1);
                                            Trans1.AddNewlyCreatedDBObject(Point1, true);

                                            Point1.UpgradeOpen();
                                            List<object> List1 = new List<object>();
                                            List<Autodesk.Gis.Map.Constants.DataType> List2 = new List<Autodesk.Gis.Map.Constants.DataType>();
                                            List1.Add(current_number);
                                            List2.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                                            Functions.Populate_object_data_table(Tables1, Point1.ObjectId, "Point_numbers", List1, List2);

                                            if (k == Poly1.NumberOfVertices - 1)
                                            {
                                                DBPoint Point2 = new DBPoint(Pt2);
                                                Point2.Layer = nume_layer_pt;
                                                BTrecord.AppendEntity(Point2);
                                                Trans1.AddNewlyCreatedDBObject(Point2, true);

                                                Point2.UpgradeOpen();
                                                List1 = new List<object>();
                                                List2 = new List<Autodesk.Gis.Map.Constants.DataType>();
                                                List1.Add(next_number);
                                                List2.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                                                Functions.Populate_object_data_table(Tables1, Point2.ObjectId, "Point_numbers", List1, List2);
                                            }

                                            TableL.Rows.Add();
                                            TableL.Rows[TableL.Rows.Count - 1]["PT_START"] = current_number;
                                            TableL.Rows[TableL.Rows.Count - 1]["PT_END"] = next_number;
                                            TableL.Rows[TableL.Rows.Count - 1]["BEARING"] = quadrant1;
                                            TableL.Rows[TableL.Rows.Count - 1]["DISTANCE"] = Functions.Get_String_Rounded(dist1, Rounding()) + "'";
                                            current_number = next_number;
                                        }
                                    }

                                }


                                if (Ent1 is Line)
                                {
                                    Line line1 = (Line)Ent1;

                                    Point3d Pt1 = line1.StartPoint;
                                    Point3d Pt2 = line1.EndPoint;


                                    double x1 = Pt1.X;
                                    double y1 = Pt1.Y;
                                    double x2 = Pt2.X;
                                    double y2 = Pt2.Y;
                                    double dist1 = Math.Pow(Math.Pow(x1 - x2, 2) + Math.Pow(y1 - y2, 2), 0.5);
                                    double bearing1 = Functions.GET_Bearing_rad(x1, y1, x2, y2);
                                    double rot_t = Functions.calculate_rotatie_text(bearing1);
                                    string quadrant1 = Functions.Get_Quadrant_bearing(bearing1, space_in_bear);
                                    string next_number = Get_next_number(current_number, 1, 2);


                                    DBPoint Point1 = new DBPoint(Pt1);
                                    Point1.Layer = nume_layer_pt;
                                    BTrecord.AppendEntity(Point1);
                                    Trans1.AddNewlyCreatedDBObject(Point1, true);

                                    Point1.UpgradeOpen();
                                    List<object> List1 = new List<object>();
                                    List<Autodesk.Gis.Map.Constants.DataType> List2 = new List<Autodesk.Gis.Map.Constants.DataType>();
                                    List1.Add(current_number);
                                    List2.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                                    Functions.Populate_object_data_table(Tables1, Point1.ObjectId, "Point_numbers", List1, List2);

                                    DBPoint Point2 = new DBPoint(Pt2);
                                    Point2.Layer = nume_layer_pt;
                                    BTrecord.AppendEntity(Point2);
                                    Trans1.AddNewlyCreatedDBObject(Point2, true);

                                    Point2.UpgradeOpen();
                                    List1 = new List<object>();
                                    List2 = new List<Autodesk.Gis.Map.Constants.DataType>();
                                    List1.Add(next_number);
                                    List2.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                                    Functions.Populate_object_data_table(Tables1, Point2.ObjectId, "Point_numbers", List1, List2);


                                    TableL.Rows.Add();
                                    TableL.Rows[TableL.Rows.Count - 1]["PT_START"] = current_number;
                                    TableL.Rows[TableL.Rows.Count - 1]["PT_END"] = next_number;
                                    TableL.Rows[TableL.Rows.Count - 1]["BEARING"] = quadrant1;
                                    TableL.Rows[TableL.Rows.Count - 1]["DISTANCE"] = Functions.Get_String_Rounded(dist1, Rounding()) + "'";
                                    current_number = next_number;


                                }

                                current_number = Get_next_number(current_number, 1, 2);
                            }




                            Insert_blocks(TableL, NumeBlock);
                            Trans1.Commit();
                            textBox_current_number.Text = current_number;
                        }

                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                Freeze_operations = false;
            }
        }

        private string Get_next_number(string current_number, int pref_len, int suff_len)
        {
            string suffix_string = current_number.Substring(pref_len, suff_len);
            string prefix = current_number.Substring(0, pref_len);
            int suffix_number = 0;
            if (Functions.IsNumeric(suffix_string) == true)
            {
                suffix_number = Convert.ToInt32(suffix_string);
            }
            int next_suffix_number = suffix_number + 1;
            string next_suffix_string = next_suffix_number.ToString();
            if (next_suffix_number < 10)
            {
                next_suffix_string = "0" + next_suffix_string;
            }
            int asc = (int)Convert.ToChar(prefix.Substring(0, 1));
            if (next_suffix_number == 100)
            {
                prefix = Convert.ToString(Convert.ToChar(asc + 1));
                next_suffix_string = "01";
            }
            string next_number = prefix + next_suffix_string;
            return next_number;
        }


        private void panel3_Click(object sender, EventArgs e)
        {
            if (Freeze_operations == false)
            {
                Freeze_operations = true;
                try
                {
                    Functions.Incarca_existing_Blocks_with_attributes_to_combobox(comboBox_blocks);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                Freeze_operations = false;
            }
        }

        private void Create_point_object_data()
        {

            {
                try
                {
                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                    using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {

                            List<string> List1 = new List<string>();
                            List<string> List2 = new List<string>();
                            List<Autodesk.Gis.Map.Constants.DataType> List3 = new List<Autodesk.Gis.Map.Constants.DataType>();

                            List1.Add("POINT_ID");
                            List2.Add("Point name");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);



                            Functions.Get_object_data_table("Point_numbers", "Generated by BRD", List1, List2, List3);


                            Trans1.Commit();
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

        }



    }
}
