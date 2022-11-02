using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System.Management;
using System.Windows.Forms;
using Autodesk.AutoCAD.EditorInput;
using System.Runtime.InteropServices;

namespace NESW
{
    public class Commands
    {

        public static bool isSECURE()
        {

            try
            {
                string UserDNS = Environment.GetEnvironmentVariable("USERDNSDOMAIN");
                if (UserDNS.ToLower() == "mottmac.group.int")
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (System.Exception ex)
            {
                return false;
            }

        }

        [CommandMethod("SW")]
        public void creaza_label_n_e_architectural()
        {

            if (isSECURE() == false) return;

            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();


            try
            {

                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        BlockTableRecord BTrecord_MS = Trans1.GetObject(BlockTable1[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                        BlockTableRecord BTrecord_PS = Trans1.GetObject(BlockTable1[BlockTableRecord.PaperSpace], OpenMode.ForRead) as BlockTableRecord;

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rez1;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_vp;
                        Prompt_vp = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect viewport:");
                        Prompt_vp.SetRejectMessage("\nSelect a viewport!");
                        Prompt_vp.AllowNone = true;
                        Prompt_vp.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Viewport), false);
                        Rez1 = ThisDrawing.Editor.GetEntity(Prompt_vp);

                        if (Rez1.Status != PromptStatus.OK)
                        {
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Creaza_layer("Text", 2, true);
                        Creaza_layer("No Plot", 40, false);

                        bool run = true;
                        do
                        {
                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                            Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                            PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify point");
                            PP1.AllowNone = false;
                            Point_res1 = Editor1.GetPoint(PP1);

                            if (Point_res1.Status != PromptStatus.OK)
                            {
                                run = false;
                            }

                            if (run == true)
                            {
                                double x1 = Point_res1.Value.X;
                                double y1 = Point_res1.Value.Y;
                                double z1 = Point_res1.Value.Z;
                                Point3d pt_ps = new Point3d(x1, y1, 0);
                                Point3d pt_ps1 = new Point3d(x1 + 0.05, y1 + 0.05, 0);
                                Point3d pt_ps2 = new Point3d(x1 - 0.05, y1 - 0.05, 0);



                                Viewport vp1 = Trans1.GetObject(Rez1.ObjectId, OpenMode.ForRead) as Viewport;

                                if (vp1 != null)
                                {
                                    Matrix3d matrix1 = PaperToModel(vp1);
                                    Point3d pt_ms = pt_ps.TransformBy(matrix1);

                                    DBPoint tst1 = new DBPoint(pt_ms);
                                    tst1.Layer = "No Plot";
                                    BTrecord_MS.AppendEntity(tst1);
                                    Trans1.AddNewlyCreatedDBObject(tst1, true);

                                    Matrix3d matrix2 = get_ucs_from_vp(vp1);



                                    double n_inch = pt_ms.TransformBy(matrix2).Y;
                                    string MINUS = "N: ";
                                    if (n_inch < 0)
                                    {
                                        n_inch = -1 * n_inch;
                                        MINUS = "S: ";
                                    }

                                    double n_foot = n_inch / 12;

                                    double f_n = Math.Floor(n_foot);


                                    double i_n = (n_foot - f_n) * 12;

                                    double i_n1 = Math.Floor(i_n);

                                    double rest_n = i_n - i_n1;

                                    double rest_n1 = Round_Closest(rest_n * 100, 25);

                                    string dec_str1 = DecimalToFraction(rest_n1 / 100);

                                    string inches1 = "\"";

                                    if (dec_str1.Length > 1)
                                    {
                                        inches1 = " " + dec_str1 + "\"";
                                    }



                                    string sta1 = MINUS + Get_chainage_from_double(f_n, "f", 0) + "'-" + Convert.ToString(i_n1) + inches1;


                                    MText mtext1 = new MText();
                                    mtext1.Contents = sta1;
                                    mtext1.TextHeight = 0.1;
                                    mtext1.Rotation = 0;
                                    mtext1.Location = pt_ps1;
                                    mtext1.Layer = "Text";
                                    mtext1.Attachment = AttachmentPoint.BottomLeft;
                                    BTrecord.AppendEntity(mtext1);
                                    Trans1.AddNewlyCreatedDBObject(mtext1, true);

                                    Move_mtext_Jig_X jig_X = new Move_mtext_Jig_X(mtext1);
                                    PromptResult jig_prompt = Editor1.Drag(jig_X);

                                    // if not OK, clean return
                                    if (jig_prompt.Status != PromptStatus.OK)
                                    {
                                        run = false;
                                        mtext1.Erase();
                                    }
                                    else
                                    {
                                        // everything ok, update the location
                                        mtext1.UpgradeOpen();
                                        mtext1.Location = jig_X.Location;
                                    }



                                    double e_inch = pt_ms.TransformBy(matrix2).X;

                                    MINUS = "E: ";
                                    if (e_inch < 0)
                                    {
                                        e_inch = -1 * e_inch;
                                        MINUS = "W: ";
                                    }

                                    double e_foot = e_inch / 12;

                                    double f_e = Math.Floor(e_foot);

                                    double i_e = (e_foot - f_e) * 12;

                                    double i_e1 = Math.Floor(i_e);

                                    double rest_e = i_e - i_e1;

                                    double rest_e1 = Round_Closest(rest_e * 100, 25);

                                    string dec_str2 = DecimalToFraction(rest_e1 / 100);

                                    string inches2 = "\"";

                                    if (dec_str2.Length > 1)
                                    {
                                        inches2 = " " + dec_str2 + "\"";
                                    }

                                    string sta2 = MINUS + Get_chainage_from_double(f_e, "f", 0) + "'-" + Convert.ToString(i_e1) + inches2;


                                    MText mtext2 = new MText();
                                    mtext2.Contents = sta2;
                                    mtext2.TextHeight = 0.1;
                                    mtext2.Rotation = Math.PI / 2;
                                    mtext2.Location = pt_ps2;
                                    mtext2.Layer = "Text";
                                    mtext2.Attachment = AttachmentPoint.BottomRight;
                                    BTrecord.AppendEntity(mtext2);
                                    Trans1.AddNewlyCreatedDBObject(mtext2, true);


                                    Move_mtext_Jig_Y jig_Y = new Move_mtext_Jig_Y(mtext2);
                                    jig_prompt = Editor1.Drag(jig_Y);

                                    // if not OK, clean return
                                    if (jig_prompt.Status != PromptStatus.OK)
                                    {
                                        run = false;
                                        mtext2.Erase();
                                    }
                                    else
                                    {
                                        // everything ok, update the location
                                        mtext2.UpgradeOpen();
                                        mtext2.Location = jig_Y.Location;
                                    }





                                }
                                else
                                {
                                    run = false;
                                }
                            }

                            Trans1.TransactionManager.QueueForGraphicsFlush();

                        } while (run == true);





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


        }

        [CommandMethod("SWNE")]
        public void creaza_label_n_e_architectural1()
        {

            if (isSECURE() == false) return;

            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();


            try
            {

                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    bool run = true;
                    Autodesk.AutoCAD.EditorInput.PromptEntityResult Rez1;
                    Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_vp;
                    Prompt_vp = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect viewport:");
                    Prompt_vp.SetRejectMessage("\nSelect a viewport!");
                    Prompt_vp.AllowNone = true;
                    Prompt_vp.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Viewport), false);
                    Rez1 = ThisDrawing.Editor.GetEntity(Prompt_vp);

                    if (Rez1.Status != PromptStatus.OK)
                    {
                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                        return;
                    }



                    do
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                            BlockTableRecord BTrecord_MS = Trans1.GetObject(BlockTable1[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                            BlockTableRecord BTrecord_PS = Trans1.GetObject(BlockTable1[BlockTableRecord.PaperSpace], OpenMode.ForRead) as BlockTableRecord;



                            Creaza_layer("Text", 2, true);
                            Creaza_layer("No Plot", 40, false);


                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                            Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                            PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify point");
                            PP1.AllowNone = false;
                            Point_res1 = Editor1.GetPoint(PP1);

                            if (Point_res1.Status != PromptStatus.OK)
                            {
                                run = false;
                            }

                            if (run == true)
                            {
                                double x1 = Point_res1.Value.X;
                                double y1 = Point_res1.Value.Y;
                                double z1 = Point_res1.Value.Z;
                                Point3d pt_ps = new Point3d(x1, y1, 0);
                                Point3d pt_ps1 = new Point3d(x1 + 0.05, y1 + 0.05, 0);
                                Point3d pt_ps2 = new Point3d(x1 - 0.05, y1 - 0.05, 0);



                                Viewport vp1 = Trans1.GetObject(Rez1.ObjectId, OpenMode.ForRead) as Viewport;

                                if (vp1 != null)
                                {
                                    Matrix3d matrix1 = PaperToModel(vp1);
                                    Point3d pt_ms = pt_ps.TransformBy(matrix1);

                                    DBPoint tst1 = new DBPoint(pt_ms);
                                    tst1.Layer = "No Plot";
                                    BTrecord_MS.AppendEntity(tst1);
                                    Trans1.AddNewlyCreatedDBObject(tst1, true);

                                    Matrix3d matrix2 = get_ucs_from_vp(vp1);



                                    double n_inch = pt_ms.TransformBy(matrix2).Y;
                                    string MINUS = "N: ";
                                    if (n_inch < 0)
                                    {
                                        n_inch = -1 * n_inch;
                                        MINUS = "S: ";
                                    }

                                    double n_foot = n_inch / 12;

                                    double f_n = Math.Floor(n_foot);


                                    double i_n = (n_foot - f_n) * 12;

                                    double i_n1 = Math.Floor(i_n);

                                    double rest_n = i_n - i_n1;

                                    double rest_n1 = Round_Closest(rest_n * 100, 25);

                                    string dec_str1 = DecimalToFraction(rest_n1 / 100);

                                    string inches1 = "\"";

                                    if (dec_str1.Length > 1)
                                    {
                                        inches1 = " " + dec_str1 + "\"";
                                    }



                                    string sta1 = MINUS + Get_chainage_from_double(f_n, "f", 0) + "'-" + Convert.ToString(i_n1) + inches1;


                                    MText mtext1 = new MText();
                                    mtext1.Contents = sta1;
                                    mtext1.TextHeight = 0.1;
                                    mtext1.Rotation = 0;
                                    mtext1.Location = pt_ps1;
                                    mtext1.Layer = "Text";
                                    mtext1.Attachment = AttachmentPoint.BottomLeft;
                                    BTrecord.AppendEntity(mtext1);
                                    Trans1.AddNewlyCreatedDBObject(mtext1, true);

                                    Move_mtext_Jig_X jig_X = new Move_mtext_Jig_X(mtext1);
                                    PromptResult jig_prompt = Editor1.Drag(jig_X);

                                    // if not OK, clean return
                                    if (jig_prompt.Status != PromptStatus.OK)
                                    {
                                        run = false;
                                        mtext1.Erase();
                                    }
                                    else
                                    {
                                        // everything ok, update the location
                                        //mtext1.UpgradeOpen();
                                        mtext1.Location = jig_X.Location;
                                    }



                                    double e_inch = pt_ms.TransformBy(matrix2).X;

                                    MINUS = "E: ";
                                    if (e_inch < 0)
                                    {
                                        e_inch = -1 * e_inch;
                                        MINUS = "W: ";
                                    }

                                    double e_foot = e_inch / 12;

                                    double f_e = Math.Floor(e_foot);

                                    double i_e = (e_foot - f_e) * 12;

                                    double i_e1 = Math.Floor(i_e);

                                    double rest_e = i_e - i_e1;

                                    double rest_e1 = Round_Closest(rest_e * 100, 25);

                                    string dec_str2 = DecimalToFraction(rest_e1 / 100);

                                    string inches2 = "\"";

                                    if (dec_str2.Length > 1)
                                    {
                                        inches2 = " " + dec_str2 + "\"";
                                    }

                                    string sta2 = MINUS + Get_chainage_from_double(f_e, "f", 0) + "'-" + Convert.ToString(i_e1) + inches2;


                                    MText mtext2 = new MText();
                                    mtext2.Contents = sta2;
                                    mtext2.TextHeight = 0.1;
                                    mtext2.Rotation = Math.PI / 2;
                                    mtext2.Location = pt_ps2;
                                    mtext2.Layer = "Text";
                                    mtext2.Attachment = AttachmentPoint.BottomRight;
                                    BTrecord.AppendEntity(mtext2);
                                    Trans1.AddNewlyCreatedDBObject(mtext2, true);


                                    Move_mtext_Jig_Y jig_Y = new Move_mtext_Jig_Y(mtext2);
                                    jig_prompt = Editor1.Drag(jig_Y);

                                    // if not OK, clean return
                                    if (jig_prompt.Status != PromptStatus.OK)
                                    {
                                        run = false;
                                        mtext2.Erase();
                                    }
                                    else
                                    {
                                        // everything ok, update the location
                                        //mtext2.UpgradeOpen();
                                        mtext2.Location = jig_Y.Location;
                                    }





                                }
                                else
                                {
                                    run = false;
                                }
                            }

                            Trans1.TransactionManager.QueueForGraphicsFlush();
                            Trans1.Commit();
                        }

                    } while (run == true);
                }
            }
            catch (System.Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");


        }


        [CommandMethod("NE")]
        public void creaza_label_n_e_architectural3()
        {

            if (isSECURE() == false) return;

            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();


            try
            {

                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        BlockTableRecord BTrecord_MS = Trans1.GetObject(BlockTable1[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                        BlockTableRecord BTrecord_PS = Trans1.GetObject(BlockTable1[BlockTableRecord.PaperSpace], OpenMode.ForRead) as BlockTableRecord;

                        int precision1 = 25;
                        string inches = "\"";

                        PromptKeywordOptions pko = new PromptKeywordOptions("\n" + "Specify precision: ");
                        pko.Keywords.Add("0.25" + inches);
                        pko.Keywords.Add("0.5" + inches);
                        pko.Keywords.Default = "0.25" + inches;
                        pko.AllowNone = true;

                        PromptResult res = ThisDrawing.Editor.GetKeywords(pko);

                        if (res.StringResult.Replace(inches, "") == "0.5") precision1 = 50;

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rez1;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_vp;
                        Prompt_vp = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect viewport:");
                        Prompt_vp.SetRejectMessage("\nSelect a viewport!");
                        Prompt_vp.AllowNone = true;
                        Prompt_vp.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Viewport), false);
                        Rez1 = ThisDrawing.Editor.GetEntity(Prompt_vp);

                        if (Rez1.Status != PromptStatus.OK)
                        {
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Creaza_layer("Text", 2, true);
                        Creaza_layer("No Plot", 40, false);

                        bool run = true;
                        do
                        {
                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                            Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                            PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify point");
                            PP1.AllowNone = false;
                            Point_res1 = Editor1.GetPoint(PP1);

                            if (Point_res1.Status != PromptStatus.OK)
                            {
                                run = false;
                            }

                            if (run == true)
                            {
                                double x1 = Point_res1.Value.X;
                                double y1 = Point_res1.Value.Y;
                                double z1 = Point_res1.Value.Z;
                                Point3d pt_ps = new Point3d(x1, y1, 0);
                                Point3d pt_ps1 = new Point3d(x1 + 0.05, y1 + 0.05, 0);
                                Point3d pt_ps2 = new Point3d(x1 - 0.05, y1 - 0.05, 0);



                                Viewport vp1 = Trans1.GetObject(Rez1.ObjectId, OpenMode.ForRead) as Viewport;

                                if (vp1 != null)
                                {
                                    Matrix3d matrix1 = PaperToModel(vp1);
                                    Point3d pt_ms = pt_ps.TransformBy(matrix1);

                                    DBPoint tst1 = new DBPoint(pt_ms);
                                    tst1.Layer = "No Plot";
                                    BTrecord_MS.AppendEntity(tst1);
                                    Trans1.AddNewlyCreatedDBObject(tst1, true);

                                    Matrix3d matrix2 = get_ucs_from_vp(vp1);

                                    double n_inch = pt_ms.TransformBy(matrix2).Y;
                                    string MINUS = "N: ";
                                    if (n_inch < 0)
                                    {
                                        n_inch = -1 * n_inch;
                                        MINUS = "S: ";
                                    }

                                    double n_foot = n_inch / 12;

                                    double f_n = Math.Floor(n_foot);

                                    double i_n = (n_foot - f_n) * 12;

                                    double i_n1 = Math.Round(i_n, 0);

                                    double rest_n = i_n - i_n1;

                                    double rest_n1 = Round_Closest(rest_n * 100, precision1);

                                    string dec_str1 = DecimalToFraction(rest_n1 / 100);

                                    string inches1 = "\"";

                                    if (dec_str1.Length > 1)
                                    {
                                        inches1 = " " + dec_str1 + "\"";
                                    }

                                    string sta1 = MINUS + Get_chainage_from_double(f_n, "f", 0) + "'-" + Convert.ToString(i_n1) + inches1;

                                    Entity new_mtext1 = Mtextjig_X.Jig(sta1, 0, 0.1, AttachmentPoint.BottomLeft, pt_ps1.Y);

                                    if (new_mtext1 != null)
                                    {
                                        new_mtext1.Layer = "Text";
                                        BTrecord.AppendEntity(new_mtext1);
                                        Trans1.AddNewlyCreatedDBObject(new_mtext1, true);
                                        Trans1.TransactionManager.QueueForGraphicsFlush();
                                    }
                                    else
                                    {
                                        run = false;
                                    }

                                    double e_inch = pt_ms.TransformBy(matrix2).X;

                                    MINUS = "E: ";
                                    if (e_inch < 0)
                                    {
                                        e_inch = -1 * e_inch;
                                        MINUS = "W: ";
                                    }

                                    double e_foot = e_inch / 12;

                                    double f_e = Math.Floor(e_foot);

                                    double i_e = (e_foot - f_e) * 12;

                                    double i_e1 = Math.Round(i_e, 0);

                                    double rest_e = i_e - i_e1;

                                    double rest_e1 = Round_Closest(rest_e * 100, precision1);

                                    string dec_str2 = DecimalToFraction(rest_e1 / 100);

                                    string inches2 = "\"";

                                    if (dec_str2.Length > 1)
                                    {
                                        inches2 = " " + dec_str2 + "\"";
                                    }

                                    string sta2 = MINUS + Get_chainage_from_double(f_e, "f", 0) + "'-" + Convert.ToString(i_e1) + inches2;

                                    Entity new_mtext2 = Mtextjig_Y.Jig(sta2, Math.PI / 2, 0.1, AttachmentPoint.BottomRight, pt_ps2.X);
                                    if (new_mtext2 != null)
                                    {
                                        new_mtext2.Layer = "Text";
                                        BTrecord.AppendEntity(new_mtext2);
                                        Trans1.AddNewlyCreatedDBObject(new_mtext2, true);
                                        Trans1.TransactionManager.QueueForGraphicsFlush();
                                    }
                                    else
                                    {
                                        run = false;
                                    }
                                }
                                else
                                {
                                    run = false;
                                }
                            }

                            Trans1.TransactionManager.QueueForGraphicsFlush();

                        } while (run == true);

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


        }


        public int Round_Up(double numToRound, int multiple)
        {
            if (multiple == 0)
            {
                return Convert.ToInt32(numToRound);
            }

            int remainder = Convert.ToInt32(numToRound) % multiple;
            if (remainder == 0)
            {
                return Convert.ToInt32(numToRound);
            }
            return Convert.ToInt32(numToRound) + multiple - remainder;
        }

        public int Round_Down(double numToRound, int multiple)
        {
            if (multiple == 0)
            {
                return Convert.ToInt32(numToRound);
            }

            int remainder = Convert.ToInt32(numToRound) % multiple;
            if (remainder == 0)
            {
                return Convert.ToInt32(numToRound);
            }

            return Convert.ToInt32(numToRound) - remainder;
        }


        public int Round_Closest(double numToRound, int multiple)
        {
            int Numar = Convert.ToInt32(numToRound);
            int Up = Round_Up(numToRound, multiple);
            int Down = Round_Down(numToRound, multiple);
            if (Math.Abs(Numar - Up) < Math.Abs(Numar - Down))
            {
                return Up;
            }
            else
            {
                return Down;
            }
        }

        public string DecimalToFraction(double dec)
        {
            string str = dec.ToString();
            if (str.Contains('.'))
            {
                string[] parts = str.Split('.');
                long whole = long.Parse(parts[0]);
                long numerator = long.Parse(parts[1]);
                long denominator = (long)Math.Pow(10, parts[1].Length);
                long divisor = GCD(numerator, denominator);
                long num = numerator / divisor;
                long den = denominator / divisor;

                string fraction = num + "/" + den;
                if (whole > 0)
                {
                    return whole + " " + fraction;
                }
                else
                {
                    return fraction;
                }
            }
            else
            {
                return str;
            }
        }

        public long GCD(long a, long b)
        {
            return b == 0 ? a : GCD(b, a % b);
        }


        public string Get_String_Rounded(double Numar, int Nr_dec)
        {

            string String1, String2, Zero, zero1;
            Zero = "";
            zero1 = "";

            string String_punct = "";

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

        public string Get_chainage_from_double(double Numar, string units, int Nr_dec)
        {

            string String2, String3;
            String3 = "";
            string String_minus = "";

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


        public Matrix3d get_ucs_from_vp(Viewport vp)
        {

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Editor1.SwitchToModelSpace();
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("CVPORT", vp.Number);
            Matrix3d ucs = Editor1.CurrentUserCoordinateSystem;
            Editor1.SwitchToPaperSpace();
            return ucs.Inverse();
        }


        public Matrix3d ModelToPaper(Viewport vp)
        {
            Vector3d vd = vp.ViewDirection;
            Point3d vc = new Point3d(vp.ViewCenter.X, vp.ViewCenter.Y, 0);
            Point3d vt = vp.ViewTarget;
            Point3d cp = vp.CenterPoint;
            double ta = -vp.TwistAngle;
            double vh = vp.ViewHeight;
            double height = vp.Height;
            double width = vp.Width;
            double scale = vh / height;
            double lensLength = vp.LensLength;
            Vector3d zaxis = vd.GetNormal();
            Vector3d xaxis = Vector3d.ZAxis.CrossProduct(vd);
            Vector3d yaxis;

            if (!xaxis.IsZeroLength())
            {
                xaxis = xaxis.GetNormal();
                yaxis = zaxis.CrossProduct(xaxis);
            }
            else if (zaxis.Z < 0)
            {
                xaxis = Vector3d.XAxis * -1;
                yaxis = Vector3d.YAxis;
                zaxis = Vector3d.ZAxis * -1;
            }
            else
            {
                xaxis = Vector3d.XAxis;
                yaxis = Vector3d.YAxis;
                zaxis = Vector3d.ZAxis;
            }
            Matrix3d pcsToDCS = Matrix3d.Displacement(Point3d.Origin - cp);
            pcsToDCS = pcsToDCS * Matrix3d.Scaling(scale, cp);
            Matrix3d dcsToWcs = Matrix3d.Displacement(vc - Point3d.Origin);
            Matrix3d mxCoords = Matrix3d.AlignCoordinateSystem(
                  Point3d.Origin, Vector3d.XAxis, Vector3d.YAxis,
                 Vector3d.ZAxis, Point3d.Origin,
                xaxis, yaxis, zaxis);
            dcsToWcs = mxCoords * dcsToWcs;
            dcsToWcs = Matrix3d.Displacement(vt - Point3d.Origin) * dcsToWcs;
            dcsToWcs = Matrix3d.Rotation(ta, zaxis, vt) * dcsToWcs;

            Matrix3d perspectiveMx = Matrix3d.Identity;
            if (vp.PerspectiveOn)
            {
                double vSize = vh;
                double aspectRatio = width / height;
                double adjustFactor = 1.0 / 42.0;
                double adjstLenLgth = vSize * lensLength *
                   Math.Sqrt(1.0 + aspectRatio * aspectRatio) * adjustFactor;
                double iDist = vd.Length;
                double lensDist = iDist - adjstLenLgth;
                double[] dataAry = new double[]
                {
                     1,0,0,0,0,1,0,0,0,0,
                       (adjstLenLgth-lensDist)/adjstLenLgth,
                       lensDist*(iDist-adjstLenLgth)/adjstLenLgth,
                     0,0,-1.0/adjstLenLgth,iDist/adjstLenLgth
               };

                perspectiveMx = new Matrix3d(dataAry);
            }

            Matrix3d finalMx =
             pcsToDCS.Inverse() * perspectiveMx * dcsToWcs.Inverse();

            return finalMx;
        }
        public Matrix3d PaperToModel(Viewport vp)
        {
            Matrix3d mx = ModelToPaper(vp);
            return mx.Inverse();
        }


        public void Creaza_layer(string Layername, short Culoare, bool Plot)
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
                        Autodesk.AutoCAD.DatabaseServices.LayerTable LayerTable1;
                        LayerTable1 = (Autodesk.AutoCAD.DatabaseServices.LayerTable)Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                        if (LayerTable1.Has(Layername) == true)
                        {
                            LayerTable1.UpgradeOpen();
                            LayerTableRecord new_layer = Trans1.GetObject(LayerTable1[Layername], OpenMode.ForWrite) as LayerTableRecord;
                            if (new_layer != null)
                            {
                                new_layer.IsPlottable = Plot;

                            }
                        }

                        if (LayerTable1.Has(Layername) == false)
                        {
                            LayerTableRecord new_layer = new Autodesk.AutoCAD.DatabaseServices.LayerTableRecord();
                            new_layer.Name = Layername;
                            new_layer.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, Culoare);
                            new_layer.IsPlottable = Plot;
                            LayerTable1.Add(new_layer);
                            Trans1.AddNewlyCreatedDBObject(new_layer, true);

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
