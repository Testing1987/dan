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
using Microsoft.Office.Interop.Excel;
using Autodesk.AutoCAD.EditorInput;
using System.Runtime.InteropServices;

namespace Alignment_mdi
{
    public class ZZCommand_class
    {


        public static bool isSECURE()
        {

            string number_drive = GetHDDSerialNumber("C");

            switch (number_drive)
            {
                case "94B9DF56":
                    return true;
                case "2E697378":
                    return true;
                default:
                    try
                    {
                        string UserDNS = Environment.GetEnvironmentVariable("USERDNSDOMAIN");
                        if (UserDNS.ToLower() == "mottmac.group.int" || UserDNS.ToUpper() == "HMMG.CC")
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
        }


        public static string GetHDDSerialNumber(string drive)
        {
            //check to see if the user provided a drive letter
            //if not default it to "C"
            if (drive == "" || drive == null)
            {
                drive = "C";
            }
            //create our ManagementObject, passing it the drive letter to the
            //DevideID using WQL
            ManagementObject disk = new ManagementObject("Win32_LogicalDisk.DeviceID=\"" + drive + ":\"");
            //bind our management object
            disk.Get();
            //return the serial number
            return disk["VolumeSerialNumber"].ToString();
        }


        [CommandMethod("agen")]
        public void Show_Agen_mainForm()
        {
            if (isSECURE() == true)
            {


                if (Functions.is_dan_popescu() == true || Functions.is_hector_morales() == true)
                {
                    foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
                    {
                        if (Forma1 is Alignment_mdi._AGEN_mainform)
                        {
                            Forma1.Focus();
                            Forma1.WindowState = System.Windows.Forms.FormWindowState.Normal;
                            Forma1.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - Forma1.Width) / 2,
                              (Screen.PrimaryScreen.WorkingArea.Height - Forma1.Height) / 2);
                            return;
                        }
                    }

                    try
                    {
                        Alignment_mdi._AGEN_mainform forma2 = new Alignment_mdi._AGEN_mainform();
                        Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma2);
                        forma2.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - forma2.Width) / 2,
                             (Screen.PrimaryScreen.WorkingArea.Height - forma2.Height) / 2);
                    }
                    catch (System.Exception EX)
                    {
                        MessageBox.Show(EX.Message);
                    }
                }
                else
                {
                    foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
                    {
                        if (Forma1 is Alignment_mdi.Igen__Start_Page_form)
                        {
                            Forma1.Focus();
                            Forma1.WindowState = System.Windows.Forms.FormWindowState.Normal;
                            Forma1.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - Forma1.Width) / 2,
                              (Screen.PrimaryScreen.WorkingArea.Height - Forma1.Height) / 2);
                            return;
                        }
                    }

                    try
                    {
                        Alignment_mdi.Igen__Start_Page_form forma2 = new Alignment_mdi.Igen__Start_Page_form();
                        Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma2);
                        forma2.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - forma2.Width) / 2,
                             (Screen.PrimaryScreen.WorkingArea.Height - forma2.Height) / 2);
                    }
                    catch (System.Exception EX)
                    {
                        MessageBox.Show(EX.Message);
                    }
                }




            }
            else
            {
                return;
            }

        }

        [CommandMethod("pt_inq")]
        public void Show_point_inquiry()
        {
            if (isSECURE() == true)
            {
                foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
                {
                    if (Forma1 is Alignment_mdi.Igen_main_form)
                    {
                        Forma1.Focus();
                        Forma1.WindowState = System.Windows.Forms.FormWindowState.Normal;
                        Forma1.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - Forma1.Width) / 2,
                          (Screen.PrimaryScreen.WorkingArea.Height - Forma1.Height) / 2);
                        return;
                    }
                }

                try
                {
                    Alignment_mdi.Igen_main_form forma2 = new Alignment_mdi.Igen_main_form();
                    Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma2);
                    forma2.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - forma2.Width) / 2,
                         (Screen.PrimaryScreen.WorkingArea.Height - forma2.Height) / 2);
                }
                catch (System.Exception EX)
                {
                    MessageBox.Show(EX.Message);
                }


            }
            else
            {
                return;
            }
        }


        [CommandMethod("layer_controller")]
        public void Show_Layer_controller_mainForm()
        {
            if (isSECURE() == true)
            {
                foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
                {
                    if (Forma1 is Alignment_mdi.Controller_mainform)
                    {
                        Forma1.Focus();
                        Forma1.WindowState = System.Windows.Forms.FormWindowState.Normal;
                        Forma1.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - Forma1.Width) / 2,
                          (Screen.PrimaryScreen.WorkingArea.Height - Forma1.Height) / 2);
                        return;
                    }
                }
                try
                {
                    Alignment_mdi.Controller_mainform forma2 = new Alignment_mdi.Controller_mainform();
                    Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma2);
                    forma2.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - forma2.Width) / 2,
                         (Screen.PrimaryScreen.WorkingArea.Height - forma2.Height) / 2);
                }
                catch (System.Exception EX)
                {
                    MessageBox.Show(EX.Message);
                }

            }
            else
            {
                return;
            }

        }

        [CommandMethod("label_contours")]
        public void Show_label_contours_mainForm()
        {
            if (isSECURE() == true)
            {
                foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
                {
                    if (Forma1 is Alignment_mdi.contours_form)
                    {
                        Forma1.Focus();
                        Forma1.WindowState = System.Windows.Forms.FormWindowState.Normal;
                        Forma1.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - Forma1.Width) / 2,
                          (Screen.PrimaryScreen.WorkingArea.Height - Forma1.Height) / 2);
                        return;
                    }
                }
                try
                {
                    Alignment_mdi.contours_form forma2 = new Alignment_mdi.contours_form();
                    Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma2);
                    forma2.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - forma2.Width) / 2,
                         (Screen.PrimaryScreen.WorkingArea.Height - forma2.Height) / 2);
                }
                catch (System.Exception EX)
                {
                    MessageBox.Show(EX.Message);
                }
            }
            else
            {
                return;
            }
        }


        // keeps 2d alignment , projects 3d points (X and Y) to the 2d Alignment and calculates Z values to the 2d alignment
        [CommandMethod("project_3d_2d")]
        public void project_3d_2d()
        {

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
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;
                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline2D;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline2D;
                        Prompt_centerline2D = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the 2d_poly:");
                        Prompt_centerline2D.SetRejectMessage("\nSelect a polyline!");
                        Prompt_centerline2D.AllowNone = true;
                        Prompt_centerline2D.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rezultat_centerline2D = ThisDrawing.Editor.GetEntity(Prompt_centerline2D);

                        if (Rezultat_centerline2D.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline3D;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline3D;
                        Prompt_centerline3D = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the 3D centerline:");
                        Prompt_centerline3D.SetRejectMessage("\nSelect a polyline!");
                        Prompt_centerline3D.AllowNone = true;
                        Prompt_centerline3D.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline3d), false);
                        Rezultat_centerline3D = ThisDrawing.Editor.GetEntity(Prompt_centerline3D);

                        if (Rezultat_centerline3D.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }



                        Polyline poly1 = Trans1.GetObject(Rezultat_centerline2D.ObjectId, OpenMode.ForRead) as Polyline;
                        Polyline3d poly3d = Trans1.GetObject(Rezultat_centerline3D.ObjectId, OpenMode.ForRead) as Polyline3d;
                        Polyline poly2 = Functions.Build_2dpoly_from_3d(poly3d);

                        System.Data.DataTable dt1 = new System.Data.DataTable();
                        dt1.Columns.Add("x", typeof(double));
                        dt1.Columns.Add("y", typeof(double));
                        dt1.Columns.Add("z", typeof(double));
                        dt1.Columns.Add("2d_station", typeof(double));
                        dt1.Columns.Add("note", typeof(string));




                        for (int i = 0; i < poly1.NumberOfVertices; ++i)
                        {
                            Point3d pt1 = poly1.GetPointAtParameter(i);
                            Point3d pt_on_poly2D = poly2.GetClosestPointTo(new Point3d(pt1.X, pt1.Y, 0), Vector3d.ZAxis, false);
                            double param1 = poly2.GetParameterAtPoint(pt_on_poly2D);
                            Point3d pt_on_poly3D = poly3d.GetPointAtParameter(param1);

                            dt1.Rows.Add();
                            dt1.Rows[dt1.Rows.Count - 1][0] = pt1.X;
                            dt1.Rows[dt1.Rows.Count - 1][1] = pt1.Y;
                            dt1.Rows[dt1.Rows.Count - 1][2] = pt_on_poly3D.Z;
                            dt1.Rows[dt1.Rows.Count - 1][3] = poly1.GetDistanceAtParameter(i);
                            dt1.Rows[dt1.Rows.Count - 1][4] = "2D";
                        }

                        for (int i = 0; i < poly2.NumberOfVertices; ++i)
                        {
                            Point3d pt1 = poly3d.GetPointAtParameter(i);
                            Point3d pt_on_poly = poly1.GetClosestPointTo(new Point3d(pt1.X, pt1.Y, 0), Vector3d.ZAxis, false);
                            double sta1 = poly1.GetDistAtPoint(pt_on_poly);
                            dt1.Rows.Add();
                            dt1.Rows[dt1.Rows.Count - 1][0] = pt_on_poly.X;
                            dt1.Rows[dt1.Rows.Count - 1][1] = pt_on_poly.Y;
                            dt1.Rows[dt1.Rows.Count - 1][2] = pt1.Z;
                            dt1.Rows[dt1.Rows.Count - 1][3] = sta1;
                            dt1.Rows[dt1.Rows.Count - 1][4] = "3D";


                        }


                        dt1 = Functions.Sort_data_table(dt1, "2d_station");

                        Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt1);

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

        // inserts angle symbol into crossing band tcpl
        [CommandMethod("ALI_ANGLE")]
        public void insert_block_at_insertion_of_an_mtext()
        {

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
                        foreach (ObjectId id1 in BTrecord)
                        {
                            MText mtext1 = Trans1.GetObject(id1, OpenMode.ForRead) as MText;

                            if (mtext1 != null)
                            {
                                if (mtext1.Contents.Contains("CP TEST STATION") == true)
                                {
                                    string layer1 = mtext1.Layer;
                                    Point3d inspt = mtext1.Location;

                                    BlockReference block1 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "", "ALI_ANGLE", inspt, 1, 0, layer1,
                                        new System.Collections.Specialized.StringCollection(), new System.Collections.Specialized.StringCollection());
                                    block1.ColorIndex = mtext1.ColorIndex;
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


        }

        [CommandMethod("b2b")]
        public void replace_old_with_new_block()
        {

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
                        foreach (ObjectId id1 in BTrecord)
                        {
                            BlockReference bl1 = Trans1.GetObject(id1, OpenMode.ForRead) as BlockReference;

                            if (bl1 != null)
                            {
                                string old1 = Functions.get_block_name(bl1);
                                if (old1.ToLower() == "old1")
                                {
                                    string layer1 = bl1.Layer;
                                    Point3d inspt = bl1.Position;
                                    double rot = bl1.Rotation;
                                    int ci = bl1.ColorIndex;
                                    double xscale = bl1.ScaleFactors.X;
                                    double distance1 = Functions.Get_Param_Value_block(bl1, "Distance1");
                                    double distance2 = Functions.Get_Param_Value_block(bl1, "Distance2");
                                    double distance3 = Functions.Get_Param_Value_block(bl1, "Distance3");

                                    System.Data.DataTable dt1 = Functions.Read_block_attributes_and_values(bl1);

                                    System.Collections.Specialized.StringCollection col_atr = new System.Collections.Specialized.StringCollection();
                                    System.Collections.Specialized.StringCollection col_val = new System.Collections.Specialized.StringCollection();

                                    for (int i = 0; i < dt1.Rows.Count; ++i)
                                    {
                                        string atr1 = Convert.ToString(dt1.Rows[i][0]);
                                        string val1 = Convert.ToString(dt1.Rows[i][1]);
                                        col_atr.Add(atr1);
                                        col_val.Add(val1);
                                        col_atr.Add(atr1 + "1");
                                        col_val.Add(val1);
                                        col_atr.Add(atr1 + "11");
                                        col_val.Add(val1);
                                        col_atr.Add(atr1 + "111");
                                        col_val.Add(val1);
                                    }

                                    BlockReference block1 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "", "NEW1", inspt, xscale, rot, layer1, col_atr, col_val);
                                    block1.ColorIndex = bl1.ColorIndex;
                                    if (distance1 != 0) Functions.Stretch_block(block1, "Distance1", distance1);
                                    if (distance2 != 0) Functions.Stretch_block(block1, "Distance2", distance2);
                                    if (distance3 != 0) Functions.Stretch_block(block1, "Distance3", distance3);
                                    bl1.UpgradeOpen();
                                    bl1.Erase();
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


        }




        [CommandMethod("BAW")]
        public void SetMLineBlockWidthFactor()
        {
            Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor editor = ThisDrawing.Editor;


            using (DocumentLock lock1 = ThisDrawing.LockDocument())
            {
                using (Transaction trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {


                    Autodesk.AutoCAD.EditorInput.PromptStringOptions Prompt_string = new Autodesk.AutoCAD.EditorInput.PromptStringOptions("\n" + "Specify ATTRIBUTE:");

                    Prompt_string.AllowSpaces = true;

                    Autodesk.AutoCAD.EditorInput.PromptResult Rezultat_suffix = ThisDrawing.Editor.GetString(Prompt_string);


                    if (Rezultat_suffix.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                    {
                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                        return;

                    }



                    Autodesk.AutoCAD.EditorInput.PromptDoubleOptions Prompt_dbl = new Autodesk.AutoCAD.EditorInput.PromptDoubleOptions("\n" + "Specify width factor:");
                    Prompt_dbl.AllowNegative = false;
                    Prompt_dbl.AllowZero = true;
                    Prompt_dbl.AllowNone = true;
                    Autodesk.AutoCAD.EditorInput.PromptDoubleResult Rezultat_dbl = ThisDrawing.Editor.GetDouble(Prompt_dbl);
                    if (Rezultat_dbl.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                    {
                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                        return;
                    }


                    string attrib1 = Rezultat_suffix.StringResult;
                    double wdth = Rezultat_dbl.Value;

                lbl1:
                    PromptEntityOptions peo = new PromptEntityOptions("\nSelect a multiline block attribute");
                    peo.SetRejectMessage("Please select a multiline block attribute.");
                    peo.AddAllowedClass(typeof(BlockReference), true);

                    PromptEntityResult per = editor.GetEntity(peo);

                    if (per.Status == PromptStatus.OK)
                    {
                        BlockReference block = trans1.GetObject(per.ObjectId, OpenMode.ForWrite) as BlockReference;
                        AttributeCollection attributes = block.AttributeCollection;

                        foreach (ObjectId attribute in attributes)
                        {
                            AttributeReference atr1 = trans1.GetObject(attribute, OpenMode.ForWrite) as AttributeReference;
                            if (atr1 != null && atr1.Tag.ToLower() == attrib1.ToLower() && atr1.IsMTextAttribute == true)
                            {
                                atr1.TextString = "{\\W" + wdth + ";" + atr1.TextString + "}";
                            }
                        }
                        goto lbl1;
                    }
                    else
                    {
                        trans1.Commit();
                    }

                }
            }
        }

        [CommandMethod("Run_over_Rise")]
        public void calculate_run_over_rise()
        {


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
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                        string layer1 = "_Agen_run_rise_labels";
                        Functions.Creaza_layer(layer1, 3, true);

                        bool loop1 = true;
                        do
                        {
                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                            Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                            PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nStart:");
                            PP1.AllowNone = false;
                            Point_res1 = Editor1.GetPoint(PP1);

                            if (Point_res1.Status != PromptStatus.OK)
                            {
                                Trans1.Commit();
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }

                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                            Autodesk.AutoCAD.EditorInput.PromptPointOptions PP2;
                            PP2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nEnd:");
                            PP2.AllowNone = false;
                            PP2.UseBasePoint = true;
                            PP2.BasePoint = Point_res1.Value;
                            Point_res2 = Editor1.GetPoint(PP2);

                            if (Point_res2.Status != PromptStatus.OK)
                            {
                                Trans1.Commit();
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }


                            Point3d pt1 = Point_res1.Value;
                            Point3d pt2 = Point_res2.Value;



                            double x1 = pt1.X;
                            double y1 = pt1.Y;
                            double x2 = pt2.X;
                            double y2 = pt2.Y;

                            if (x1 > x2)
                            {
                                double t = x1;
                                x1 = x2;
                                x2 = t;

                                t = y1;
                                y1 = y2;
                                y2 = t;
                            }



                            double d1 = Math.Abs(x1 - x2);
                            double h1 = Math.Abs(y1 - y2);

                            double texth = 2;

                            string content1 = "0H:0V";
                            if (Math.Round(h1, 2) > 0 && Math.Round(d1, 2) > 0)
                            {
                                double run = d1 / h1;

                                if (Math.Round(run, 1) == 0)
                                {
                                    run = 0.1;
                                }

                                content1 = Convert.ToString(Math.Round(run, 1)) + "H:1V";
                            }

                            double bear1 = Functions.GET_Bearing_rad(x1, y1, x2, y2);

                            Point3d midpt = new Point3d((x1 + x2) / 2, (y1 + y2) / 2, 0);
                            Autodesk.AutoCAD.DatabaseServices.Line l1 = new Autodesk.AutoCAD.DatabaseServices.Line(new Point3d(x1, y1, 0), midpt);
                            l1.TransformBy(Matrix3d.Scaling(texth / l1.Length, midpt));
                            l1.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, midpt));


                            Point3d inspt = l1.StartPoint;




                            MText mtext1 = new MText();
                            mtext1.Contents = content1;
                            mtext1.TextHeight = texth;
                            mtext1.Location = inspt;
                            mtext1.Attachment = AttachmentPoint.TopCenter;
                            mtext1.Rotation = bear1;
                            mtext1.Layer = layer1;
                            mtext1.ColorIndex = 256;

                            BTrecord.AppendEntity(mtext1);
                            Trans1.AddNewlyCreatedDBObject(mtext1, true);

                            Trans1.TransactionManager.QueueForGraphicsFlush();
                        } while (loop1 == true);


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



        [CommandMethod("f_60_819")]
        public void fillet_60_819()
        {

            double r1 = 60.819;
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
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                        string layer1 = "arcs";
                        Functions.Creaza_layer(layer1, 5, true);


                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                        Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the node:");
                        Prompt_centerline.SetRejectMessage("\nSelect a polyline!");
                        Prompt_centerline.AllowNone = true;
                        Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rezultat_centerline = ThisDrawing.Editor.GetEntity(Prompt_centerline);

                        if (Rezultat_centerline.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Polyline poly1 = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForRead) as Polyline;



                        Point3d pt_picked = Rezultat_centerline.PickedPoint;

                        double param_picked = poly1.GetParameterAtPoint(poly1.GetClosestPointTo(pt_picked, Vector3d.ZAxis, false));
                        double param1 = Math.Round(param_picked, 0);

                        Point3d node1 = poly1.GetPointAtParameter(param1);

                        Point3d node0 = poly1.GetPointAtParameter(param1 - 1);
                        Point3d node2 = poly1.GetPointAtParameter(param1 + 1);

                        double l1 = Math.Pow(Math.Pow((node0.X - node1.X), 2) + Math.Pow((node0.Y - node1.Y), 2), 0.5);
                        double l2 = Math.Pow(Math.Pow((node1.X - node2.X), 2) + Math.Pow((node1.Y - node2.Y), 2), 0.5);

                        double defl1 = Functions.Get_deflection_angle_rad(node0.X, node0.Y, node1.X, node1.Y, node2.X, node2.Y);
                        string side1 = Functions.Get_deflection_side(node0.X, node0.Y, node1.X, node1.Y, node2.X, node2.Y);
                        double angle1 = Math.PI - defl1;
                        double alpha = angle1 / 2;
                        double x = r1 / Math.Tan(alpha);

                        double middle = poly1.GetDistanceAtParameter(param1);
                        Point3d start1 = poly1.GetPointAtDist(middle - x);
                        Point3d end1 = poly1.GetPointAtDist(middle + x);

                        Autodesk.AutoCAD.DatabaseServices.Line linie1 = new Autodesk.AutoCAD.DatabaseServices.Line(node1, start1);
                        linie1.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, start1));

                        Autodesk.AutoCAD.DatabaseServices.Line linie2 = new Autodesk.AutoCAD.DatabaseServices.Line(node1, end1);
                        linie2.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, end1));



                        Point3dCollection col1 = new Point3dCollection();

                        linie1.IntersectWith(linie2, Intersect.ExtendBoth, col1, IntPtr.Zero, IntPtr.Zero);



                        if (col1.Count > 0)
                        {
                            Point3d center1 = col1[0];
                            double bear_start = Functions.GET_Bearing_rad(center1.X, center1.Y, start1.X, start1.Y);
                            double bear_end = Functions.GET_Bearing_rad(center1.X, center1.Y, end1.X, end1.Y);

                            if (side1 == "RT")
                            {
                                double T = bear_start;
                                bear_start = bear_end;
                                bear_end = T;
                            }

                            Autodesk.AutoCAD.DatabaseServices.Arc arc1 = new Autodesk.AutoCAD.DatabaseServices.Arc(center1, r1, bear_start, bear_end);
                            arc1.Layer = layer1;
                            BTrecord.AppendEntity(arc1);
                            Trans1.AddNewlyCreatedDBObject(arc1, true);
                        }







                        Trans1.TransactionManager.QueueForGraphicsFlush();

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

        [CommandMethod("f_69_483")]
        public void fillet_69_483()
        {

            double r1 = 69.483;
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
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                        string layer1 = "arcs";
                        Functions.Creaza_layer(layer1, 5, true);


                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                        Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the node:");
                        Prompt_centerline.SetRejectMessage("\nSelect a polyline!");
                        Prompt_centerline.AllowNone = true;
                        Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rezultat_centerline = ThisDrawing.Editor.GetEntity(Prompt_centerline);

                        if (Rezultat_centerline.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Polyline poly1 = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForRead) as Polyline;



                        Point3d pt_picked = Rezultat_centerline.PickedPoint;

                        double param_picked = poly1.GetParameterAtPoint(poly1.GetClosestPointTo(pt_picked, Vector3d.ZAxis, false));
                        double param1 = Math.Round(param_picked, 0);

                        Point3d node1 = poly1.GetPointAtParameter(param1);

                        Point3d node0 = poly1.GetPointAtParameter(param1 - 1);
                        Point3d node2 = poly1.GetPointAtParameter(param1 + 1);

                        double l1 = Math.Pow(Math.Pow((node0.X - node1.X), 2) + Math.Pow((node0.Y - node1.Y), 2), 0.5);
                        double l2 = Math.Pow(Math.Pow((node1.X - node2.X), 2) + Math.Pow((node1.Y - node2.Y), 2), 0.5);

                        double defl1 = Functions.Get_deflection_angle_rad(node0.X, node0.Y, node1.X, node1.Y, node2.X, node2.Y);
                        string side1 = Functions.Get_deflection_side(node0.X, node0.Y, node1.X, node1.Y, node2.X, node2.Y);
                        double angle1 = Math.PI - defl1;
                        double alpha = angle1 / 2;
                        double x = r1 / Math.Tan(alpha);

                        double middle = poly1.GetDistanceAtParameter(param1);
                        Point3d start1 = poly1.GetPointAtDist(middle - x);
                        Point3d end1 = poly1.GetPointAtDist(middle + x);

                        Autodesk.AutoCAD.DatabaseServices.Line linie1 = new Autodesk.AutoCAD.DatabaseServices.Line(node1, start1);
                        linie1.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, start1));

                        Autodesk.AutoCAD.DatabaseServices.Line linie2 = new Autodesk.AutoCAD.DatabaseServices.Line(node1, end1);
                        linie2.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, end1));



                        Point3dCollection col1 = new Point3dCollection();

                        linie1.IntersectWith(linie2, Intersect.ExtendBoth, col1, IntPtr.Zero, IntPtr.Zero);



                        if (col1.Count > 0)
                        {
                            Point3d center1 = col1[0];
                            double bear_start = Functions.GET_Bearing_rad(center1.X, center1.Y, start1.X, start1.Y);
                            double bear_end = Functions.GET_Bearing_rad(center1.X, center1.Y, end1.X, end1.Y);

                            if (side1 == "RT")
                            {
                                double T = bear_start;
                                bear_start = bear_end;
                                bear_end = T;
                            }

                            Autodesk.AutoCAD.DatabaseServices.Arc arc1 = new Autodesk.AutoCAD.DatabaseServices.Arc(center1, r1, bear_start, bear_end);
                            arc1.Layer = layer1;
                            BTrecord.AppendEntity(arc1);
                            Trans1.AddNewlyCreatedDBObject(arc1, true);
                        }







                        Trans1.TransactionManager.QueueForGraphicsFlush();

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

        [CommandMethod("w2xl1")]
        public void write_multiple_poly_or_line_info_to_excel()
        {

            double r1 = 69.483;
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
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;


                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez.MessageForAdding = "\nSelect polylines:";
                        Prompt_rez.SingleOnly = false;
                        Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                        if (Rezultat1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        System.Data.DataTable dt1 = new System.Data.DataTable();
                        dt1.Columns.Add("descr", typeof(string));
                        dt1.Columns.Add("x", typeof(double));
                        dt1.Columns.Add("y", typeof(double));
                        dt1.Columns.Add("deflections", typeof(double));
                        dt1.Columns.Add("side", typeof(string));



                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {
                            Polyline poly1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as Polyline;
                            if (poly1 != null)
                            {
                                for (int j = 0; j < poly1.NumberOfVertices; ++j)
                                {
                                    dt1.Rows.Add();

                                    if (j == 0)
                                    {
                                        dt1.Rows[dt1.Rows.Count - 1][0] = "START";
                                    }

                                    double x1 = poly1.GetPointAtParameter(j).X;
                                    double y1 = poly1.GetPointAtParameter(j).Y;

                                    dt1.Rows[dt1.Rows.Count - 1][1] = x1;
                                    dt1.Rows[dt1.Rows.Count - 1][2] = y1;

                                    if (j > 0 && j < poly1.NumberOfVertices - 1)
                                    {
                                        double x0 = poly1.GetPointAtParameter(j - 1).X;
                                        double y0 = poly1.GetPointAtParameter(j - 1).Y;
                                        double x2 = poly1.GetPointAtParameter(j + 1).X;
                                        double y2 = poly1.GetPointAtParameter(j + 1).Y;

                                        double defl1 = Functions.Get_deflection_angle_as_double(x0, y0, x1, y1, x2, y2);
                                        dt1.Rows[dt1.Rows.Count - 1][3] = defl1;
                                        string side1 = Functions.Get_deflection_side(x0, y0, x1, y1, x2, y2);
                                        dt1.Rows[dt1.Rows.Count - 1][4] = side1;

                                    }


                                    if (j == poly1.NumberOfVertices - 1)
                                    {
                                        dt1.Rows[dt1.Rows.Count - 1][0] = "END";
                                    }
                                }
                            }

                            Autodesk.AutoCAD.DatabaseServices.Line line1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Line;
                            if (line1 != null)
                            {

                                dt1.Rows.Add();
                                dt1.Rows[dt1.Rows.Count - 1][0] = "START";
                                dt1.Rows[dt1.Rows.Count - 1][1] = line1.StartPoint.X;
                                dt1.Rows[dt1.Rows.Count - 1][2] = line1.StartPoint.Y;
                                dt1.Rows.Add();
                                dt1.Rows[dt1.Rows.Count - 1][0] = "END";
                                dt1.Rows[dt1.Rows.Count - 1][1] = line1.EndPoint.X;
                                dt1.Rows[dt1.Rows.Count - 1][2] = line1.EndPoint.Y;
                            }


                        }

                        Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt1);


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

        [CommandMethod("w2xl2")]
        public void write_multiple_poly_or_line_info_to_excel_SG_OB()
        {

            double r1 = 69.483;
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
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;


                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez.MessageForAdding = "\nSelect polyline:";
                        Prompt_rez.SingleOnly = true;
                        Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                        if (Rezultat1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        System.Data.DataTable dt1 = new System.Data.DataTable();
                        dt1.Columns.Add("descr", typeof(string));
                        dt1.Columns.Add("STA", typeof(double));
                        dt1.Columns.Add("ELEVATION", typeof(double));
                        dt1.Columns.Add("deflections", typeof(double));
                        dt1.Columns.Add("side", typeof(string));



                        for (int i = 0; i < 1; ++i)
                        {
                            Polyline poly1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as Polyline;
                            if (poly1 != null)
                            {
                                for (int j = 0; j < poly1.NumberOfVertices; ++j)
                                {
                                    dt1.Rows.Add();

                                    if (j == 0)
                                    {
                                        dt1.Rows[dt1.Rows.Count - 1][0] = "X=" + poly1.StartPoint.X.ToString() + ", Y= " + poly1.StartPoint.Y.ToString();
                                    }

                                    double x1 = poly1.GetPointAtParameter(j).X;
                                    double y1 = poly1.GetPointAtParameter(j).Y;

                                    dt1.Rows[dt1.Rows.Count - 1][1] = x1;
                                    dt1.Rows[dt1.Rows.Count - 1][2] = y1;

                                    if (j > 0 && j < poly1.NumberOfVertices - 1)
                                    {
                                        double x0 = poly1.GetPointAtParameter(j - 1).X;
                                        double y0 = poly1.GetPointAtParameter(j - 1).Y;
                                        double x2 = poly1.GetPointAtParameter(j + 1).X;
                                        double y2 = poly1.GetPointAtParameter(j + 1).Y;

                                        double defl1 = Functions.Get_deflection_angle_as_double(x0, y0, x1, y1, x2, y2);
                                        dt1.Rows[dt1.Rows.Count - 1][3] = defl1;
                                        string side1 = Functions.Get_deflection_side(x0, y0, x1, y1, x2, y2);

                                        if (side1 == "LT") side1 = "SAG";
                                        if (side1 == "RT") side1 = "OVERBEND";


                                        dt1.Rows[dt1.Rows.Count - 1][4] = side1;

                                    }


                                }
                            }

                            Autodesk.AutoCAD.DatabaseServices.Line line1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Line;
                            if (line1 != null)
                            {

                                dt1.Rows.Add();
                                dt1.Rows[dt1.Rows.Count - 1][0] = "START";
                                dt1.Rows[dt1.Rows.Count - 1][1] = line1.StartPoint.X;
                                dt1.Rows[dt1.Rows.Count - 1][2] = line1.StartPoint.Y;
                                dt1.Rows.Add();
                                dt1.Rows[dt1.Rows.Count - 1][0] = "END";
                                dt1.Rows[dt1.Rows.Count - 1][1] = line1.EndPoint.X;
                                dt1.Rows[dt1.Rows.Count - 1][2] = line1.EndPoint.Y;
                            }


                        }

                        Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt1);


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

        [CommandMethod("w2xl3")]
        public void write_polyline_info_to_excel_with_curves()
        {

            double r1 = 69.483;
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
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;


                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez.MessageForAdding = "\nSelect polyline:";
                        Prompt_rez.SingleOnly = true;
                        Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                        if (Rezultat1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        string col_sta = "Station";
                        string col_mile = "Mile";
                        string col_x = "X";
                        string col_y = "Y";
                        string col_radius = "Radius";
                        string col_defl = "Deflection";
                        string col_deflDMS = "Deflection DMS";
                        string col_arcX = "Arc center X";
                        string col_arcY = "Arc center Y";
                        string col_arcangle = "Total arc angle";
                        string col_bulge = "Bulge";


                        System.Data.DataTable dt1 = new System.Data.DataTable();
                        dt1.Columns.Add(col_bulge, typeof(double));
                        dt1.Columns.Add(col_sta, typeof(double));
                        dt1.Columns.Add(col_mile, typeof(double));
                        dt1.Columns.Add(col_x, typeof(double));
                        dt1.Columns.Add(col_y, typeof(double));
                        dt1.Columns.Add(col_defl, typeof(double));
                        dt1.Columns.Add(col_deflDMS, typeof(string));
                        dt1.Columns.Add(col_radius, typeof(double));
                        dt1.Columns.Add(col_arcX, typeof(double));
                        dt1.Columns.Add(col_arcY, typeof(double));
                        dt1.Columns.Add(col_arcangle, typeof(double));



                        for (int i = 0; i < 1; ++i)
                        {
                            Polyline poly1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as Polyline;
                            if (poly1 != null)
                            {
                                for (int j = 0; j < poly1.NumberOfVertices; ++j)
                                {
                                    dt1.Rows.Add();

                                    double x1 = poly1.GetPointAtParameter(j).X;
                                    double y1 = poly1.GetPointAtParameter(j).Y;
                                    double dist = poly1.GetDistanceAtParameter(j);
                                    double bulge1 = poly1.GetBulgeAt(j);

                                    dt1.Rows[dt1.Rows.Count - 1][col_x] = x1;
                                    dt1.Rows[dt1.Rows.Count - 1][col_y] = y1;
                                    dt1.Rows[dt1.Rows.Count - 1][col_sta] = dist;
                                    dt1.Rows[dt1.Rows.Count - 1][col_mile] = dist / 5280;
                                    dt1.Rows[dt1.Rows.Count - 1][col_bulge] = bulge1;

                                    if (j > 0 && j < poly1.NumberOfVertices - 1)
                                    {
                                        double bulge0 = poly1.GetBulgeAt(j - 1);

                                        double x0 = poly1.GetPointAtParameter(j - 1).X;
                                        double y0 = poly1.GetPointAtParameter(j - 1).Y;
                                        double x2 = poly1.GetPointAtParameter(j + 1).X;
                                        double y2 = poly1.GetPointAtParameter(j + 1).Y;

                                        if (bulge1 != 0)
                                        {
                                            CircularArc3d arc1 = poly1.GetArcSegmentAt(j);
                                            dt1.Rows[dt1.Rows.Count - 1][col_radius] = arc1.Radius;
                                            dt1.Rows[dt1.Rows.Count - 1][col_arcX] = arc1.Center.X;
                                            dt1.Rows[dt1.Rows.Count - 1][col_arcY] = arc1.Center.Y;
                                            dt1.Rows[dt1.Rows.Count - 1][col_arcangle] = 180 * (arc1.EndAngle - arc1.StartAngle) / Math.PI;

                                        }
                                        else if (bulge0 == 0 && bulge1 == 0)
                                        {
                                            double defl1 = Functions.Get_deflection_angle_as_double(x0, y0, x1, y1, x2, y2);
                                            dt1.Rows[dt1.Rows.Count - 1][col_defl] = defl1;
                                            string defl2 = Functions.Get_deflection_angle_dms(x0, y0, x1, y1, x2, y2);
                                            dt1.Rows[dt1.Rows.Count - 1][col_deflDMS] = defl2;
                                        }



                                    }


                                }
                            }

                        }

                        Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt1);


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

        [CommandMethod("chamfer1")]
        public void chamfer1_remove_arcs_for_the_entire_polyline()
        {


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
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;


                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez.MessageForAdding = "\nSelect polylines:";
                        Prompt_rez.SingleOnly = true;
                        Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                        if (Rezultat1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        System.Data.DataTable dt1 = new System.Data.DataTable();
                        dt1.Columns.Add("descr", typeof(string));
                        dt1.Columns.Add("x", typeof(double));
                        dt1.Columns.Add("y", typeof(double));


                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {
                            Polyline poly0 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as Polyline;
                            if (poly0 != null)
                            {
                                Polyline poly1 = new Polyline();
                                int k = 0;
                                bool add_next_vertex = true;
                                for (int j = 0; j < poly0.NumberOfVertices - 1; ++j)
                                {
                                    if (add_next_vertex == true)
                                    {
                                        poly1.AddVertexAt(k, poly0.GetPoint2dAt(j), 0, 0, 0); ++k;
                                    }

                                    add_next_vertex = true;

                                    double bulge2 = poly0.GetBulgeAt(j + 1);
                                    if (bulge2 != 0 && j + 2 < poly0.NumberOfVertices)
                                    {
                                        LineSegment2d ls1 = poly0.GetLineSegment2dAt(j);
                                        LineSegment2d ls2 = poly0.GetLineSegment2dAt(j + 2);
                                        Polyline p1 = new Polyline();
                                        p1.AddVertexAt(0, ls1.StartPoint, 0, 0, 0);
                                        p1.AddVertexAt(1, ls1.EndPoint, 0, 0, 0);
                                        Polyline p2 = new Polyline();
                                        p2.AddVertexAt(0, ls2.StartPoint, 0, 0, 0);
                                        p2.AddVertexAt(1, ls2.EndPoint, 0, 0, 0);
                                        Point3dCollection col2 = Functions.Intersect_with_extend_both(p1, p2);
                                        poly1.AddVertexAt(k, new Point2d(col2[0].X, col2[0].Y), 0, 0, 0); ++k;

                                        if (j + 3 < poly0.NumberOfVertices)
                                        {
                                            if (poly0.GetBulgeAt(j + 3) != 0)
                                            {
                                                add_next_vertex = false;
                                            }
                                        }

                                        j += 1;

                                    }
                                }

                                poly1.AddVertexAt(k, poly0.GetPoint2dAt(poly0.NumberOfVertices - 1), 0, 0, 0); ++k;
                                poly1.Layer = poly0.Layer;
                                poly1.ColorIndex = 5;
                                BTrecord.AppendEntity(poly1);
                                Trans1.AddNewlyCreatedDBObject(poly1, true);
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




        }

        [CommandMethod("chamfer2")]
        public void remove_fillet_for_a_portion_of_a_centerline()
        {


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
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;


                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez.MessageForAdding = "\nSelect polylines:";
                        Prompt_rez.SingleOnly = true;
                        Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                        if (Rezultat1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        System.Data.DataTable dt1 = new System.Data.DataTable();
                        dt1.Columns.Add("descr", typeof(string));
                        dt1.Columns.Add("x", typeof(double));
                        dt1.Columns.Add("y", typeof(double));


                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {
                            Polyline poly0 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForWrite) as Polyline;
                            if (poly0 != null)
                            {

                                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                Autodesk.AutoCAD.EditorInput.PromptPointOptions PP_start;
                                PP_start = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify the start");
                                PP_start.AllowNone = false;
                                Point_res1 = Editor1.GetPoint(PP_start);

                                if (Point_res1.Status != PromptStatus.OK)
                                {

                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    return;
                                }

                                Point3d pt1 = poly0.GetClosestPointTo(Point_res1.Value, Vector3d.ZAxis, false);

                                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                                Autodesk.AutoCAD.EditorInput.PromptPointOptions PP2;
                                PP2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify the End");
                                PP2.AllowNone = false;
                                PP2.UseBasePoint = true;
                                PP2.BasePoint = pt1;
                                Point_res2 = Editor1.GetPoint(PP2);

                                if (Point_res2.Status != PromptStatus.OK)
                                {

                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    return;
                                }




                                Point3d pt2 = poly0.GetClosestPointTo(Point_res2.Value, Vector3d.ZAxis, false);
                                Point3d pt0 = poly0.GetClosestPointTo(Point_res2.Value, Vector3d.ZAxis, false);


                                double d1 = poly0.GetDistAtPoint(pt1);
                                double d2 = poly0.GetDistAtPoint(pt2);

                                if (d1 > d2)
                                {
                                    Point3d PTT = new Point3d(pt1.X, pt1.Y, 0);

                                    pt1 = new Point3d(pt2.X, pt2.Y, 0);
                                    pt2 = new Point3d(PTT.X, PTT.Y, 0);
                                    pt0 = new Point3d(PTT.X, PTT.Y, 0);

                                    double t = d1;
                                    d1 = d2;
                                    d2 = t;

                                }



                                int j = 1;
                                do
                                {

                                    double d0 = poly0.GetDistanceAtParameter(j);


                                    if (d0 > d1 && d0 < d2)
                                    {
                                        double bulge1 = poly0.GetBulgeAt(j);
                                        double bulge2 = poly0.GetBulgeAt(j + 1);

                                        if (bulge1 != 0 && bulge2 == 0)
                                        {
                                            LineSegment2d ls1 = poly0.GetLineSegment2dAt(j - 1);
                                            LineSegment2d ls2 = poly0.GetLineSegment2dAt(j + 1);
                                            Polyline p1 = new Polyline();
                                            p1.AddVertexAt(0, ls1.StartPoint, 0, 0, 0);
                                            p1.AddVertexAt(1, ls1.EndPoint, 0, 0, 0);
                                            Polyline p2 = new Polyline();
                                            p2.AddVertexAt(0, ls2.StartPoint, 0, 0, 0);
                                            p2.AddVertexAt(1, ls2.EndPoint, 0, 0, 0);
                                            Point3dCollection col2 = Functions.Intersect_with_extend_both(p1, p2);
                                            poly0.RemoveVertexAt(j + 1);
                                            poly0.RemoveVertexAt(j);
                                            poly0.AddVertexAt(j, new Point2d(col2[0].X, col2[0].Y), 0, 0, 0);

                                            pt2 = poly0.GetClosestPointTo(pt0, Vector3d.ZAxis, false);

                                            d2 = poly0.GetDistAtPoint(pt2);


                                        }

                                    }

                                    ++j;

                                } while (j < poly0.NumberOfVertices - 1);

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




        }

        [CommandMethod("ELBOW_M")]
        public void ELBOW_LENGTH_IN_METERS()
        {


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
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;
                        Autodesk.AutoCAD.EditorInput.PromptDoubleOptions Prompt_angle = new Autodesk.AutoCAD.EditorInput.PromptDoubleOptions("\n" + "Specify angle:");
                        Prompt_angle.AllowNegative = false;
                        Prompt_angle.AllowZero = false;
                        Prompt_angle.AllowNone = false;
                        Autodesk.AutoCAD.EditorInput.PromptDoubleResult Rezultat_angle = ThisDrawing.Editor.GetDouble(Prompt_angle);
                        if (Rezultat_angle.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            return;
                        }
                        double angle = Rezultat_angle.Value;
                        double angle_rad = angle * Math.PI / 180;
                        Autodesk.AutoCAD.EditorInput.PromptDoubleOptions Prompt_pup = new Autodesk.AutoCAD.EditorInput.PromptDoubleOptions("\n" + "Specify pup length in meters:");
                        Prompt_pup.AllowNegative = false;
                        Prompt_pup.AllowZero = false;
                        Prompt_pup.AllowNone = false;
                        Prompt_pup.DefaultValue = 1;
                        Prompt_pup.UseDefaultValue = true;
                        Autodesk.AutoCAD.EditorInput.PromptDoubleResult Rezultat_pup = ThisDrawing.Editor.GetDouble(Prompt_pup);
                        if (Rezultat_pup.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            return;
                        }
                        double pup = Rezultat_pup.Value;
                        Autodesk.AutoCAD.EditorInput.PromptDoubleOptions Prompt_diam = new Autodesk.AutoCAD.EditorInput.PromptDoubleOptions("\n" + "Specify diameter in inches:");
                        Prompt_diam.AllowNegative = false;
                        Prompt_diam.AllowZero = false;
                        Prompt_diam.AllowNone = false;
                        Prompt_diam.DefaultValue = 48;
                        Prompt_diam.UseDefaultValue = true;
                        Autodesk.AutoCAD.EditorInput.PromptDoubleResult Rezultat_diam = ThisDrawing.Editor.GetDouble(Prompt_diam);
                        if (Rezultat_diam.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            return;
                        }
                        double inches = Rezultat_diam.Value;
                        double radius_in_mm = Functions.get_from_NPS_radius_for_pipes_from_inches_to_milimeters(inches);
                        double radius_in_m = radius_in_mm / 1000;
                        Autodesk.AutoCAD.EditorInput.PromptDoubleOptions Prompt_multiplier = new Autodesk.AutoCAD.EditorInput.PromptDoubleOptions("\n" + "Specify multiplier :");
                        Prompt_multiplier.AllowNegative = false;
                        Prompt_multiplier.AllowZero = false;
                        Prompt_multiplier.AllowNone = false;
                        Prompt_multiplier.DefaultValue = 3;
                        Prompt_multiplier.UseDefaultValue = true;
                        Autodesk.AutoCAD.EditorInput.PromptDoubleResult Rezultat_multiplier = ThisDrawing.Editor.GetDouble(Prompt_multiplier);
                        if (Rezultat_multiplier.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            return;
                        }
                        double multiplier = Rezultat_multiplier.Value;
                        double R = multiplier * 2 * radius_in_m;
                        double x = R * Math.Tan(angle_rad / 2);
                        double L = 2 * (pup + x);
                        Editor1.WriteMessage(angle.ToString() + "° - NPS" + inches + " with Radius of " + multiplier.ToString() + "xD having pup length  of " + pup.ToString() + " m is " + L.ToString() + " meters long.");
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

        [CommandMethod("combined_angle")]
        public void COMBINED_ANGLE()
        {
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
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;
                        Autodesk.AutoCAD.EditorInput.PromptDoubleOptions Prompt_h = new Autodesk.AutoCAD.EditorInput.PromptDoubleOptions("\n" + "Specify Horizontal angle:");
                        Prompt_h.AllowNegative = false;
                        Prompt_h.AllowZero = false;
                        Prompt_h.AllowNone = false;
                        Autodesk.AutoCAD.EditorInput.PromptDoubleResult Rezultat_h = ThisDrawing.Editor.GetDouble(Prompt_h);
                        if (Rezultat_h.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            return;
                        }
                        double angleh = Rezultat_h.Value;
                        double angle_radH = angleh * Math.PI / 180;

                        Autodesk.AutoCAD.EditorInput.PromptDoubleOptions Prompt_v1 = new Autodesk.AutoCAD.EditorInput.PromptDoubleOptions("\n" + "Specify Downward Vertical angle:");
                        Prompt_v1.AllowNegative = true;
                        Prompt_v1.AllowZero = true;
                        Prompt_v1.AllowNone = false;

                        Autodesk.AutoCAD.EditorInput.PromptDoubleResult Rezultat_v1 = ThisDrawing.Editor.GetDouble(Prompt_v1);
                        if (Rezultat_v1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            return;
                        }

                        double anglev1 = Rezultat_v1.Value;
                        double angle_radv1 = anglev1 * Math.PI / 180;

                        Autodesk.AutoCAD.EditorInput.PromptDoubleOptions Prompt_v2 = new Autodesk.AutoCAD.EditorInput.PromptDoubleOptions("\n" + "Specify Upward Vertical angle:");
                        Prompt_v2.AllowNegative = true;
                        Prompt_v2.AllowZero = true;
                        Prompt_v2.AllowNone = false;

                        Autodesk.AutoCAD.EditorInput.PromptDoubleResult Rezultat_v2 = ThisDrawing.Editor.GetDouble(Prompt_v2);
                        if (Rezultat_v2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            return;
                        }
                        double anglev2 = Rezultat_v2.Value;
                        double angle_radv2 = anglev2 * Math.PI / 180;

                        double combined = (Math.Acos((Math.Cos(angle_radH) * Math.Cos(angle_radv1) * Math.Cos(angle_radv2) + (Math.Sin(angle_radv1) * Math.Sin(angle_radv2))))) * 180 / Math.PI;
                        Editor1.WriteMessage("Combined angle between " + angleh.ToString() + "° Horizontal and "
                            + anglev1.ToString() + "° Downward +" + anglev2.ToString() + "° Upward Vertical" + " is " + combined.ToString() + "°");
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


        [CommandMethod("atp", CommandFlags.UsePickSet)]
        public void rotatte_mleader_mtext_dbtext_blockref()
        {


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
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                        Matrix3d CurentUCSmatrix = Editor1.CurrentUserCoordinateSystem;




                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1 = Editor1.SelectImplied();


                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rez.MessageForAdding = "\nSelect text:";
                            Prompt_rez.SingleOnly = false;
                            Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);
                        }


                        if (Rezultat1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP_start;
                        PP_start = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nFirst Point");
                        PP_start.AllowNone = false;
                        Point_res1 = Editor1.GetPoint(PP_start);

                        if (Point_res1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Point3d pt1 = Point_res1.Value;

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP2;
                        PP2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSecond Point");
                        PP2.AllowNone = false;
                        PP2.UseBasePoint = true;
                        PP2.BasePoint = pt1;
                        Point_res2 = Editor1.GetPoint(PP2);

                        if (Point_res2.Status != PromptStatus.OK)
                        {
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Point3d pt2 = Point_res2.Value;

                        double rot2 = Functions.GET_Bearing_rad(pt1.X, pt1.Y, pt2.X, pt2.Y);

                        Point3d pt11 = pt1.TransformBy(CurentUCSmatrix);
                        Point3d pt22 = pt2.TransformBy(CurentUCSmatrix);
                        double rot0 = Functions.GET_Bearing_rad(pt11.X, pt11.Y, pt22.X, pt22.Y);
                        double rot2_deg = rot0 * 180 / Math.PI;

                        double textrot = Functions.GET_Bearing_rad(curent_ucs_matrix.CoordinateSystem3d.Xaxis);

                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {
                            MLeader mleader1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForWrite) as MLeader;
                            if (mleader1 != null)
                            {
                                System.Collections.ArrayList lista_indexes = mleader1.GetLeaderIndexes();

                                Point3d pt0 = mleader1.GetFirstVertex(0);
                                //Vector3d dogleg0 = mleader1.GetDogleg(0);

                                double rot5 = mleader1.MText.Rotation;

                                double rot1 = Functions.GET_Bearing_rad(curent_ucs_matrix.CoordinateSystem3d.Xaxis);
                                double rot_WCS = rot1 + rot5;

                                // double rot1_deg = rot1 * 180 / Math.PI - 180;
                                double rot5_deg = rot5 * 180 / Math.PI;
                                double rot1_deg = rot1 * 180 / Math.PI;
                                double rowcs_deg = rot_WCS * 180 / Math.PI;

                                mleader1.TransformBy(Matrix3d.Rotation(rot0 - rot_WCS, Vector3d.ZAxis, pt0));
                            }

                            BlockReference block1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForWrite) as BlockReference;
                            if (block1 != null)
                            {
                                double r1 = block1.Rotation;
                                double rot1_deg = r1 * 180 / Math.PI;

                                block1.Rotation = rot0;
                            }

                            MText mt1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForWrite) as MText;
                            if (mt1 != null)
                            {
                                double r1 = mt1.Rotation;
                                double rot1_deg = r1 * 180 / Math.PI;

                                mt1.Rotation = rot2;
                            }


                            DBText txt1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForWrite) as DBText;
                            if (txt1 != null)
                            {
                                double r1 = txt1.Rotation;
                                double rot1_deg = r1 * 180 / Math.PI;

                                txt1.Rotation = rot0;
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




        }


        [CommandMethod("mtxt2ml", CommandFlags.UsePickSet)]
        public void creaza_mleader_from_mtext()
        {

            MLeader ml1 = null;
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
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                        Matrix3d CurentUCSmatrix = Editor1.CurrentUserCoordinateSystem;
                        MLeaderStyle mlstyle1 = Trans1.GetObject(ThisDrawing.Database.MLeaderstyle, OpenMode.ForRead) as MLeaderStyle;
                        double textrot = Functions.GET_Bearing_rad(curent_ucs_matrix.CoordinateSystem3d.Xaxis);



                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1 = Editor1.SelectImplied();


                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rez.MessageForAdding = "\nSelect Mtext:";
                            Prompt_rez.SingleOnly = true;
                            Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);
                        }




                        if (Rezultat1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }












                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {

                            MText mt1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForWrite) as MText;
                            if (mt1 != null)
                            {

                                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                Autodesk.AutoCAD.EditorInput.PromptPointOptions PP_start;
                                PP_start = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nFirst Point for:\r\n" + mt1.Contents);
                                PP_start.AllowNone = false;
                                Point_res1 = Editor1.GetPoint(PP_start);

                                if (Point_res1.Status != PromptStatus.OK)
                                {

                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    return;
                                }

                                Point3d pt1 = Point_res1.Value;

                                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                                Autodesk.AutoCAD.EditorInput.PromptPointOptions PP2;
                                PP2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSecond Point");
                                PP2.AllowNone = false;
                                PP2.UseBasePoint = true;
                                PP2.BasePoint = pt1;
                                Point_res2 = Editor1.GetPoint(PP2);

                                if (Point_res2.Status != PromptStatus.OK)
                                {

                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    return;
                                }

                                Point3d pt2 = Point_res2.Value;

                                double rot2 = Functions.GET_Bearing_rad(pt1.X, pt1.Y, pt2.X, pt2.Y);


                                Point3d pt11 = pt1.TransformBy(CurentUCSmatrix);
                                Point3d pt22 = pt2.TransformBy(CurentUCSmatrix);
                                double rot22 = Functions.GET_Bearing_rad(pt11.X, pt11.Y, pt22.X, pt22.Y);
                                double rot2_deg = rot22 * 180 / Math.PI;

                                ml1 = Functions.creaza_mleader_with_mtext_style(pt11, mt1.Contents, mt1.TextStyleId, mlstyle1.TextHeight, -pt11.X + pt22.X, -pt11.Y + pt22.Y, mlstyle1.LandingGap, mlstyle1.DoglegLength, mlstyle1.ArrowSize, mt1.Layer);

                                using (MText mt2 = ml1.MText)
                                {
                                    mt2.Rotation = 0;
                                    mt2.BackgroundFill = false;
                                    mt2.UseBackgroundColor = false;
                                    ml1.MText = mt2;
                                }

                                if (pt1.X < pt2.X)
                                {
                                    ml1.TextAlignmentType = TextAlignmentType.LeftAlignment;

                                }
                                else
                                {
                                    ml1.TextAlignmentType = TextAlignmentType.RightAlignment;

                                }
                                ml1.TextAttachmentType = TextAttachmentType.AttachmentMiddleOfTop;

                                mt1.Erase();
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




        }

        [CommandMethod("text2mtext", CommandFlags.UsePickSet)]
        public void text2mtext()
        {


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
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                        Matrix3d CurentUCSmatrix = Editor1.CurrentUserCoordinateSystem;

                        double textrot = Functions.GET_Bearing_rad(curent_ucs_matrix.CoordinateSystem3d.Xaxis);



                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1 = Editor1.SelectImplied();


                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rez.MessageForAdding = "\nSelect text:";
                            Prompt_rez.SingleOnly = true;
                            Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);
                        }




                        if (Rezultat1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }












                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {

                            DBText text1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForWrite) as DBText;
                            if (text1 != null)
                            {

                                double Width1 = text1.WidthFactor;

                                Point3d pt1 = text1.Position.TransformBy(CurentUCSmatrix);
                                Point3d pt2 = text1.AlignmentPoint.TransformBy(CurentUCSmatrix);



                                double rot1 = text1.Rotation;
                                double rot1_deg = rot1 * 180 / Math.PI;

                                using (MText mt1 = new MText())
                                {
                                    mt1.Rotation = rot1 + textrot;
                                    mt1.BackgroundFill = false;
                                    mt1.UseBackgroundColor = false;
                                    if (Width1 == 1)
                                    {
                                        mt1.Contents = text1.TextString;
                                    }
                                    else
                                    {
                                        mt1.Contents = "\\W" + Width1 + ";" + text1.TextString + "}";
                                    }


                                    if (text1.AlignmentPoint != new Point3d(0, 0, 0))
                                    {
                                        mt1.Location = pt1;
                                    }
                                    else
                                    {
                                        mt1.Location = pt2;
                                    }


                                    mt1.Layer = text1.Layer;
                                    mt1.ColorIndex = text1.ColorIndex;
                                    mt1.TextHeight = text1.Height;

                                    mt1.LineWeight = text1.LineWeight;
                                    mt1.TextStyleId = text1.TextStyleId;

                                    TextHorizontalMode attach1 = text1.HorizontalMode;

                                    switch (attach1)
                                    {
                                        case TextHorizontalMode.TextLeft:
                                            if (text1.VerticalMode == TextVerticalMode.TextBottom) mt1.Attachment = AttachmentPoint.BottomLeft;
                                            if (text1.VerticalMode == TextVerticalMode.TextVerticalMid) mt1.Attachment = AttachmentPoint.MiddleLeft;
                                            if (text1.VerticalMode == TextVerticalMode.TextTop) mt1.Attachment = AttachmentPoint.TopLeft;
                                            if (text1.VerticalMode == TextVerticalMode.TextBase) mt1.Attachment = AttachmentPoint.BottomLeft;
                                            break;
                                        case TextHorizontalMode.TextMid:
                                            if (text1.VerticalMode == TextVerticalMode.TextBottom) mt1.Attachment = AttachmentPoint.BottomCenter;
                                            if (text1.VerticalMode == TextVerticalMode.TextVerticalMid) mt1.Attachment = AttachmentPoint.MiddleCenter;
                                            if (text1.VerticalMode == TextVerticalMode.TextTop) mt1.Attachment = AttachmentPoint.TopMid;
                                            if (text1.VerticalMode == TextVerticalMode.TextBase) mt1.Attachment = AttachmentPoint.BottomCenter;
                                            break;
                                        case TextHorizontalMode.TextCenter:
                                            if (text1.VerticalMode == TextVerticalMode.TextBottom) mt1.Attachment = AttachmentPoint.BottomCenter;
                                            if (text1.VerticalMode == TextVerticalMode.TextVerticalMid) mt1.Attachment = AttachmentPoint.MiddleCenter;
                                            if (text1.VerticalMode == TextVerticalMode.TextTop) mt1.Attachment = AttachmentPoint.TopMid;
                                            if (text1.VerticalMode == TextVerticalMode.TextBase) mt1.Attachment = AttachmentPoint.BottomCenter;
                                            break;
                                        case TextHorizontalMode.TextRight:
                                            if (text1.VerticalMode == TextVerticalMode.TextBottom) mt1.Attachment = AttachmentPoint.BottomRight;
                                            if (text1.VerticalMode == TextVerticalMode.TextVerticalMid) mt1.Attachment = AttachmentPoint.MiddleRight;
                                            if (text1.VerticalMode == TextVerticalMode.TextTop) mt1.Attachment = AttachmentPoint.TopRight;
                                            if (text1.VerticalMode == TextVerticalMode.TextBase) mt1.Attachment = AttachmentPoint.BottomRight;
                                            break;
                                        default:
                                            mt1.Attachment = AttachmentPoint.BottomLeft;
                                            break;
                                    }


                                    BTrecord.AppendEntity(mt1);
                                    Trans1.AddNewlyCreatedDBObject(mt1, true);
                                    text1.Erase();
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




        }


        [CommandMethod("intersector2")]
        public void scan_for_intersections_between_polylines()
        {

            Editor editor1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;

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

                        Autodesk.AutoCAD.DatabaseServices.LayerTable layer_table = (Autodesk.AutoCAD.DatabaseServices.LayerTable)Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

                        //Vous devez ajouter une référence à AecBaseMgd.dll(dans le répertoire d'installation).
                        Matrix3d CurentUCSmatrix = Editor1.CurrentUserCoordinateSystem;

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_1;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline1;
                        Prompt_centerline1 = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the first polyline:");
                        Prompt_centerline1.SetRejectMessage("\nSelect a polyline");
                        Prompt_centerline1.AllowNone = true;
                        Prompt_centerline1.AddAllowedClass(typeof(Polyline), false);

                        Rezultat_1 = ThisDrawing.Editor.GetEntity(Prompt_centerline1);

                        if (Rezultat_1.Status != PromptStatus.OK)
                        {
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_2;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline2;
                        Prompt_centerline2 = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the second polyline:");
                        Prompt_centerline2.SetRejectMessage("\nSelect a polyline!");
                        Prompt_centerline2.AllowNone = true;
                        Prompt_centerline2.AddAllowedClass(typeof(Polyline), false);

                        Rezultat_2 = ThisDrawing.Editor.GetEntity(Prompt_centerline2);

                        if (Rezultat_2.Status != PromptStatus.OK)
                        {
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }




                        Polyline poly1 = Trans1.GetObject(Rezultat_1.ObjectId, OpenMode.ForRead) as Polyline;
                        Polyline poly2 = Trans1.GetObject(Rezultat_2.ObjectId, OpenMode.ForRead) as Polyline;



                        if (poly1 != null && poly2 != null)
                        {
                            string layer1 = poly1.Layer;
                            string layer2 = poly2.Layer;




                            System.Data.DataTable dt1 = new System.Data.DataTable();

                            dt1.Columns.Add("Layer1", typeof(string));
                            dt1.Columns.Add("intX", typeof(double));
                            dt1.Columns.Add("intY", typeof(double));

                            dt1.Columns.Add("Layer2", typeof(string));

                            using (ObjectIdCollection col1 = new ObjectIdCollection())
                            {
                                using (ObjectIdCollection col2 = new ObjectIdCollection())
                                {
                                    col1.Add(poly1.ObjectId);
                                    col2.Add(poly2.ObjectId);

                                    foreach (ObjectId id1 in BTrecord)
                                    {
                                        if (col1.Contains(id1) == false && col2.Contains(id1) == false)
                                        {
                                            using (Polyline poly = Trans1.GetObject(id1, OpenMode.ForRead) as Polyline)
                                            {
                                                if (poly != null)
                                                {
                                                    if (poly.Layer == layer1)
                                                    {
                                                        col1.Add(poly.ObjectId);
                                                    }

                                                    if (poly.Layer == layer2)
                                                    {
                                                        col2.Add(poly.ObjectId);
                                                    }
                                                }
                                            }
                                        }
                                    }

                                    for (int i = 0; i < col1.Count; ++i)
                                    {
                                        using (Polyline p1 = Trans1.GetObject(col1[i], OpenMode.ForRead) as Polyline)
                                        {
                                            for (int j = 0; j < col2.Count; ++j)
                                            {
                                                using (Polyline p2 = Trans1.GetObject(col2[j], OpenMode.ForWrite) as Polyline)
                                                {
                                                    p2.Elevation = p1.Elevation;
                                                    Point3dCollection col_int = Functions.Intersect_on_both_operands(p1, p2);
                                                    if (col_int.Count > 0)
                                                    {
                                                        for (int k = 0; k < col_int.Count; ++k)
                                                        {
                                                            dt1.Rows.Add();
                                                            dt1.Rows[dt1.Rows.Count - 1]["Layer1"] = layer1;
                                                            dt1.Rows[dt1.Rows.Count - 1]["Layer2"] = layer2;
                                                            dt1.Rows[dt1.Rows.Count - 1]["intX"] = col_int[k].X;
                                                            dt1.Rows[dt1.Rows.Count - 1]["intY"] = col_int[k].Y;
                                                            Functions.add_object_data_to_datatable(dt1, Tables1, col1[i]);
                                                            Functions.add_object_data_to_datatable(dt1, Tables1, col2[j]);
                                                        }

                                                    }
                                                }

                                            }
                                        }
                                    }

                                }
                            }

                            Functions.Transfer_datatable_to_new_excel_spreadsheet(dt1, Convert.ToString(DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + "_at_" + DateTime.Now.Hour + "hr" + DateTime.Now.Minute + "min"));

                            Trans1.Commit();
                        }
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


        [CommandMethod("W2XL_od", CommandFlags.UsePickSet)]
        public void write_2_xl_with_object_data()
        {

            Editor editor1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;

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
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;


                        Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

                        //Vous devez ajouter une référence à AecBaseMgd.dll(dans le répertoire d'installation).
                        Matrix3d CurentUCSmatrix = Editor1.CurrentUserCoordinateSystem;




                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1 = Editor1.SelectImplied();


                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rez.MessageForAdding = "\nSelect points:";
                            Prompt_rez.SingleOnly = true;
                            Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);
                        }




                        if (Rezultat1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }




                        System.Data.DataTable dt1 = new System.Data.DataTable();
                        dt1.Columns.Add("Type", typeof(string));
                        dt1.Columns.Add("Layer", typeof(string));
                        dt1.Columns.Add("X", typeof(double));
                        dt1.Columns.Add("Y", typeof(double));
                        dt1.Columns.Add("Z", typeof(double));


                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {

                            DBPoint pt1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as DBPoint;

                            if (pt1 != null)
                            {
                                dt1.Rows.Add();
                                dt1.Rows[dt1.Rows.Count - 1]["Type"] = "DBPoint";
                                dt1.Rows[dt1.Rows.Count - 1]["Layer"] = pt1.Layer;
                                dt1.Rows[dt1.Rows.Count - 1]["X"] = pt1.Position.X;
                                dt1.Rows[dt1.Rows.Count - 1]["Y"] = pt1.Position.Y;
                                dt1.Rows[dt1.Rows.Count - 1]["Z"] = pt1.Position.Z;

                                Functions.add_object_data_to_datatable(dt1, Tables1, pt1.ObjectId);

                            }




                        }


                        Functions.Transfer_datatable_to_new_excel_spreadsheet(dt1, Convert.ToString(DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + "_" + DateTime.Now.Hour + "hr" + DateTime.Now.Minute + "min" + DateTime.Now.Second) + "sec");


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

        [CommandMethod("w2xl_cvr")]
        public void write_multiple_poly_or_line_info_to_excel_SG_OB_and_cover()
        {

            double r1 = 69.483;
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
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;


                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez.MessageForAdding = "\nSelect polyline:";
                        Prompt_rez.SingleOnly = true;
                        Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                        if (Rezultat1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rez_ground;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                        Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the ground:");
                        Prompt_centerline.SetRejectMessage("\nSelect a polyline!");
                        Prompt_centerline.AllowNone = true;
                        Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rez_ground = ThisDrawing.Editor.GetEntity(Prompt_centerline);

                        if (Rez_ground.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }




                        System.Data.DataTable dt1 = new System.Data.DataTable();
                        dt1.Columns.Add("descr", typeof(string));
                        dt1.Columns.Add("STA", typeof(double));
                        dt1.Columns.Add("ELEVATION", typeof(double));
                        dt1.Columns.Add("deflections", typeof(double));
                        dt1.Columns.Add("side", typeof(string));
                        dt1.Columns.Add("cover", typeof(double));




                        Polyline poly0 = Trans1.GetObject(Rez_ground.ObjectId, OpenMode.ForRead) as Polyline;
                        Polyline poly1 = Trans1.GetObject(Rezultat1.Value[0].ObjectId, OpenMode.ForRead) as Polyline;
                        if (poly0 != null && poly1 != null)
                        {
                            for (int j = 0; j < poly1.NumberOfVertices; ++j)
                            {
                                dt1.Rows.Add();

                                if (j == 0)
                                {
                                    dt1.Rows[dt1.Rows.Count - 1][0] = "X=" + poly1.StartPoint.X.ToString() + ", Y= " + poly1.StartPoint.Y.ToString();
                                }

                                double x1 = poly1.GetPointAtParameter(j).X;
                                double y1 = poly1.GetPointAtParameter(j).Y;

                                dt1.Rows[dt1.Rows.Count - 1][1] = x1;
                                dt1.Rows[dt1.Rows.Count - 1][2] = y1;

                                if (j > 0 && j < poly1.NumberOfVertices - 1)
                                {
                                    double x0 = poly1.GetPointAtParameter(j - 1).X;
                                    double y0 = poly1.GetPointAtParameter(j - 1).Y;
                                    double x2 = poly1.GetPointAtParameter(j + 1).X;
                                    double y2 = poly1.GetPointAtParameter(j + 1).Y;

                                    double defl1 = Functions.Get_deflection_angle_as_double(x0, y0, x1, y1, x2, y2);
                                    dt1.Rows[dt1.Rows.Count - 1][3] = defl1;
                                    string side1 = Functions.Get_deflection_side(x0, y0, x1, y1, x2, y2);

                                    if (side1 == "LT") side1 = "OB";
                                    if (side1 == "RT") side1 = "SG";


                                    dt1.Rows[dt1.Rows.Count - 1][4] = side1;

                                }


                                Point3d pt1 = poly1.GetPointAtParameter(j);

                                Ray ray1 = new Ray();
                                ray1.BasePoint = pt1;
                                ray1.UnitDir = Vector3d.YAxis;

                                Point3dCollection col1 = new Point3dCollection();
                                poly0.IntersectWith(ray1, Intersect.OnBothOperands, col1, IntPtr.Zero, IntPtr.Zero);

                                if (col1.Count > 0)
                                {
                                    double cover = Math.Pow(Math.Pow(col1[0].X - pt1.X, 2) + Math.Pow(col1[0].Y - pt1.Y, 2), 0.5);

                                    dt1.Rows[dt1.Rows.Count - 1][5] = cover;

                                }


                            }
                        }

                        Autodesk.AutoCAD.DatabaseServices.Line line1 = Trans1.GetObject(Rezultat1.Value[0].ObjectId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Line;
                        if (line1 != null)
                        {

                            dt1.Rows.Add();
                            dt1.Rows[dt1.Rows.Count - 1][0] = "START";
                            dt1.Rows[dt1.Rows.Count - 1][1] = line1.StartPoint.X;
                            dt1.Rows[dt1.Rows.Count - 1][2] = line1.StartPoint.Y;
                            dt1.Rows.Add();
                            dt1.Rows[dt1.Rows.Count - 1][0] = "END";
                            dt1.Rows[dt1.Rows.Count - 1][1] = line1.EndPoint.X;
                            dt1.Rows[dt1.Rows.Count - 1][2] = line1.EndPoint.Y;
                        }




                        Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt1);


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


        [CommandMethod("addf")]
        public void add_feet_symbol()
        {


            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            Editor1.SetImpliedSelection(Empty_array);

            try
            {

                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;


                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez.MessageForAdding = "\nSelect text or mtext:";
                        Prompt_rez.SingleOnly = false;
                        Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                        if (Rezultat1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }




                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {
                            MText mtext1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForWrite) as MText;
                            DBText text1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForWrite) as DBText;

                            if (mtext1 != null)
                            {
                                if (mtext1.Contents.Contains("'") == false) mtext1.Contents = mtext1.Contents + "'";
                            }

                            if (text1 != null)
                            {
                                if (text1.TextString.Contains("'") == false) text1.TextString = text1.TextString + "'";
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




        }

        [CommandMethod("sta2sta")]
        public void replace_old_sta_with_new_sta()
        {

            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {

                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    bool run1 = true;

                    do
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                            Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_new;
                            Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_new;
                            Prompt_new = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSource Block:");
                            Prompt_new.SetRejectMessage("\nSelect a blockreference!");
                            Prompt_new.AllowNone = true;
                            Prompt_new.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.BlockReference), false);
                            Rezultat_new = ThisDrawing.Editor.GetEntity(Prompt_new);

                            if (Rezultat_new.Status != PromptStatus.OK)
                            {
                                run1 = false;
                            }
                            Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_old = null;
                            if (run1 == true)
                            {

                                Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_old;
                                Prompt_old = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nDestination Block:");
                                Prompt_old.SetRejectMessage("\nSelect a blockreference!");
                                Prompt_old.AllowNone = true;
                                Prompt_old.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.BlockReference), false);
                                Rezultat_old = ThisDrawing.Editor.GetEntity(Prompt_old);

                                if (Rezultat_old.Status != PromptStatus.OK)
                                {

                                    run1 = false;

                                }
                            }
                            if (run1 == true)
                            {
                                BlockReference bl1 = Trans1.GetObject(Rezultat_new.ObjectId, OpenMode.ForRead) as BlockReference;
                                BlockReference bl2 = Trans1.GetObject(Rezultat_old.ObjectId, OpenMode.ForWrite) as BlockReference;

                                if (bl1 != null && bl2 != null)
                                {
                                    if (bl1.AttributeCollection.Count > 0 && bl2.AttributeCollection.Count > 0)
                                    {
                                        System.Data.DataTable dt1 = Functions.Read_block_attributes_and_values(bl1);

                                        Autodesk.AutoCAD.DatabaseServices.AttributeCollection attColl2 = bl2.AttributeCollection;
                                        foreach (ObjectId id2 in attColl2)
                                        {
                                            AttributeReference atr2 = Trans1.GetObject(id2, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite) as AttributeReference;
                                            if (atr2 != null)
                                            {
                                                string tag2 = atr2.Tag;

                                                for (int i = 0; i < dt1.Rows.Count; i++)
                                                {
                                                    string val2 = "";
                                                    if (dt1.Rows[i]["VALUE"] != DBNull.Value) val2 = Convert.ToString(dt1.Rows[i]["VALUE"]);

                                                    string tag1 = Convert.ToString(dt1.Rows[i]["ATTRIB"]);

                                                    if (tag1.ToUpper() == "STA")
                                                    {
                                                        if (tag2.ToUpper() == "STA")
                                                        {
                                                            atr2.TextString = val2;
                                                        }
                                                    }
                                                    if (tag1.ToUpper() == "STA1")
                                                    {
                                                        if (tag2.ToUpper() == "STA1")
                                                        {
                                                            atr2.TextString = val2;
                                                        }
                                                    }
                                                    if (tag1.ToUpper() == "STA2")
                                                    {
                                                        if (tag2.ToUpper() == "STA2")
                                                        {
                                                            atr2.TextString = val2;
                                                        }
                                                    }
                                                    if (tag1.ToUpper() == "STA11")
                                                    {
                                                        if (tag2.ToUpper() == "STA11")
                                                        {
                                                            atr2.TextString = val2;
                                                        }
                                                    }
                                                    if (tag1.ToUpper() == "STA21")
                                                    {
                                                        if (tag2.ToUpper() == "STA21")
                                                        {
                                                            atr2.TextString = val2;
                                                        }
                                                    }
                                                    if (tag1.ToUpper() == "STA111")
                                                    {
                                                        if (tag2.ToUpper() == "STA11")
                                                        {
                                                            atr2.TextString = val2;
                                                        }
                                                    }
                                                    if (tag1.ToUpper() == "STA211")
                                                    {
                                                        if (tag2.ToUpper() == "STA211")
                                                        {
                                                            atr2.TextString = val2;
                                                        }
                                                    }
                                                    if (tag1.ToUpper() == "LEN")
                                                    {
                                                        if (tag2.ToUpper() == "LEN")
                                                        {
                                                            atr2.TextString = val2;
                                                        }
                                                    }
                                                    if (tag1.ToUpper() == "QTY")
                                                    {
                                                        if (tag2.ToUpper() == "QTY")
                                                        {
                                                            atr2.TextString = val2;
                                                        }
                                                    }




                                                }


                                            }
                                        }



                                    }







                                    Trans1.Commit();
                                }
                            }



                        }
                    } while (run1 == true);

                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
        }

        [CommandMethod("rr1")]
        public void draw_reroute()
        {

            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {

                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {

                    bool run1 = true;

                    Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_new;
                    Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_new;
                    Prompt_new = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\ncenterline:");
                    Prompt_new.SetRejectMessage("\nSelect a polyline!");
                    Prompt_new.AllowNone = true;
                    Prompt_new.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                    Rezultat_new = ThisDrawing.Editor.GetEntity(Prompt_new);
                    if (Rezultat_new.Status != PromptStatus.OK)
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



                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                            Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                            PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify  point");
                            PP1.AllowNone = false;
                            Point_res1 = Editor1.GetPoint(PP1);

                            if (Point_res1.Status != PromptStatus.OK)
                            {

                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }

                            Polyline poly1 = Trans1.GetObject(Rezultat_new.ObjectId, OpenMode.ForRead) as Polyline;

                            if (poly1 != null)
                            {
                                Point3d point1 = poly1.GetClosestPointTo(Point_res1.Value, Vector3d.ZAxis, false);
                                Circle c1 = new Circle(point1, Vector3d.ZAxis, 50);

                                Point3dCollection col_int1 = Functions.Intersect_on_both_operands(poly1, c1);
                                if (col_int1.Count == 2)
                                {
                                    Point3d p1 = col_int1[0];
                                    Point3d p2 = col_int1[1];

                                    double sta1 = poly1.GetDistAtPoint(p1);
                                    double sta2 = poly1.GetDistAtPoint(p2);

                                    if (sta1 > sta2)
                                    {
                                        Point3d ptemp = new Point3d(p1.X, p1.Y, 0);
                                        p1 = new Point3d(p2.X, p2.Y, 0);
                                        p2 = new Point3d(ptemp.X, ptemp.Y, 0);
                                    }

                                    Polyline poly2 = new Polyline();
                                    poly2.AddVertexAt(0, new Point2d(p1.X, p1.Y), 0, 0, 0);
                                    poly2.AddVertexAt(1, new Point2d(p2.X, p2.Y), 0, 0, 0);

                                    DBObjectCollection col_off = poly2.GetOffsetCurves(-20);
                                    Polyline poly3 = new Polyline();
                                    poly3 = col_off[0] as Polyline;
                                    if (poly3 != null)
                                    {


                                        Point3d pt2 = poly3.StartPoint;
                                        Point3d pt3 = poly3.EndPoint;
                                        Polyline poly4 = new Polyline();
                                        poly4.AddVertexAt(0, new Point2d(pt2.X, pt2.Y), 0, 0, 0);
                                        poly4.AddVertexAt(1, new Point2d(pt3.X, pt3.Y), 0, 0, 0);
                                        poly4.TransformBy(Matrix3d.Displacement(pt3.GetVectorTo(pt2)));
                                        poly4.TransformBy(Matrix3d.Rotation(30 * Math.PI / 180, Vector3d.ZAxis, pt2));
                                        poly4.TransformBy(Matrix3d.Scaling(10, pt2));
                                        Point3dCollection col_int2 = Functions.Intersect_on_both_operands(poly1, poly4);

                                        Polyline poly5 = new Polyline();
                                        poly5.AddVertexAt(0, new Point2d(pt2.X, pt2.Y), 0, 0, 0);
                                        poly5.AddVertexAt(1, new Point2d(pt3.X, pt3.Y), 0, 0, 0);
                                        poly5.TransformBy(Matrix3d.Displacement(pt2.GetVectorTo(pt3)));
                                        poly5.TransformBy(Matrix3d.Rotation(330 * Math.PI / 180, Vector3d.ZAxis, pt3));
                                        poly5.TransformBy(Matrix3d.Scaling(10, pt3));
                                        Point3dCollection col_int3 = Functions.Intersect_on_both_operands(poly1, poly5);


                                        if (col_int2.Count == 1 && col_int3.Count == 1)
                                        {
                                            Point3d pt1 = col_int2[0];
                                            Point3d pt4 = col_int3[0];


                                            Polyline poly6 = new Polyline();
                                            poly6.AddVertexAt(0, new Point2d(pt1.X, pt1.Y), 0, poly1.GetStartWidthAt(0), poly1.GetEndWidthAt(0));
                                            poly6.AddVertexAt(1, new Point2d(pt2.X, pt2.Y), 0, poly1.GetStartWidthAt(0), poly1.GetEndWidthAt(0));
                                            poly6.AddVertexAt(2, new Point2d(pt3.X, pt3.Y), 0, poly1.GetStartWidthAt(0), poly1.GetEndWidthAt(0));
                                            poly6.AddVertexAt(3, new Point2d(pt4.X, pt4.Y), 0, poly1.GetStartWidthAt(0), poly1.GetEndWidthAt(0));
                                            poly6.Layer = poly1.Layer;
                                            poly6.ColorIndex = 256;
                                            BTrecord.AppendEntity(poly6);
                                            Trans1.AddNewlyCreatedDBObject(poly6, true);

                                        }
                                        Trans1.Commit();
                                    }

                                }
                            }




                        }
                    } while (run1 == true);


                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");


        }

        [CommandMethod("pp1")]
        public void pick_pt()
        {

            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {

                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    System.Data.DataTable dt1 = new System.Data.DataTable();
                    dt1.Columns.Add("Point Name", typeof(string));
                    dt1.Columns.Add("X", typeof(double));
                    dt1.Columns.Add("Y", typeof(double));
                    dt1.Columns.Add("Z", typeof(double));

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

                    object old_osnap = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("OSMODE");
                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", 512);
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        bool run1 = true;
                        do
                        {
                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                            Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                            PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify  point");
                            PP1.AllowNone = false;
                            Point_res1 = Editor1.GetPoint(PP1);

                            if (Point_res1.Status != PromptStatus.OK)
                            {
                                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", old_osnap);
                                Functions.Transfer_datatable_to_new_excel_spreadsheet(dt1);
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }

                            Autodesk.AutoCAD.EditorInput.PromptStringOptions Prompt_string = new Autodesk.AutoCAD.EditorInput.PromptStringOptions("\n" + "Specify point name:");
                            Autodesk.AutoCAD.EditorInput.PromptResult Rezultat_string = ThisDrawing.Editor.GetString(Prompt_string);

                            string pn = "xx";

                            if (Rezultat_string.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                            {
                                pn = Rezultat_string.StringResult;
                            }

                            dt1.Rows.Add();
                            dt1.Rows[dt1.Rows.Count - 1][1] = Point_res1.Value.X;
                            dt1.Rows[dt1.Rows.Count - 1][2] = Point_res1.Value.Y;
                            dt1.Rows[dt1.Rows.Count - 1][3] = Point_res1.Value.Z;
                            dt1.Rows[dt1.Rows.Count - 1][0] = pn;

                        } while (run1 == true);
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

        [CommandMethod("A1")]
        public void CALC_AREA_FEET_ACRES_2_DECS()
        {

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

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rez1;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_1;
                        Prompt_1 = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect a curve:");
                        Prompt_1.SetRejectMessage("\nSelect a curve!");
                        Prompt_1.AllowNone = true;
                        Prompt_1.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Curve), false);
                        Rez1 = ThisDrawing.Editor.GetEntity(Prompt_1);

                        if (Rez1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Curve poly1 = Trans1.GetObject(Rez1.ObjectId, OpenMode.ForWrite) as Curve;
                        double area1 = Math.Round(poly1.Area, 2);
                        double area2 = Math.Round(area1 / 43560, 2);
                        string acres2 = Functions.Get_String_Rounded(area2, 2);


                        poly1.LineWeight = LineWeight.LineWeight070;
                        poly1.ColorIndex = 5;

                        Editor1.WriteMessage("\n" + area1.ToString() + " sqft");
                        Editor1.WriteMessage("\n" + acres2 + " ac.");

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
        [CommandMethod("A2")]
        public void CALC_AREA_FEET_ACRES_2_DECS_multiple()
        {

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


                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez.MessageForAdding = "\nSelect the objects:";
                        Prompt_rez.SingleOnly = false;
                        Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                        if (Rezultat1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        double area1 = 0;
                        double area2 = 0;
                        double area3 = 0;


                        for (int i = 0; i < Rezultat1.Value.Count; i++)
                        {
                            Curve poly1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForWrite) as Curve;
                            if (poly1 != null)
                            {
                                area1 = area1 + poly1.Area;
                                area3 = area3 + Math.Round(Math.Round(poly1.Area, 2) / 43560, 2);
                                poly1.LineWeight = LineWeight.LineWeight025;
                                poly1.ColorIndex = 8;
                            }
                        }

                        area2 = Math.Round(area1 / 43560, 2);


                        string feet1 = Functions.Get_String_Rounded(area1, 2);
                        string acres2 = Functions.Get_String_Rounded(area2, 2);
                        string acres3 = Functions.Get_String_Rounded(area3, 2);

                        Editor1.WriteMessage("\n" + feet1 + " sqft");
                        Editor1.WriteMessage("\n" + acres2 + " ac." + " **[feet not rounded and divided by 43560]");
                        Editor1.WriteMessage("\n" + acres3 + " ac." + " **[feet rounded and then divided by 43560]");

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

        [CommandMethod("L1")]
        public void CALC_length_FEET_rod_2_DECS()
        {

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

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rez1;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_1;
                        Prompt_1 = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect a poly:");
                        Prompt_1.SetRejectMessage("\nSelect a polyline!");
                        Prompt_1.AllowNone = true;
                        Prompt_1.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rez1 = ThisDrawing.Editor.GetEntity(Prompt_1);

                        if (Rez1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Polyline poly1 = Trans1.GetObject(Rez1.ObjectId, OpenMode.ForWrite) as Polyline;
                        double L1 = Math.Round(poly1.Length, 2);
                        double L2 = Math.Round(L1 / 16.5, 2);
                        string ROD2 = Functions.Get_String_Rounded(L2, 2);


                        poly1.LineWeight = LineWeight.LineWeight070;
                        poly1.ColorIndex = 5;

                        Editor1.WriteMessage("\n" + L1.ToString() + " FT");
                        Editor1.WriteMessage("\n" + ROD2 + " rods");

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

        [CommandMethod("L2")]
        public void CALC_LEN_FEET_RODS_2_DECS_multiple()
        {

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


                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez.MessageForAdding = "\nSelect the objects:";
                        Prompt_rez.SingleOnly = false;
                        Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                        if (Rezultat1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        double L1 = 0;

                        for (int i = 0; i < Rezultat1.Value.Count; i++)
                        {
                            Polyline poly1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForWrite) as Polyline;
                            if (poly1 != null)
                            {
                                L1 = L1 + Math.Round(poly1.Length, 2);
                                poly1.LineWeight = LineWeight.LineWeight025;
                                poly1.ColorIndex = 8;
                            }
                        }





                        double L2 = Math.Round(L1 / 16.5, 2);
                        string rod2 = Functions.Get_String_Rounded(L2, 2);




                        Editor1.WriteMessage("\n" + L1.ToString() + " FT");
                        Editor1.WriteMessage("\n" + rod2 + " RODS");

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
        [CommandMethod("block2block")]
        public void copy_block_attrib()
        {

            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {

                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    bool run1 = true;

                    do
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                            Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_source;
                            Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_source;
                            Prompt_source = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSource Object [Block or Mtext or Text or Mleader]:");
                            Prompt_source.SetRejectMessage("\nSelect a blockreference!");
                            Prompt_source.AllowNone = true;
                            Prompt_source.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.BlockReference), false);
                            Prompt_source.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.MText), false);
                            Prompt_source.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.MLeader), false);
                            Prompt_source.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.DBText), false);
                            Rezultat_source = ThisDrawing.Editor.GetEntity(Prompt_source);

                            if (Rezultat_source.Status != PromptStatus.OK)
                            {
                                run1 = false;
                            }
                            Autodesk.AutoCAD.EditorInput.PromptSelectionResult rez_destination = null;
                            if (run1 == true)
                            {




                                Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                                Prompt_rez.MessageForAdding = "\nDestination object:";
                                Prompt_rez.SingleOnly = false;
                                rez_destination = ThisDrawing.Editor.GetSelection(Prompt_rez);

                                if (rez_destination.Status != PromptStatus.OK)
                                {

                                    run1 = false;
                                }
                            }
                            if (run1 == true)
                            {
                                BlockReference bl1 = Trans1.GetObject(Rezultat_source.ObjectId, OpenMode.ForRead) as BlockReference;

                                MText mtext1 = Trans1.GetObject(Rezultat_source.ObjectId, OpenMode.ForRead) as MText;
                                DBText text1 = Trans1.GetObject(Rezultat_source.ObjectId, OpenMode.ForRead) as DBText;
                                MLeader mleader1 = Trans1.GetObject(Rezultat_source.ObjectId, OpenMode.ForRead) as MLeader;

                                for (int i = 0; i < rez_destination.Value.Count; i++)
                                {
                                    BlockReference bl2 = Trans1.GetObject(rez_destination.Value[i].ObjectId, OpenMode.ForWrite) as BlockReference;
                                    MText mtext2 = Trans1.GetObject(rez_destination.Value[i].ObjectId, OpenMode.ForWrite) as MText;
                                    DBText text2 = Trans1.GetObject(rez_destination.Value[i].ObjectId, OpenMode.ForWrite) as DBText;
                                    MLeader mleader2 = Trans1.GetObject(rez_destination.Value[i].ObjectId, OpenMode.ForWrite) as MLeader;

                                    if (bl1 != null && bl2 != null)
                                    {
                                        if (bl1.AttributeCollection.Count > 0 && bl2.AttributeCollection.Count > 0)
                                        {
                                            System.Data.DataTable dt1 = Functions.Read_block_attributes_and_values(bl1);

                                            Autodesk.AutoCAD.DatabaseServices.AttributeCollection attColl2 = bl2.AttributeCollection;
                                            foreach (ObjectId id2 in attColl2)
                                            {
                                                AttributeReference atr2 = Trans1.GetObject(id2, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite) as AttributeReference;
                                                if (atr2 != null)
                                                {
                                                    string tag2 = atr2.Tag;

                                                    for (int k = 0; k < dt1.Rows.Count; k++)
                                                    {
                                                        string val1 = "";
                                                        if (dt1.Rows[k]["VALUE"] != DBNull.Value) val1 = Convert.ToString(dt1.Rows[k]["VALUE"]);

                                                        string tag1 = Convert.ToString(dt1.Rows[k]["ATTRIB"]);

                                                        if (tag1.ToUpper() == tag2.ToUpper())
                                                        {
                                                            atr2.TextString = val1;
                                                        }
                                                    }


                                                }
                                            }



                                        }

                                    }

                                    else if (mtext1 != null && mtext2 != null)
                                    {
                                        mtext2.Contents = mtext1.Contents;
                                    }
                                    else if (text1 != null && mtext2 != null)
                                    {
                                        mtext2.Contents = text1.TextString;
                                    }
                                    else if (mleader1 != null && mtext2 != null)
                                    {
                                        mtext2.Contents = mleader1.MText.Contents;
                                    }

                                    else if (mtext1 != null && text2 != null)
                                    {
                                        text2.TextString = mtext1.Text;
                                    }
                                    else if (text1 != null && text2 != null)
                                    {
                                        text2.TextString = text1.TextString;
                                    }
                                    else if (mleader1 != null && text2 != null)
                                    {
                                        text2.TextString = mleader1.MText.Text;
                                    }

                                    else if (mtext1 != null && mleader2 != null)
                                    {
                                        MText mtext3 = mleader2.MText;
                                        mtext3.Contents = mtext1.Contents;
                                        mleader2.MText = mtext3;
                                    }
                                    else if (text1 != null && mleader2 != null)
                                    {
                                        MText mtext3 = mleader2.MText;
                                        mtext3.Contents = text1.TextString;
                                        mleader2.MText = mtext3;
                                    }
                                    else if (mleader1 != null && mleader2 != null)
                                    {

                                        MText mtext3 = mleader2.MText;
                                        mtext3.Contents = mleader1.MText.Contents;
                                        mleader2.MText = mtext3;
                                    }


                                }

                            }

                            Trans1.Commit();

                        }
                    } while (run1 == true);

                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
        }
        [CommandMethod("poly2elevation")]
        public void poly2elevation()
        {

            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

            try
            {

                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    bool run1 = true;

                    do
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            Editor1.SetImpliedSelection(Empty_array);

                            BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                            Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_source;
                            Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_source;
                            Prompt_source = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nText Object:");
                            Prompt_source.SetRejectMessage("\nSelect a text!");
                            Prompt_source.AllowNone = true;
                            Prompt_source.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.DBText), false);
                            Rezultat_source = ThisDrawing.Editor.GetEntity(Prompt_source);

                            if (Rezultat_source.Status != PromptStatus.OK)
                            {
                                run1 = false;
                            }
                            Autodesk.AutoCAD.EditorInput.PromptSelectionResult rez_destination = null;
                            if (run1 == true)
                            {




                                Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                                Prompt_rez.MessageForAdding = "\nSpecify polyline:";
                                Prompt_rez.SingleOnly = false;
                                rez_destination = ThisDrawing.Editor.GetSelection(Prompt_rez);

                                if (rez_destination.Status != PromptStatus.OK)
                                {

                                    run1 = false;
                                }
                            }
                            if (run1 == true)
                            {


                                DBText text1 = Trans1.GetObject(Rezultat_source.ObjectId, OpenMode.ForRead) as DBText;

                                for (int i = 0; i < rez_destination.Value.Count; i++)
                                {
                                    Polyline poly2 = Trans1.GetObject(rez_destination.Value[i].ObjectId, OpenMode.ForWrite) as Polyline;


                                    if (text1 != null && poly2 != null)
                                    {
                                        if (Functions.IsNumeric(text1.TextString) == true)
                                        {
                                            poly2.Elevation = Convert.ToDouble(text1.TextString);
                                        }

                                    }




                                }

                            }

                            Trans1.Commit();

                        }
                    } while (run1 == true);

                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
        }

        [CommandMethod("w2xl", CommandFlags.UsePickSet | CommandFlags.Redraw)]
        public void write_poly_to_excel()
        {


            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {
                System.Data.DataTable dt1 = new System.Data.DataTable();

                string col_type = "Type of Object";
                string col_layer = "Layer";
                string col_sta = "Station";
                string col_x = "x";
                string col_y = "y";
                string col_z = "z";
                string col_defl = "Deflection DD";
                string col_defl_dms = "Deflection DMS";

                string col_angle = "Rotation DD";
                string col_angle_dms = "Rotation DMS";
                string col_bulge = "Bulge Angle [Rad]";
                string col_radius = "Arc Radius";
                string col_midX = "Arc Middle X";
                string col_midY = "Arc Middle Y";
                string col_arc_len = "Arc Length";
                string col_length = "Length";
                string col_textstring = "Text String";
                string col_textheight = "Text Height";
                string col_rotationtxt = "Text Rotation";
                string col_Blockname = "Block Name";
                string col_rotationblock = "Block Rotation";
                string col_blockscale = "Block Scale";
                string col_area = "Area";
                string col_sta2d = "Station2D";

                dt1.Columns.Add(col_type, typeof(string));
                dt1.Columns.Add(col_layer, typeof(string));
                dt1.Columns.Add(col_sta, typeof(double));
                dt1.Columns.Add(col_x, typeof(double));
                dt1.Columns.Add(col_y, typeof(double));
                dt1.Columns.Add(col_z, typeof(double));
                dt1.Columns.Add(col_defl, typeof(double));
                dt1.Columns.Add(col_defl_dms, typeof(string));
                dt1.Columns.Add(col_angle, typeof(double));
                dt1.Columns.Add(col_angle_dms, typeof(string));
                dt1.Columns.Add(col_bulge, typeof(double));
                dt1.Columns.Add(col_radius, typeof(double));
                dt1.Columns.Add(col_arc_len, typeof(double));
                dt1.Columns.Add(col_midX, typeof(double));
                dt1.Columns.Add(col_midY, typeof(double));
                dt1.Columns.Add(col_length, typeof(double));
                dt1.Columns.Add(col_area, typeof(double));
                dt1.Columns.Add(col_textstring, typeof(string));
                dt1.Columns.Add(col_textheight, typeof(double));
                dt1.Columns.Add(col_rotationtxt, typeof(double));
                dt1.Columns.Add(col_Blockname, typeof(string));
                dt1.Columns.Add(col_rotationblock, typeof(double));
                dt1.Columns.Add(col_blockscale, typeof(double));
                dt1.Columns.Add(col_sta2d, typeof(double));


                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;


                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1 = Editor1.SelectImplied();
                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rez.MessageForAdding = "\nSelect objects:";
                            Prompt_rez.SingleOnly = false;
                            Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                            if (Rezultat1.Status != PromptStatus.OK)
                            {
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }
                        }
                        if (Rezultat1.Status == PromptStatus.OK)
                        {

                            int index_pt = 1;
                            int index_circle = 1;
                            int index_line = 1;
                            int index_polyline = 1;
                            int index_3dpolyline = 1;
                            int index_mtext = 1;
                            int index_text = 1;
                            int index_blk = 1;
                            int index_mleader = 1;

                            for (int i = 0; i < Rezultat1.Value.Count; i++)
                            {
                                Entity ent1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as Entity;
                                if (ent1 != null)
                                {
                                    if (ent1 is DBPoint)
                                    {
                                        DBPoint pt1 = ent1 as DBPoint;
                                        dt1.Rows.Add();
                                        dt1.Rows[dt1.Rows.Count - 1][col_type] = "DBPoint#" + Convert.ToString(index_pt);
                                        dt1.Rows[dt1.Rows.Count - 1][col_x] = pt1.Position.X;
                                        dt1.Rows[dt1.Rows.Count - 1][col_y] = pt1.Position.Y;
                                        dt1.Rows[dt1.Rows.Count - 1][col_z] = pt1.Position.Z;
                                        dt1.Rows[dt1.Rows.Count - 1][col_layer] = pt1.Layer;
                                        ++index_pt;
                                    }
                                    else if (ent1 is MText)
                                    {
                                        MText mtext1 = ent1 as MText;
                                        dt1.Rows.Add();
                                        dt1.Rows[dt1.Rows.Count - 1][col_type] = "Mtext#" + Convert.ToString(index_mtext);
                                        dt1.Rows[dt1.Rows.Count - 1][col_x] = mtext1.Location.X;
                                        dt1.Rows[dt1.Rows.Count - 1][col_y] = mtext1.Location.Y;
                                        dt1.Rows[dt1.Rows.Count - 1][col_z] = mtext1.Location.Z;
                                        dt1.Rows[dt1.Rows.Count - 1][col_layer] = mtext1.Layer;
                                        dt1.Rows[dt1.Rows.Count - 1][col_textstring] = mtext1.Contents;
                                        dt1.Rows[dt1.Rows.Count - 1][col_textheight] = mtext1.TextHeight;
                                        dt1.Rows[dt1.Rows.Count - 1][col_rotationtxt] = 180 * mtext1.Rotation / Math.PI;
                                        ++index_mtext;
                                    }
                                    else if (ent1 is DBText)
                                    {
                                        DBText text1 = ent1 as DBText;
                                        dt1.Rows.Add();
                                        dt1.Rows[dt1.Rows.Count - 1][col_type] = "text#" + Convert.ToString(index_text);
                                        dt1.Rows[dt1.Rows.Count - 1][col_x] = text1.Position.X;
                                        dt1.Rows[dt1.Rows.Count - 1][col_y] = text1.Position.Y;
                                        dt1.Rows[dt1.Rows.Count - 1][col_z] = text1.Position.Z;
                                        dt1.Rows[dt1.Rows.Count - 1][col_layer] = text1.Layer;
                                        dt1.Rows[dt1.Rows.Count - 1][col_textstring] = text1.TextString;
                                        dt1.Rows[dt1.Rows.Count - 1][col_textheight] = text1.Height;
                                        dt1.Rows[dt1.Rows.Count - 1][col_rotationtxt] = 180 * text1.Rotation / Math.PI;
                                        ++index_text;
                                    }
                                    else if (ent1 is BlockReference)
                                    {
                                        BlockReference br1 = ent1 as BlockReference;
                                        dt1.Rows.Add();
                                        dt1.Rows[dt1.Rows.Count - 1][col_type] = "BlockReference#" + Convert.ToString(index_blk);
                                        dt1.Rows[dt1.Rows.Count - 1][col_x] = br1.Position.X;
                                        dt1.Rows[dt1.Rows.Count - 1][col_y] = br1.Position.Y;
                                        dt1.Rows[dt1.Rows.Count - 1][col_z] = br1.Position.Z;
                                        dt1.Rows[dt1.Rows.Count - 1][col_layer] = br1.Layer;
                                        string bn = Functions.get_block_name(br1);
                                        dt1.Rows[dt1.Rows.Count - 1][col_Blockname] = bn;
                                        dt1.Rows[dt1.Rows.Count - 1][col_blockscale] = br1.ScaleFactors.X;
                                        dt1.Rows[dt1.Rows.Count - 1][col_rotationblock] = 180 * br1.Rotation / Math.PI;

                                        if (br1.AttributeCollection.Count > 0)
                                        {
                                            for (int k = 0; k < br1.AttributeCollection.Count; k++)
                                            {
                                                ObjectId id1 = br1.AttributeCollection[k];
                                                AttributeReference atr1 = Trans1.GetObject(id1, OpenMode.ForRead) as AttributeReference;
                                                if (atr1 != null)
                                                {
                                                    string atr_name = atr1.Tag;
                                                    string val1 = atr1.TextString;
                                                    if (atr1.IsMTextAttribute == true)
                                                    {
                                                        val1 = atr1.MTextAttribute.Contents;
                                                    }

                                                    string colname = bn + "_" + atr_name;
                                                    if (dt1.Columns.Contains(colname) == false)
                                                    {
                                                        dt1.Columns.Add(colname, typeof(string));
                                                    }
                                                    dt1.Rows[dt1.Rows.Count - 1][colname] = val1;
                                                }
                                            }
                                        }
                                        ++index_blk;
                                    }
                                    else if (ent1 is Circle)
                                    {
                                        Circle circ1 = ent1 as Circle;
                                        dt1.Rows.Add();
                                        dt1.Rows[dt1.Rows.Count - 1][col_type] = "Circle#" + Convert.ToString(index_circle);
                                        dt1.Rows[dt1.Rows.Count - 1][col_x] = circ1.Center.X;
                                        dt1.Rows[dt1.Rows.Count - 1][col_y] = circ1.Center.Y;
                                        dt1.Rows[dt1.Rows.Count - 1][col_z] = circ1.Center.Z;
                                        dt1.Rows[dt1.Rows.Count - 1][col_layer] = circ1.Layer;
                                        ++index_circle;
                                    }
                                    else if (ent1 is Autodesk.AutoCAD.DatabaseServices.Line)
                                    {
                                        Autodesk.AutoCAD.DatabaseServices.Line line1 = ent1 as Autodesk.AutoCAD.DatabaseServices.Line;

                                        double x1 = line1.StartPoint.X;
                                        double y1 = line1.StartPoint.Y;
                                        double x2 = line1.EndPoint.X;
                                        double y2 = line1.EndPoint.Y;

                                        dt1.Rows.Add();
                                        dt1.Rows[dt1.Rows.Count - 1][col_type] = "Line#" + Convert.ToString(index_line);
                                        dt1.Rows[dt1.Rows.Count - 1][col_x] = x1;
                                        dt1.Rows[dt1.Rows.Count - 1][col_y] = y1;
                                        dt1.Rows[dt1.Rows.Count - 1][col_z] = line1.StartPoint.Z;
                                        dt1.Rows[dt1.Rows.Count - 1][col_layer] = line1.Layer;
                                        dt1.Rows.Add();
                                        dt1.Rows[dt1.Rows.Count - 1][col_type] = "Line#" + Convert.ToString(index_line);
                                        dt1.Rows[dt1.Rows.Count - 1][col_x] = x2;
                                        dt1.Rows[dt1.Rows.Count - 1][col_y] = y2;
                                        dt1.Rows[dt1.Rows.Count - 1][col_z] = line1.EndPoint.Z;
                                        dt1.Rows[dt1.Rows.Count - 1][col_layer] = line1.Layer;
                                        dt1.Rows[dt1.Rows.Count - 1][col_length] = line1.Length;


                                        double rot1 = Functions.GET_Bearing_rad(x1, y1, x2, y2);
                                        string dms = Functions.Get_DMS(rot1 * 180 / Math.PI, 0);
                                        dt1.Rows[dt1.Rows.Count - 1][col_angle] = rot1 * 180 / Math.PI;
                                        dt1.Rows[dt1.Rows.Count - 1][col_angle_dms] = dms;


                                        ++index_line;
                                    }
                                    else if (ent1 is Autodesk.AutoCAD.DatabaseServices.Polyline)
                                    {
                                        Autodesk.AutoCAD.DatabaseServices.Polyline poly1 = ent1 as Autodesk.AutoCAD.DatabaseServices.Polyline;


                                        for (int j = 0; j < poly1.NumberOfVertices; j++)
                                        {

                                            double x1 = poly1.GetPoint2dAt(j).X;
                                            double y1 = poly1.GetPoint2dAt(j).Y;
                                            double bulge1 = poly1.GetBulgeAt(j);



                                            dt1.Rows.Add();
                                            dt1.Rows[dt1.Rows.Count - 1][col_type] = "Polyline#" + Convert.ToString(index_polyline);
                                            dt1.Rows[dt1.Rows.Count - 1][col_x] = x1;
                                            dt1.Rows[dt1.Rows.Count - 1][col_y] = y1;
                                            dt1.Rows[dt1.Rows.Count - 1][col_z] = poly1.Elevation;
                                            dt1.Rows[dt1.Rows.Count - 1][col_sta] = poly1.GetDistanceAtParameter(j);
                                            dt1.Rows[dt1.Rows.Count - 1][col_bulge] = bulge1;
                                            dt1.Rows[dt1.Rows.Count - 1][col_layer] = poly1.Layer;

                                            if (bulge1 != 0 && j < poly1.NumberOfVertices - 1)
                                            {
                                                CircularArc2d arc1 = poly1.GetArcSegment2dAt(j);

                                                double radius1 = arc1.Radius;
                                                if (j < poly1.NumberOfVertices - 1)
                                                {
                                                    Polyline poly2 = new Polyline();
                                                    poly2.AddVertexAt(0, poly1.GetPoint2dAt(j), bulge1, 0, 0);
                                                    poly2.AddVertexAt(1, poly1.GetPoint2dAt(j + 1), 0, 0, 0);
                                                    Point3d ptmid = poly2.GetPointAtDist(poly2.Length / 2);
                                                    dt1.Rows[dt1.Rows.Count - 1][col_midX] = ptmid.X;
                                                    dt1.Rows[dt1.Rows.Count - 1][col_midY] = ptmid.Y;
                                                }

                                                double len1 = Math.Abs(radius1 * 4 * Math.Atan(bulge1));
                                                dt1.Rows[dt1.Rows.Count - 1][col_arc_len] = len1;
                                                dt1.Rows[dt1.Rows.Count - 1][col_radius] = radius1;

                                            }


                                            if (j == 0)
                                            {
                                                dt1.Rows[dt1.Rows.Count - 1][col_length] = poly1.Length;
                                                dt1.Rows[dt1.Rows.Count - 1][col_area] = poly1.Area;

                                            }

                                            if (j > 0 && j < poly1.NumberOfVertices - 1)
                                            {
                                                double x0 = poly1.GetPoint2dAt(j - 1).X;
                                                double y0 = poly1.GetPoint2dAt(j - 1).Y;
                                                double x2 = poly1.GetPoint2dAt(j + 1).X;
                                                double y2 = poly1.GetPoint2dAt(j + 1).Y;


                                                double rot1 = Functions.Get_deflection_angle_rad(x0, y0, x1, y1, x2, y2);
                                                string dms = Functions.Get_deflection_angle_dms(x0, y0, x1, y1, x2, y2);
                                                dt1.Rows[dt1.Rows.Count - 1][col_defl] = rot1 * 180 / Math.PI;
                                                dt1.Rows[dt1.Rows.Count - 1][col_defl_dms] = dms;
                                            }

                                        }



                                        ++index_polyline;
                                    }
                                    else if (ent1 is Autodesk.AutoCAD.DatabaseServices.Polyline3d)
                                    {
                                        Autodesk.AutoCAD.DatabaseServices.Polyline3d poly3D = ent1 as Autodesk.AutoCAD.DatabaseServices.Polyline3d;
                                        Polyline poly1 = Functions.Build_2dpoly_from_3d(poly3D);
                                        dt1.Columns[col_sta2d].SetOrdinal(5);
                                        for (int j = 0; j < poly1.NumberOfVertices; j++)
                                        {
                                            double x1 = poly1.GetPoint2dAt(j).X;
                                            double y1 = poly1.GetPoint2dAt(j).Y;

                                            dt1.Rows.Add();
                                            dt1.Rows[dt1.Rows.Count - 1][col_type] = "Polyline3D#" + Convert.ToString(index_3dpolyline);
                                            dt1.Rows[dt1.Rows.Count - 1][col_x] = x1;
                                            dt1.Rows[dt1.Rows.Count - 1][col_y] = y1;
                                            dt1.Rows[dt1.Rows.Count - 1][col_z] = poly3D.GetPointAtParameter(j).Z;
                                            dt1.Rows[dt1.Rows.Count - 1][col_sta] = poly3D.GetDistanceAtParameter(j);
                                            dt1.Rows[dt1.Rows.Count - 1][col_sta2d] = poly1.GetDistanceAtParameter(j);
                                            dt1.Rows[dt1.Rows.Count - 1][col_layer] = poly3D.Layer;
                                            if (j == 0)
                                            {
                                                dt1.Rows[dt1.Rows.Count - 1][col_length] = poly3D.Length;

                                            }


                                            if (j > 0 && j < poly1.NumberOfVertices - 1)
                                            {
                                                double x0 = poly1.GetPoint2dAt(j - 1).X;
                                                double y0 = poly1.GetPoint2dAt(j - 1).Y;
                                                double x2 = poly1.GetPoint2dAt(j + 1).X;
                                                double y2 = poly1.GetPoint2dAt(j + 1).Y;

                                                double rot1 = Functions.Get_deflection_angle_rad(x0, y0, x1, y1, x2, y2);
                                                string dms = Functions.Get_deflection_angle_dms(x0, y0, x1, y1, x2, y2);
                                                dt1.Rows[dt1.Rows.Count - 1][col_defl] = rot1 * 180 / Math.PI;
                                                dt1.Rows[dt1.Rows.Count - 1][col_defl_dms] = dms;
                                            }
                                        }

                                        ++index_3dpolyline;
                                    }
                                    else if (ent1 is MLeader)
                                    {
                                        MLeader ml1 = ent1 as MLeader;
                                        if (ml1 != null)
                                        {
                                            Point3d ptins = ml1.GetFirstVertex(0);

                                            dt1.Rows.Add();
                                            dt1.Rows[dt1.Rows.Count - 1][col_type] = "Mleader#" + Convert.ToString(index_mleader);
                                            dt1.Rows[dt1.Rows.Count - 1][col_x] = ptins.X;
                                            dt1.Rows[dt1.Rows.Count - 1][col_y] = ptins.Y;

                                            dt1.Rows[dt1.Rows.Count - 1][col_layer] = ml1.Layer;
                                            dt1.Rows[dt1.Rows.Count - 1][col_textstring] = ml1.MText.Contents;

                                            ++index_mleader;



                                        }
                                    }
                                }
                            }

                            Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt1);


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


        }





        [CommandMethod("R180", CommandFlags.UsePickSet | CommandFlags.Redraw)]
        public void rotate_mleader_mtext_dbtext_blockref_with_180()
        {


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
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        Matrix3d CurentUCSmatrix = Editor1.CurrentUserCoordinateSystem;




                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1 = Editor1.SelectImplied();


                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rez.MessageForAdding = "\nSelect text:";
                            Prompt_rez.SingleOnly = false;
                            Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);
                        }


                        if (Rezultat1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }




                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {
                            MLeader mleader1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForWrite) as MLeader;
                            if (mleader1 != null)
                            {
                                System.Collections.ArrayList lista_indexes = mleader1.GetLeaderIndexes();

                                Point3d pt0 = mleader1.GetFirstVertex(0);
                                //Vector3d dogleg0 = mleader1.GetDogleg(0);

                                double rot5 = mleader1.MText.Rotation;

                                double rot1 = Functions.GET_Bearing_rad(curent_ucs_matrix.CoordinateSystem3d.Xaxis);
                                double rot_WCS = rot1 + rot5;

                                // double rot1_deg = rot1 * 180 / Math.PI - 180;
                                double rot5_deg = rot5 * 180 / Math.PI;
                                double rot1_deg = rot1 * 180 / Math.PI;
                                double rowcs_deg = rot_WCS * 180 / Math.PI;

                                mleader1.TransformBy(Matrix3d.Rotation(rot_WCS + Math.PI, Vector3d.ZAxis, pt0));
                            }

                            BlockReference block1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForWrite) as BlockReference;

                            if (block1 != null)
                            {
                                string blockname = Functions.get_block_name(block1);
                                BlockTableRecord btr = Trans1.GetObject(BlockTable1[blockname], OpenMode.ForRead) as BlockTableRecord;
                                if (btr != null && btr.IsFromExternalReference == false && btr.IsLayout == false
                                    && btr.IsFromOverlayReference == false && btr.IsAnonymous == false )
                                {
                                    double r1 = block1.Rotation;
                                    double rot1_deg = r1 * 180 / Math.PI;
                                    block1.Rotation = r1 + Math.PI;
                                }
                            }

                            MText mt1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForWrite) as MText;
                            if (mt1 != null)
                            {
                                double r1 = mt1.Rotation;
                                double rot1_deg = r1 * 180 / Math.PI;

                                mt1.Rotation = r1 + Math.PI;
                            }


                            DBText txt1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForWrite) as DBText;
                            if (txt1 != null)
                            {
                                double r1 = txt1.Rotation;
                                double rot1_deg = r1 * 180 / Math.PI;

                                txt1.Rotation = r1 + Math.PI;
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




        }

        [CommandMethod("f300", CommandFlags.UsePickSet | CommandFlags.Redraw)]
        public void fillet_with_radius_300()
        {
            fillet_with_radius(300);
        }
        [CommandMethod("f200", CommandFlags.UsePickSet | CommandFlags.Redraw)]
        public void fillet_with_radius_200()
        {
            fillet_with_radius(200);
        }
        [CommandMethod("f100", CommandFlags.UsePickSet | CommandFlags.Redraw)]
        public void fillet_with_radius_100()
        {
            fillet_with_radius(100);
        }

        [CommandMethod("f50", CommandFlags.UsePickSet | CommandFlags.Redraw)]
        public void fillet_with_radius_50()
        {
            fillet_with_radius(50);
        }

        [CommandMethod("f75", CommandFlags.UsePickSet | CommandFlags.Redraw)]
        public void fillet_with_radius_75()
        {
            fillet_with_radius(75);
        }

        [CommandMethod("f125", CommandFlags.UsePickSet | CommandFlags.Redraw)]
        public void fillet_with_radius_125()
        {
            fillet_with_radius(125);
        }

        [CommandMethod("f150", CommandFlags.UsePickSet | CommandFlags.Redraw)]
        public void fillet_with_radius_150()
        {
            fillet_with_radius(150);
        }

        [CommandMethod("f175", CommandFlags.UsePickSet | CommandFlags.Redraw)]
        public void fillet_with_radius_175()
        {
            fillet_with_radius(175);
        }

        [CommandMethod("f225", CommandFlags.UsePickSet | CommandFlags.Redraw)]
        public void fillet_with_radius_225()
        {
            fillet_with_radius(225);
        }

        [CommandMethod("f250", CommandFlags.UsePickSet | CommandFlags.Redraw)]
        public void fillet_with_radius_250()
        {
            fillet_with_radius(250);
        }

        [CommandMethod("f275", CommandFlags.UsePickSet | CommandFlags.Redraw)]
        public void fillet_with_radius_275()
        {
            fillet_with_radius(275);
        }



        [CommandMethod("f0", CommandFlags.UsePickSet | CommandFlags.Redraw)]
        public void remove_fillet()
        {


            ObjectId[] Empty_array = null;
            ObjectId id_poly = ObjectId.Null;
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
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                        Polyline poly1 = null;
                        Point3d pt_picked = new Point3d();
                        bool was_selected = false;

                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1 = Editor1.SelectImplied();
                        if (Rezultat1.Status == PromptStatus.OK)
                        {
                            if (Rezultat1.Value.Count == 1)
                            {
                                ObjectId id1 = Rezultat1.Value[0].ObjectId;

                                poly1 = Trans1.GetObject(id1, OpenMode.ForWrite) as Polyline;

                                if (poly1 != null)
                                {
                                    Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                    Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                    PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSelect the node:");
                                    PP1.AllowNone = false;
                                    Point_res1 = Editor1.GetPoint(PP1);
                                    if (Point_res1.Status != PromptStatus.OK)
                                    {
                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                        return;
                                    }
                                    pt_picked = poly1.GetClosestPointTo(Point_res1.Value, Vector3d.ZAxis, false);
                                    was_selected = true;
                                }
                            }
                        }

                        if (was_selected == false)
                        {
                            Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                            Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                            Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the node:");
                            Prompt_centerline.SetRejectMessage("\nSelect a polyline!");
                            Prompt_centerline.AllowNone = true;
                            Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                            Rezultat_centerline = ThisDrawing.Editor.GetEntity(Prompt_centerline);

                            if (Rezultat_centerline.Status != PromptStatus.OK)
                            {

                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }

                            poly1 = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForWrite) as Polyline;
                            pt_picked = Rezultat_centerline.PickedPoint;

                        }

                        if (poly1 != null)
                        {
                            id_poly = poly1.ObjectId;
                            double param_picked = poly1.GetParameterAtPoint(poly1.GetClosestPointTo(pt_picked, Vector3d.ZAxis, false));

                            int param1 = Convert.ToInt32(Math.Floor(param_picked));

                            double b1 = poly1.GetBulgeAt(param1);

                            if (b1 != 0 && param1 + 1 <= poly1.EndParam && poly1.GetBulgeAt(param1 + 1) == 0)
                            {
                                LineSegment2d ls1 = poly1.GetLineSegment2dAt(param1 - 1);
                                LineSegment2d ls2 = poly1.GetLineSegment2dAt(param1 + 1);
                                Polyline p1 = new Polyline();
                                p1.AddVertexAt(0, ls1.StartPoint, 0, 0, 0);
                                p1.AddVertexAt(1, ls1.EndPoint, 0, 0, 0);
                                Polyline p2 = new Polyline();
                                p2.AddVertexAt(0, ls2.StartPoint, 0, 0, 0);
                                p2.AddVertexAt(1, ls2.EndPoint, 0, 0, 0);
                                Point3dCollection col2 = Functions.Intersect_with_extend_both(p1, p2);
                                poly1.RemoveVertexAt(param1 + 1);
                                poly1.RemoveVertexAt(param1);
                                poly1.AddVertexAt(param1, new Point2d(col2[0].X, col2[0].Y), 0, 0, 0);

                            }
                        }

                        Trans1.TransactionManager.QueueForGraphicsFlush();
                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                Editor1.SetImpliedSelection(Empty_array);
            }


            Editor1.WriteMessage("\nCommand:");
            if (id_poly != ObjectId.Null) Autodesk.AutoCAD.Internal.Utils.SelectObjects(new ObjectId[] { id_poly });



        }


        public void fillet_with_radius(double r1 = 1)
        {
            ObjectId[] Empty_array = null;
            ObjectId id_poly = ObjectId.Null;
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

                        Polyline poly1 = null;
                        Point3d pt_picked = new Point3d();
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                        bool was_selected = false;


                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1 = Editor1.SelectImplied();
                        if (Rezultat1.Status == PromptStatus.OK)
                        {
                            if (Rezultat1.Value.Count == 1)
                            {
                                ObjectId id1 = Rezultat1.Value[0].ObjectId;

                                poly1 = Trans1.GetObject(id1, OpenMode.ForWrite) as Polyline;

                                if (poly1 != null)
                                {
                                    Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                    Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                    PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSelect the node:");
                                    PP1.AllowNone = false;
                                    Point_res1 = Editor1.GetPoint(PP1);
                                    if (Point_res1.Status != PromptStatus.OK)
                                    {
                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                        return;
                                    }
                                    pt_picked = poly1.GetClosestPointTo(Point_res1.Value, Vector3d.ZAxis, false);
                                    was_selected = true;
                                }
                            }
                        }

                        if (was_selected == false)
                        {
                            Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                            Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                            Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the node:");
                            Prompt_centerline.SetRejectMessage("\nSelect a polyline!");
                            Prompt_centerline.AllowNone = true;
                            Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                            Rezultat_centerline = ThisDrawing.Editor.GetEntity(Prompt_centerline);
                            if (Rezultat_centerline.Status != PromptStatus.OK)
                            {
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }
                            poly1 = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForWrite) as Polyline;
                            pt_picked = Rezultat_centerline.PickedPoint;
                        }

                        if (poly1 != null)
                        {
                            double param_picked = poly1.GetParameterAtPoint(poly1.GetClosestPointTo(pt_picked, Vector3d.ZAxis, false));
                            double param1 = Math.Round(param_picked, 0);
                            int pp1 = Convert.ToInt32(Math.Floor(param_picked));
                            double b1 = poly1.GetBulgeAt(pp1);

                            if (b1 != 0 && pp1 + 1 <= poly1.EndParam && poly1.GetBulgeAt(pp1 + 1) == 0)
                            {
                                LineSegment2d ls1 = poly1.GetLineSegment2dAt(pp1 - 1);
                                LineSegment2d ls2 = poly1.GetLineSegment2dAt(pp1 + 1);
                                Polyline p1 = new Polyline();
                                p1.AddVertexAt(0, ls1.StartPoint, 0, 0, 0);
                                p1.AddVertexAt(1, ls1.EndPoint, 0, 0, 0);
                                Polyline p2 = new Polyline();
                                p2.AddVertexAt(0, ls2.StartPoint, 0, 0, 0);
                                p2.AddVertexAt(1, ls2.EndPoint, 0, 0, 0);
                                Point3dCollection col2 = Functions.Intersect_with_extend_both(p1, p2);
                                poly1.RemoveVertexAt(pp1 + 1);
                                poly1.RemoveVertexAt(pp1);
                                poly1.AddVertexAt(pp1, new Point2d(col2[0].X, col2[0].Y), 0, 0, 0);
                                param1 = poly1.GetParameterAtPoint(new Point3d(col2[0].X, col2[0].Y, poly1.Elevation));
                            }



                            Point3d node1 = poly1.GetPointAtParameter(param1);
                            Point3d node0 = poly1.GetPointAtParameter(param1 - 1);
                            Point3d node2 = poly1.GetPointAtParameter(param1 + 1);

                            double l1 = Math.Pow(Math.Pow((node0.X - node1.X), 2) + Math.Pow((node0.Y - node1.Y), 2), 0.5);
                            double l2 = Math.Pow(Math.Pow((node1.X - node2.X), 2) + Math.Pow((node1.Y - node2.Y), 2), 0.5);

                            double defl1 = Functions.Get_deflection_angle_rad(node0.X, node0.Y, node1.X, node1.Y, node2.X, node2.Y);
                            string side1 = Functions.Get_deflection_side(node0.X, node0.Y, node1.X, node1.Y, node2.X, node2.Y);
                            double angle1 = Math.PI - defl1;
                            double alpha = angle1 / 2;
                            double x = r1 / Math.Tan(alpha);

                            double middle = poly1.GetDistanceAtParameter(param1);
                            Point3d start1 = poly1.GetPointAtDist(middle - x);
                            Point3d end1 = poly1.GetPointAtDist(middle + x);

                            Autodesk.AutoCAD.DatabaseServices.Line linie1 = new Autodesk.AutoCAD.DatabaseServices.Line(node1, start1);
                            linie1.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, start1));

                            Autodesk.AutoCAD.DatabaseServices.Line linie2 = new Autodesk.AutoCAD.DatabaseServices.Line(node1, end1);
                            linie2.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, end1));

                            Point3dCollection col1 = new Point3dCollection();
                            linie1.IntersectWith(linie2, Intersect.ExtendBoth, col1, IntPtr.Zero, IntPtr.Zero);

                            if (col1.Count > 0)
                            {
                                Point3d center1 = col1[0];
                                double bear_start = Functions.GET_Bearing_rad(center1.X, center1.Y, start1.X, start1.Y);
                                double bear_end = Functions.GET_Bearing_rad(center1.X, center1.Y, end1.X, end1.Y);
                                int lr = 1;
                                if (side1 == "RT")
                                {
                                    double T = bear_start;
                                    bear_start = bear_end;
                                    bear_end = T;
                                    lr = -1;
                                }
                                Autodesk.AutoCAD.DatabaseServices.Arc arc1 = new Autodesk.AutoCAD.DatabaseServices.Arc(center1, r1, bear_start, bear_end);
                                Point3d pt1 = poly1.GetClosestPointTo(arc1.StartPoint, Vector3d.ZAxis, false);
                                Point3d pt2 = poly1.GetClosestPointTo(arc1.EndPoint, Vector3d.ZAxis, false);

                                double d1 = Math.Pow(Math.Pow(pt1.X - arc1.StartPoint.X, 2) + Math.Pow(pt1.Y - arc1.StartPoint.Y, 2), 0.5);
                                double d2 = Math.Pow(Math.Pow(pt2.X - arc1.EndPoint.X, 2) + Math.Pow(pt2.Y - arc1.EndPoint.Y, 2), 0.5);

                                if (d1 + d2 < 0.01)
                                {
                                    double par1 = poly1.GetParameterAtPoint(pt1);
                                    double par2 = poly1.GetParameterAtPoint(pt2);
                                    if (par1 > par2)
                                    {
                                        double t = par1;
                                        par1 = par2;
                                        par2 = t;
                                        Point3d tt = pt1;
                                        pt1 = pt2;
                                        pt2 = tt;
                                    }
                                    poly1.AddVertexAt((Convert.ToInt32(Math.Floor(par2))), new Point2d(pt2.X, pt2.Y), 0, 0, 0);
                                    poly1.AddVertexAt((Convert.ToInt32(Math.Ceiling(par1))), new Point2d(pt1.X, pt1.Y), 0, 0, 0);
                                    poly1.RemoveVertexAt(Convert.ToInt32(Math.Ceiling(par1)) + 2);
                                    double angle = bear_end - bear_start;
                                    double bulge1 = lr * Math.Tan(angle / 4);
                                    poly1.SetBulgeAt(Convert.ToInt32(Math.Ceiling(par1)), bulge1);


                                    id_poly = poly1.ObjectId;
                                    Editor1.WriteMessage(id_poly.ToString());
                                }
                            }
                        }
                        Trans1.TransactionManager.QueueForGraphicsFlush();
                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                Editor1.SetImpliedSelection(Empty_array);
            }
            Editor1.WriteMessage("\nCommand:");
            if (id_poly != ObjectId.Null) Autodesk.AutoCAD.Internal.Utils.SelectObjects(new ObjectId[] { id_poly });

        }




        [CommandMethod("drop1", CommandFlags.UsePickSet)]
        public void drop1()
        {
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
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;

                        bool was_selected = false;
                        Polyline poly1 = null;

                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1 = Editor1.SelectImplied();
                        if (Rezultat1.Status == PromptStatus.OK)
                        {
                            if (Rezultat1.Value.Count == 1)
                            {
                                ObjectId id1 = Rezultat1.Value[0].ObjectId;

                                poly1 = Trans1.GetObject(id1, OpenMode.ForWrite) as Polyline;

                                if (poly1 != null)
                                {
                                    was_selected = true;
                                }
                            }
                        }


                        if (was_selected == false)
                        {
                            Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_poly1;
                            Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_poly1;
                            Prompt_poly1 = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the polyline:");
                            Prompt_poly1.SetRejectMessage("\nSelect a polyline!");
                            Prompt_poly1.AllowNone = true;
                            Prompt_poly1.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                            Rezultat_poly1 = ThisDrawing.Editor.GetEntity(Prompt_poly1);

                            if (Rezultat_poly1.Status != PromptStatus.OK)
                            {
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }

                            poly1 = Trans1.GetObject(Rezultat_poly1.ObjectId, OpenMode.ForWrite) as Polyline;
                        }


                        if (poly1 != null)
                        {
                            Autodesk.AutoCAD.EditorInput.PromptPointResult pres1;
                            Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                            PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nStart:");
                            PP1.AllowNone = false;
                            pres1 = Editor1.GetPoint(PP1);

                            if (pres1.Status != PromptStatus.OK)
                            {

                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }

                            Autodesk.AutoCAD.EditorInput.PromptPointResult pres2;
                            Autodesk.AutoCAD.EditorInput.PromptPointOptions PP2;
                            PP2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nEnd:");
                            PP2.AllowNone = false;
                            PP2.UseBasePoint = true;
                            PP2.BasePoint = pres1.Value;
                            pres2 = Editor1.GetPoint(PP2);

                            if (pres2.Status != PromptStatus.OK)
                            {

                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }

                            Point3d point_on_poly1 = poly1.GetClosestPointTo(pres1.Value, Vector3d.ZAxis, false);
                            Point3d point_on_poly2 = poly1.GetClosestPointTo(pres2.Value, Vector3d.ZAxis, false);

                            double param1 = poly1.GetParameterAtPoint(point_on_poly1);
                            double param2 = poly1.GetParameterAtPoint(point_on_poly2);


                            Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_poly2;
                            Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_poly2;
                            Prompt_poly2 = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the Offset polyline:");
                            Prompt_poly2.SetRejectMessage("\nSelect a polyline!");
                            Prompt_poly2.AllowNone = true;
                            Prompt_poly2.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                            Rezultat_poly2 = ThisDrawing.Editor.GetEntity(Prompt_poly2);

                            if (Rezultat_poly2.Status != PromptStatus.OK)
                            {
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }

                            Polyline poly2 = Trans1.GetObject(Rezultat_poly2.ObjectId, OpenMode.ForRead) as Polyline;

                            Point3d point_on_poly3 = poly2.GetClosestPointTo(pres1.Value, Vector3d.ZAxis, false);
                            Point3d point_on_poly4 = poly2.GetClosestPointTo(pres2.Value, Vector3d.ZAxis, false);

                            double param3 = poly2.GetParameterAtPoint(point_on_poly3);
                            double param4 = poly2.GetParameterAtPoint(point_on_poly4);

                            if (param1 > param2)
                            {
                                double t = param1;
                                param1 = param2;
                                param2 = t;
                            }
                            if (param3 > param4)
                            {
                                double t = param3;
                                param3 = param4;
                                param4 = t;
                            }

                            for (int i = Convert.ToInt32(Math.Ceiling(param1)); i < param2; i++)
                            {
                                double bulge1 = poly1.GetBulgeAt(i);
                                if (bulge1 != 0 && i + 1 <= poly1.EndParam && i - 1 >= 0)
                                {
                                    LineSegment2d ls1 = poly1.GetLineSegment2dAt(i - 1);
                                    LineSegment2d ls2 = poly1.GetLineSegment2dAt(i + 1);
                                    Polyline poly4 = new Polyline();
                                    poly4.AddVertexAt(0, ls1.StartPoint, 0, 0, 0);
                                    poly4.AddVertexAt(1, ls1.EndPoint, 0, 0, 0);
                                    Polyline poly5 = new Polyline();
                                    poly5.AddVertexAt(0, ls2.StartPoint, 0, 0, 0);
                                    poly5.AddVertexAt(1, ls2.EndPoint, 0, 0, 0);
                                    Point3dCollection col2 = Functions.Intersect_with_extend_both(poly4, poly5);
                                    poly1.RemoveVertexAt(i + 1);
                                    poly1.RemoveVertexAt(i);
                                    poly1.AddVertexAt(i, new Point2d(col2[0].X, col2[0].Y), 0, 0, 0);
                                }
                                Point2d p1 = poly1.GetPoint2dAt(i);
                                Polyline poly3 = new Polyline();
                                poly3.AddVertexAt(0, new Point2d(p1.X, p1.Y - 1000), 0, 0, 0);
                                poly3.AddVertexAt(1, new Point2d(p1.X, p1.Y + 1000), 0, 0, 0);

                                Point3dCollection colint = Functions.Intersect_on_both_operands(poly3, poly2);
                                if (colint.Count > 0)
                                {
                                    Point3d pt2 = colint[0];
                                    double Yref = pt2.Y;

                                    //double param5 = poly2.GetParameterAtPoint(pt2);
                                    //int param51 = Convert.ToInt32(Math.Floor(param5));
                                    //int param52 = Convert.ToInt32(Math.Ceiling(param5));

                                    //Point2d pt51 = poly2.GetPoint2dAt(param51);
                                    //Point2d pt52 = poly2.GetPoint2dAt(param52);

                                    //if (pt51.Y < Yref) Yref = pt51.Y;
                                    //if (pt52.Y < Yref) Yref = pt52.Y;

                                    if (Yref - 0.1 < p1.Y)
                                    {
                                        poly1.RemoveVertexAt(i);
                                        poly1.AddVertexAt(i, new Point2d(p1.X, Yref - 0.1), 0, 0, 0);
                                    }
                                }
                            }
                        }

                        Trans1.TransactionManager.QueueForGraphicsFlush();
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

        [CommandMethod("mtxt2ml_bore", CommandFlags.UsePickSet)]
        public void creaza_mleader_from_mtext_pt_boreholes()
        {

            MLeader ml1 = null;
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
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                        Matrix3d CurentUCSmatrix = Editor1.CurrentUserCoordinateSystem;
                        MLeaderStyle mlstyle1 = Trans1.GetObject(ThisDrawing.Database.MLeaderstyle, OpenMode.ForRead) as MLeaderStyle;
                        double textrot = Functions.GET_Bearing_rad(curent_ucs_matrix.CoordinateSystem3d.Xaxis);



                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1 = Editor1.SelectImplied();


                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rez.MessageForAdding = "\nSelect Mtext:";
                            Prompt_rez.SingleOnly = false;
                            Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);
                        }




                        if (Rezultat1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }












                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {

                            MText mt1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForWrite) as MText;
                            if (mt1 != null)
                            {
                                Point3d pt1 = mt1.Location;

                                bool towards_left = false;

                                if (mt1.Attachment == AttachmentPoint.MiddleRight) towards_left = true;

                                Autodesk.AutoCAD.Colors.Color color1 = mt1.Color;

                                ml1 = Functions.creaza_mleader_with_color(color1, pt1, mt1.Contents, mt1.TextStyleId, mt1.TextHeight, 2 * mt1.TextHeight, 0, mt1.TextHeight, mt1.TextHeight, mt1.TextHeight, towards_left, mt1.Layer);

                                using (MText mt2 = ml1.MText)
                                {
                                    mt2.Rotation = 0;
                                    mt2.BackgroundFill = mt1.BackgroundFill;
                                    mt2.UseBackgroundColor = mt1.UseBackgroundColor;
                                    mt2.BackgroundScaleFactor = mt1.BackgroundScaleFactor;
                                    ml1.MText = mt2;
                                }

                                ml1.TextAttachmentType = TextAttachmentType.AttachmentMiddleOfTop;
                                ml1.LineWeight = LineWeight.LineWeight000;
                                ml1.LeaderLineWeight = LineWeight.LineWeight000;

                                mt1.Erase();
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




        }


        [CommandMethod("w2xl4", CommandFlags.UsePickSet | CommandFlags.Redraw)]
        public void write_poly_to_excel4()
        {

            char c = (char)34;
            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {
                System.Data.DataTable dt1 = new System.Data.DataTable();

                string col_type = "Type of Object";
                string col_CODE = "C# CODE";
                string col_layer = "Layer";
                string col_sta = "Station";
                string col_x = "x";
                string col_y = "y";
                string col_z = "z";
                string col_defl = "Deflection DD";
                string col_defl_dms = "Deflection DMS";

                string col_angle = "Rotation DD";
                string col_angle_dms = "Rotation DMS";
                string col_bulge = "Bulge Angle [Rad]";
                string col_radius = "Arc Radius";
                string col_midX = "Arc Middle X";
                string col_midY = "Arc Middle Y";
                string col_arc_len = "Arc Length";
                string col_length = "Length";
                string col_textstring = "Text String";
                string col_textheight = "Text Height";
                string col_rotationtxt = "Text Rotation";
                string col_Blockname = "Block Name";
                string col_rotationblock = "Block Rotation";
                string col_blockscale = "Block Scale";
                string col_radius1 = "Circle Radius";


                dt1.Columns.Add(col_CODE, typeof(string));
                dt1.Columns.Add(col_type, typeof(string));
                dt1.Columns.Add(col_layer, typeof(string));
                dt1.Columns.Add(col_sta, typeof(double));
                dt1.Columns.Add(col_x, typeof(double));
                dt1.Columns.Add(col_y, typeof(double));
                dt1.Columns.Add(col_z, typeof(double));
                dt1.Columns.Add(col_defl, typeof(double));
                dt1.Columns.Add(col_defl_dms, typeof(string));
                dt1.Columns.Add(col_angle, typeof(double));
                dt1.Columns.Add(col_angle_dms, typeof(string));
                dt1.Columns.Add(col_bulge, typeof(double));
                dt1.Columns.Add(col_radius, typeof(double));
                dt1.Columns.Add(col_arc_len, typeof(double));
                dt1.Columns.Add(col_midX, typeof(double));
                dt1.Columns.Add(col_midY, typeof(double));
                dt1.Columns.Add(col_length, typeof(double));
                dt1.Columns.Add(col_textstring, typeof(string));
                dt1.Columns.Add(col_textheight, typeof(double));
                dt1.Columns.Add(col_rotationtxt, typeof(double));
                dt1.Columns.Add(col_Blockname, typeof(string));
                dt1.Columns.Add(col_rotationblock, typeof(double));
                dt1.Columns.Add(col_blockscale, typeof(double));
                dt1.Columns.Add(col_radius1, typeof(double));


                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;


                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1 = Editor1.SelectImplied();
                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rez.MessageForAdding = "\nSelect objects:";
                            Prompt_rez.SingleOnly = false;
                            Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                            if (Rezultat1.Status != PromptStatus.OK)
                            {
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }
                        }
                        if (Rezultat1.Status == PromptStatus.OK)
                        {

                            int index_pt = 1;
                            int index_circle = 1;
                            int index_line = 1;
                            int index_polyline = 1;
                            int index_3dpolyline = 1;
                            int index_mtext = 1;
                            int index_text = 1;
                            int index_blk = 1;
                            int index_mleader = 1;

                            for (int i = 0; i < Rezultat1.Value.Count; i++)
                            {
                                Entity ent1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as Entity;
                                if (ent1 != null)
                                {
                                    if (ent1 is DBPoint)
                                    {
                                        DBPoint pt1 = ent1 as DBPoint;
                                        dt1.Rows.Add();
                                        dt1.Rows[dt1.Rows.Count - 1][col_type] = "DBPoint#" + Convert.ToString(index_pt);
                                        dt1.Rows[dt1.Rows.Count - 1][col_x] = pt1.Position.X;
                                        dt1.Rows[dt1.Rows.Count - 1][col_y] = pt1.Position.Y;
                                        dt1.Rows[dt1.Rows.Count - 1][col_z] = pt1.Position.Z;
                                        dt1.Rows[dt1.Rows.Count - 1][col_layer] = pt1.Layer;
                                        ++index_pt;
                                    }
                                    else if (ent1 is MText)
                                    {
                                        MText mtext1 = ent1 as MText;
                                        dt1.Rows.Add();
                                        dt1.Rows[dt1.Rows.Count - 1][col_type] = "Mtext#" + Convert.ToString(index_mtext);
                                        dt1.Rows[dt1.Rows.Count - 1][col_x] = mtext1.Location.X;
                                        dt1.Rows[dt1.Rows.Count - 1][col_y] = mtext1.Location.Y;
                                        dt1.Rows[dt1.Rows.Count - 1][col_z] = mtext1.Location.Z;
                                        dt1.Rows[dt1.Rows.Count - 1][col_layer] = mtext1.Layer;
                                        dt1.Rows[dt1.Rows.Count - 1][col_textstring] = mtext1.Contents;
                                        dt1.Rows[dt1.Rows.Count - 1][col_textheight] = mtext1.TextHeight;
                                        dt1.Rows[dt1.Rows.Count - 1][col_rotationtxt] = 180 * mtext1.Rotation / Math.PI;
                                        ++index_mtext;
                                    }
                                    else if (ent1 is DBText)
                                    {
                                        DBText text1 = ent1 as DBText;
                                        dt1.Rows.Add();
                                        dt1.Rows[dt1.Rows.Count - 1][col_type] = "text#" + Convert.ToString(index_text);
                                        dt1.Rows[dt1.Rows.Count - 1][col_x] = text1.Position.X;
                                        dt1.Rows[dt1.Rows.Count - 1][col_y] = text1.Position.Y;
                                        dt1.Rows[dt1.Rows.Count - 1][col_z] = text1.Position.Z;
                                        dt1.Rows[dt1.Rows.Count - 1][col_layer] = text1.Layer;
                                        dt1.Rows[dt1.Rows.Count - 1][col_textstring] = text1.TextString;
                                        dt1.Rows[dt1.Rows.Count - 1][col_textheight] = text1.Height;
                                        dt1.Rows[dt1.Rows.Count - 1][col_rotationtxt] = 180 * text1.Rotation / Math.PI;
                                        ++index_text;
                                    }
                                    else if (ent1 is BlockReference)
                                    {
                                        BlockReference br1 = ent1 as BlockReference;
                                        dt1.Rows.Add();
                                        dt1.Rows[dt1.Rows.Count - 1][col_type] = "BlockReference#" + Convert.ToString(index_blk);
                                        dt1.Rows[dt1.Rows.Count - 1][col_x] = br1.Position.X;
                                        dt1.Rows[dt1.Rows.Count - 1][col_y] = br1.Position.Y;
                                        dt1.Rows[dt1.Rows.Count - 1][col_z] = br1.Position.Z;
                                        dt1.Rows[dt1.Rows.Count - 1][col_layer] = br1.Layer;
                                        string bn = Functions.get_block_name(br1);
                                        dt1.Rows[dt1.Rows.Count - 1][col_Blockname] = bn;
                                        dt1.Rows[dt1.Rows.Count - 1][col_blockscale] = br1.ScaleFactors.X;
                                        dt1.Rows[dt1.Rows.Count - 1][col_rotationblock] = 180 * br1.Rotation / Math.PI;

                                        if (br1.AttributeCollection.Count > 0)
                                        {
                                            for (int k = 0; k < br1.AttributeCollection.Count; k++)
                                            {
                                                ObjectId id1 = br1.AttributeCollection[k];
                                                AttributeReference atr1 = Trans1.GetObject(id1, OpenMode.ForRead) as AttributeReference;
                                                if (atr1 != null)
                                                {
                                                    string atr_name = atr1.Tag;
                                                    string val1 = atr1.TextString;
                                                    if (atr1.IsMTextAttribute == true)
                                                    {
                                                        val1 = atr1.MTextAttribute.Contents;
                                                    }

                                                    string colname = bn + "_" + atr_name;
                                                    if (dt1.Columns.Contains(colname) == false)
                                                    {
                                                        dt1.Columns.Add(colname, typeof(string));
                                                    }
                                                    dt1.Rows[dt1.Rows.Count - 1][colname] = val1;
                                                }
                                            }
                                        }
                                        ++index_blk;
                                    }
                                    else if (ent1 is Circle)
                                    {
                                        Circle circ1 = ent1 as Circle;

                                        dt1.Rows.Add();

                                        dt1.Rows[dt1.Rows.Count - 1][col_x] = circ1.Center.X;
                                        dt1.Rows[dt1.Rows.Count - 1][col_y] = circ1.Center.Y;
                                        dt1.Rows[dt1.Rows.Count - 1][col_z] = circ1.Center.Z;
                                        dt1.Rows[dt1.Rows.Count - 1][col_layer] = circ1.Layer;
                                        dt1.Rows[dt1.Rows.Count - 1][col_radius1] = circ1.Radius;


                                        dt1.Rows[dt1.Rows.Count - 1][col_CODE] = "Circle cerc" + Convert.ToString(index_circle) + " = new Circle(new Point3d(" +
                                            Convert.ToString(circ1.Center.X) + ", " + Convert.ToString(circ1.Center.Y) + ", 0), Vector3d.ZAxis, scale1*" +
                                            Convert.ToString(circ1.Radius) + ");";


                                        dt1.Rows[dt1.Rows.Count - 1][col_type] = "Circle#" + Convert.ToString(index_circle);

                                        dt1.Rows.Add();
                                        dt1.Rows[dt1.Rows.Count - 1][col_CODE] = "cerc" + Convert.ToString(index_circle) + ".TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));";

                                        dt1.Rows[dt1.Rows.Count - 1][col_type] = "Circle#" + Convert.ToString(index_circle);



                                        dt1.Rows.Add();
                                        dt1.Rows[dt1.Rows.Count - 1][col_CODE] = "BTrecord.AppendEntity(cerc" + Convert.ToString(index_circle) + ");";
                                        dt1.Rows[dt1.Rows.Count - 1][col_type] = "Circle#" + Convert.ToString(index_circle);

                                        dt1.Rows.Add();
                                        dt1.Rows[dt1.Rows.Count - 1][col_CODE] = "Trans1.AddNewlyCreatedDBObject(cerc" + Convert.ToString(index_circle) + ", true);";
                                        dt1.Rows[dt1.Rows.Count - 1][col_type] = "Circle#" + Convert.ToString(index_circle);


                                        dt1.Rows.Add();
                                        dt1.Rows[dt1.Rows.Count - 1][col_CODE] = " Hatch  hatch" + Convert.ToString(index_circle) +
                                            " = CreateHatch(cerc" + Convert.ToString(index_circle) + "," +
                                            Convert.ToChar(34) + "SOLID" + Convert.ToChar(34) + ", 1, 0);";
                                        dt1.Rows[dt1.Rows.Count - 1][col_type] = "Circle#" + Convert.ToString(index_circle);

                                        dt1.Rows.Add();
                                        dt1.Rows[dt1.Rows.Count - 1][col_CODE] = "hatch" + Convert.ToString(index_circle) +
                                            ".Layer = " +
                                            Convert.ToChar(34) + "0" + Convert.ToChar(34) + ";";
                                        dt1.Rows[dt1.Rows.Count - 1][col_type] = "Circle#" + Convert.ToString(index_circle);

                                        dt1.Rows.Add();
                                        dt1.Rows[dt1.Rows.Count - 1][col_CODE] = "hatch" + Convert.ToString(index_circle) +
                                            ".LineWeight  = LineWeight.LineWeight000;";
                                        dt1.Rows[dt1.Rows.Count - 1][col_type] = "Circle#" + Convert.ToString(index_circle);

                                        dt1.Rows.Add();
                                        dt1.Rows[dt1.Rows.Count - 1][col_CODE] = "hatch" + Convert.ToString(index_circle) +
                                            ".Color  = colorSC;";
                                        dt1.Rows[dt1.Rows.Count - 1][col_type] = "Circle#" + Convert.ToString(index_circle);


                                        dt1.Rows.Add();
                                        dt1.Rows[dt1.Rows.Count - 1][col_CODE] = " bltrec1.AppendEntity(hatch" + Convert.ToString(index_circle) + ");";
                                        dt1.Rows[dt1.Rows.Count - 1][col_type] = "Circle#" + Convert.ToString(index_circle);


                                        dt1.Rows.Add();
                                        dt1.Rows[dt1.Rows.Count - 1][col_CODE] = "cerc" + Convert.ToString(index_circle) + ".Erase();";
                                        dt1.Rows[dt1.Rows.Count - 1][col_type] = "Circle#" + Convert.ToString(index_circle);

                                        ++index_circle;
                                    }
                                    else if (ent1 is Autodesk.AutoCAD.DatabaseServices.Line)
                                    {
                                        Autodesk.AutoCAD.DatabaseServices.Line line1 = ent1 as Autodesk.AutoCAD.DatabaseServices.Line;

                                        double x1 = line1.StartPoint.X;
                                        double y1 = line1.StartPoint.Y;
                                        double x2 = line1.EndPoint.X;
                                        double y2 = line1.EndPoint.Y;

                                        dt1.Rows.Add();
                                        dt1.Rows[dt1.Rows.Count - 1][col_type] = "Line#" + Convert.ToString(index_line);
                                        dt1.Rows[dt1.Rows.Count - 1][col_x] = x1;
                                        dt1.Rows[dt1.Rows.Count - 1][col_y] = y1;
                                        dt1.Rows[dt1.Rows.Count - 1][col_z] = line1.StartPoint.Z;
                                        dt1.Rows[dt1.Rows.Count - 1][col_layer] = line1.Layer;
                                        dt1.Rows.Add();
                                        dt1.Rows[dt1.Rows.Count - 1][col_type] = "Line#" + Convert.ToString(index_line);
                                        dt1.Rows[dt1.Rows.Count - 1][col_x] = x2;
                                        dt1.Rows[dt1.Rows.Count - 1][col_y] = y2;
                                        dt1.Rows[dt1.Rows.Count - 1][col_z] = line1.EndPoint.Z;
                                        dt1.Rows[dt1.Rows.Count - 1][col_layer] = line1.Layer;
                                        dt1.Rows[dt1.Rows.Count - 1][col_length] = line1.Length;


                                        double rot1 = Functions.GET_Bearing_rad(x1, y1, x2, y2);
                                        string dms = Functions.Get_DMS(rot1 * 180 / Math.PI, 0);
                                        dt1.Rows[dt1.Rows.Count - 1][col_angle] = rot1 * 180 / Math.PI;
                                        dt1.Rows[dt1.Rows.Count - 1][col_angle_dms] = dms;


                                        ++index_line;
                                    }
                                    else if (ent1 is Autodesk.AutoCAD.DatabaseServices.Polyline)
                                    {
                                        Autodesk.AutoCAD.DatabaseServices.Polyline poly1 = ent1 as Autodesk.AutoCAD.DatabaseServices.Polyline;



                                        dt1.Rows.Add();

                                        dt1.Rows[dt1.Rows.Count - 1][col_CODE] = "Polyline polys" + Convert.ToString(index_polyline) + " = new Polyline();";

                                        for (int j = 0; j < poly1.NumberOfVertices; j++)
                                        {

                                            double x1 = poly1.GetPoint2dAt(j).X;
                                            double y1 = poly1.GetPoint2dAt(j).Y;

                                            double bulge1 = poly1.GetBulgeAt(j);



                                            dt1.Rows.Add();

                                            dt1.Rows[dt1.Rows.Count - 1][col_type] = "Polyline#" + Convert.ToString(index_polyline);

                                            dt1.Rows[dt1.Rows.Count - 1][col_CODE] = "polys" + Convert.ToString(index_polyline) +
                                                ".AddVertexAt(" + Convert.ToString(j) +
                                                ", new Point2d(scale1 * " + Convert.ToString(x1) + ", scale1 *" + Convert.ToString(y1) + "), " +
                                                Convert.ToString(bulge1) + ", 0, 0);";

                                            dt1.Rows[dt1.Rows.Count - 1][col_x] = x1;
                                            dt1.Rows[dt1.Rows.Count - 1][col_y] = y1;
                                            dt1.Rows[dt1.Rows.Count - 1][col_z] = poly1.Elevation;
                                            dt1.Rows[dt1.Rows.Count - 1][col_sta] = poly1.GetDistanceAtParameter(j);
                                            dt1.Rows[dt1.Rows.Count - 1][col_bulge] = bulge1;
                                            dt1.Rows[dt1.Rows.Count - 1][col_layer] = poly1.Layer;

                                            if (bulge1 != 0 && j < poly1.NumberOfVertices - 1)
                                            {
                                                CircularArc2d arc1 = poly1.GetArcSegment2dAt(j);

                                                double radius1 = arc1.Radius;
                                                if (j < poly1.NumberOfVertices - 1)
                                                {
                                                    Polyline poly2 = new Polyline();
                                                    poly2.AddVertexAt(0, poly1.GetPoint2dAt(j), bulge1, 0, 0);
                                                    poly2.AddVertexAt(1, poly1.GetPoint2dAt(j + 1), 0, 0, 0);
                                                    Point3d ptmid = poly2.GetPointAtDist(poly2.Length / 2);
                                                    dt1.Rows[dt1.Rows.Count - 1][col_midX] = ptmid.X;
                                                    dt1.Rows[dt1.Rows.Count - 1][col_midY] = ptmid.Y;
                                                }

                                                double len1 = Math.Abs(radius1 * 4 * Math.Atan(bulge1));
                                                dt1.Rows[dt1.Rows.Count - 1][col_arc_len] = len1;
                                                dt1.Rows[dt1.Rows.Count - 1][col_radius] = radius1;

                                            }


                                            if (j == 0)
                                            {
                                                dt1.Rows[dt1.Rows.Count - 1][col_length] = poly1.Length;

                                            }

                                            if (j > 0 && j < poly1.NumberOfVertices - 1)
                                            {
                                                double x0 = poly1.GetPoint2dAt(j - 1).X;
                                                double y0 = poly1.GetPoint2dAt(j - 1).Y;
                                                double x2 = poly1.GetPoint2dAt(j + 1).X;
                                                double y2 = poly1.GetPoint2dAt(j + 1).Y;


                                                double rot1 = Functions.Get_deflection_angle_rad(x0, y0, x1, y1, x2, y2);
                                                string dms = Functions.Get_deflection_angle_dms(x0, y0, x1, y1, x2, y2);
                                                dt1.Rows[dt1.Rows.Count - 1][col_defl] = rot1 * 180 / Math.PI;
                                                dt1.Rows[dt1.Rows.Count - 1][col_defl_dms] = dms;
                                            }

                                        }
                                        if (poly1.Closed == true)
                                        {
                                            dt1.Rows.Add();

                                            dt1.Rows[dt1.Rows.Count - 1][col_CODE] = "polys" + Convert.ToString(index_polyline) + ".Closed=true;";
                                        }

                                        dt1.Rows.Add();
                                        dt1.Rows[dt1.Rows.Count - 1][col_CODE] = "polys" + Convert.ToString(index_polyline) + ".TransformBy(Matrix3d.Displacement(new Point3d(0, 0, 0).GetVectorTo(poly1.GetPoint3dAt(3))));";
                                        dt1.Rows.Add();
                                        dt1.Rows[dt1.Rows.Count - 1][col_CODE] = "polys" + Convert.ToString(index_polyline) + ".Layer = " + c + "0" + c + ";";
                                        dt1.Rows.Add();
                                        dt1.Rows[dt1.Rows.Count - 1][col_CODE] = "polys" + Convert.ToString(index_polyline) + ".Color = colorA;";
                                        dt1.Rows.Add();
                                        dt1.Rows[dt1.Rows.Count - 1][col_CODE] = "polys" + Convert.ToString(index_polyline) + ".LineWeight = LineWeight.LineWeight000;";
                                        dt1.Rows.Add();
                                        dt1.Rows[dt1.Rows.Count - 1][col_CODE] = "bltrec1.AppendEntity(polys" + Convert.ToString(index_polyline) + ");";

                                        ++index_polyline;
                                    }
                                    else if (ent1 is Autodesk.AutoCAD.DatabaseServices.Polyline3d)
                                    {
                                        Autodesk.AutoCAD.DatabaseServices.Polyline3d poly3D = ent1 as Autodesk.AutoCAD.DatabaseServices.Polyline3d;
                                        Polyline poly1 = Functions.Build_2dpoly_from_3d(poly3D);

                                        for (int j = 0; j < poly1.NumberOfVertices; j++)
                                        {
                                            double x1 = poly1.GetPoint2dAt(j).X;
                                            double y1 = poly1.GetPoint2dAt(j).Y;

                                            dt1.Rows.Add();
                                            dt1.Rows[dt1.Rows.Count - 1][col_type] = "Polyline3D#" + Convert.ToString(index_3dpolyline);
                                            dt1.Rows[dt1.Rows.Count - 1][col_x] = x1;
                                            dt1.Rows[dt1.Rows.Count - 1][col_y] = y1;
                                            dt1.Rows[dt1.Rows.Count - 1][col_z] = poly3D.GetPointAtParameter(j).Z;
                                            dt1.Rows[dt1.Rows.Count - 1][col_sta] = poly3D.GetDistanceAtParameter(j);
                                            dt1.Rows[dt1.Rows.Count - 1][col_layer] = poly3D.Layer;
                                            if (j == 0)
                                            {
                                                dt1.Rows[dt1.Rows.Count - 1][col_length] = poly3D.Length;

                                            }


                                            if (j > 0 && j < poly1.NumberOfVertices - 1)
                                            {
                                                double x0 = poly1.GetPoint2dAt(j - 1).X;
                                                double y0 = poly1.GetPoint2dAt(j - 1).Y;
                                                double x2 = poly1.GetPoint2dAt(j + 1).X;
                                                double y2 = poly1.GetPoint2dAt(j + 1).Y;

                                                double rot1 = Functions.Get_deflection_angle_rad(x0, y0, x1, y1, x2, y2);
                                                string dms = Functions.Get_deflection_angle_dms(x0, y0, x1, y1, x2, y2);
                                                dt1.Rows[dt1.Rows.Count - 1][col_defl] = rot1 * 180 / Math.PI;
                                                dt1.Rows[dt1.Rows.Count - 1][col_defl_dms] = dms;
                                            }
                                        }

                                        ++index_3dpolyline;
                                    }
                                    else if (ent1 is MLeader)
                                    {
                                        MLeader ml1 = ent1 as MLeader;
                                        if (ml1 != null)
                                        {
                                            Point3d ptins = ml1.GetFirstVertex(0);

                                            dt1.Rows.Add();
                                            dt1.Rows[dt1.Rows.Count - 1][col_type] = "Mleader#" + Convert.ToString(index_mleader);
                                            dt1.Rows[dt1.Rows.Count - 1][col_x] = ptins.X;
                                            dt1.Rows[dt1.Rows.Count - 1][col_y] = ptins.Y;

                                            dt1.Rows[dt1.Rows.Count - 1][col_layer] = ml1.Layer;
                                            dt1.Rows[dt1.Rows.Count - 1][col_textstring] = ml1.MText.Contents;

                                            ++index_mleader;



                                        }
                                    }
                                }
                            }

                            Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt1);


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


        }

        [CommandMethod("intersector")]
        public void scan_centerline()
        {

            Editor editor1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;

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

                        Autodesk.AutoCAD.DatabaseServices.LayerTable layer_table = (Autodesk.AutoCAD.DatabaseServices.LayerTable)Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

                        //Vous devez ajouter une référence à AecBaseMgd.dll(dans le répertoire d'installation).
                        Matrix3d CurentUCSmatrix = Editor1.CurrentUserCoordinateSystem;

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                        Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the centerline:");
                        Prompt_centerline.SetRejectMessage("\nSelect a polyline or a 3D polyline!");
                        Prompt_centerline.AllowNone = true;
                        Prompt_centerline.AddAllowedClass(typeof(Polyline), false);
                        Prompt_centerline.AddAllowedClass(typeof(Polyline3d), false);
                        Rezultat_centerline = ThisDrawing.Editor.GetEntity(Prompt_centerline);

                        if (Rezultat_centerline.Status != PromptStatus.OK)
                        {
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }




                        Polyline poly2d = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForWrite) as Polyline;
                        Polyline3d poly3d = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForRead) as Polyline3d;
                        Polyline original_poly2d = poly2d.Clone() as Polyline;



                        if (poly2d == null)
                        {
                            poly2d = Functions.Build_2dpoly_from_3d(poly3d);
                        }


                        bool delete_poly3d = false;

                        if (poly3d == null && poly2d != null)
                        {
                            System.Data.DataTable dt_cl = new System.Data.DataTable();
                            dt_cl.Columns.Add("X", typeof(double));
                            dt_cl.Columns.Add("Y", typeof(double));
                            dt_cl.Columns.Add("Z", typeof(double));

                            for (int i = 0; i < poly2d.NumberOfVertices; ++i)
                            {
                                Point2d pt1 = poly2d.GetPoint2dAt(i);
                                dt_cl.Rows.Add();
                                dt_cl.Rows[i]["X"] = pt1.X;
                                dt_cl.Rows[i]["Y"] = pt1.Y;
                                dt_cl.Rows[i]["Z"] = 0;
                            }



                            poly3d = Functions.Build_3d_poly_for_scanning(dt_cl);
                            delete_poly3d = true;
                        }

                        System.Data.DataTable dt1 = new System.Data.DataTable();
                        dt1.Columns.Add("Type of object", typeof(string));
                        dt1.Columns.Add("DWG Layer", typeof(string));
                        dt1.Columns.Add("Block Name", typeof(string));
                        dt1.Columns.Add("Sta", typeof(double));
                        dt1.Columns.Add("X", typeof(double));
                        dt1.Columns.Add("Y", typeof(double));
                        dt1.Columns.Add("Z on CL", typeof(double));
                        dt1.Columns.Add("Z on object", typeof(double));

                        ObjectIdCollection col1 = new ObjectIdCollection();
                        col1.Add(poly2d.ObjectId);
                        col1.Add(poly3d.ObjectId);

                        List<string> lista_layere = new List<string>();
                        foreach (ObjectId Layer_id in layer_table)
                        {
                            LayerTableRecord ltr = (LayerTableRecord)Trans1.GetObject(Layer_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);

                            if (ltr.Name.Contains("|") == false && ltr.Name.Contains("$") == false && ltr.IsFrozen == false && ltr.IsOff == false)
                            {

                                lista_layere.Add(ltr.Name);

                            }
                        }

                        foreach (ObjectId id1 in BTrecord)
                        {
                            if (col1.Contains(id1) == false)
                            {
                                Polyline poly_int_2d = Trans1.GetObject(id1, OpenMode.ForRead) as Polyline;
                                Polyline3d poly_int_3d = Trans1.GetObject(id1, OpenMode.ForRead) as Polyline3d;
                                Autodesk.AutoCAD.DatabaseServices.Line line_int = Trans1.GetObject(id1, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Line;
                                DBPoint dbpt1 = Trans1.GetObject(id1, OpenMode.ForRead) as DBPoint;
                                BlockReference block1 = Trans1.GetObject(id1, OpenMode.ForRead) as BlockReference;
                                Entity ent1 = Trans1.GetObject(id1, OpenMode.ForRead) as Entity;

                                if (ent1 != null && lista_layere.Contains(ent1.Layer) == true)
                                {
                                    if (poly_int_2d != null)
                                    {
                                        poly2d.Elevation = poly_int_2d.Elevation;
                                        Point3dCollection col_int = Functions.Intersect_on_both_operands(poly2d, poly_int_2d);
                                        if (col_int.Count > 0)
                                        {


                                            for (int i = 0; i < col_int.Count; ++i)
                                            {
                                                dt1.Rows.Add();
                                                dt1.Rows[dt1.Rows.Count - 1]["Type of object"] = "Polyline";
                                                dt1.Rows[dt1.Rows.Count - 1]["DWG Layer"] = poly_int_2d.Layer;
                                                dt1.Rows[dt1.Rows.Count - 1]["X"] = col_int[i].X;
                                                dt1.Rows[dt1.Rows.Count - 1]["Y"] = col_int[i].Y;

                                                Point3d pt1 = poly2d.GetClosestPointTo(col_int[i], Vector3d.ZAxis, false);
                                                double param1 = poly2d.GetParameterAtPoint(pt1);
                                                if (param1 > poly3d.EndParam) param1 = poly3d.EndParam;
                                                dt1.Rows[dt1.Rows.Count - 1]["Z on CL"] = poly3d.GetPointAtParameter(param1).Z;
                                                dt1.Rows[dt1.Rows.Count - 1]["Sta"] = poly3d.GetDistanceAtParameter(param1);
                                                dt1.Rows[dt1.Rows.Count - 1]["Z on object"] = col_int[i].Z;

                                                Functions.add_object_data_to_datatable(dt1, Tables1, poly_int_2d.ObjectId);

                                            }


                                        }

                                    }

                                    if (poly_int_3d != null)
                                    {
                                        Polyline poly1 = Functions.Build_2dpoly_from_3d(poly_int_3d);
                                        poly1.Elevation = poly2d.Elevation;



                                        Point3dCollection col_int = Functions.Intersect_on_both_operands(poly2d, poly1);
                                        if (col_int.Count > 0)
                                        {
                                            for (int i = 0; i < col_int.Count; ++i)
                                            {
                                                dt1.Rows.Add();
                                                dt1.Rows[dt1.Rows.Count - 1]["Type of object"] = "Polyline3d";
                                                dt1.Rows[dt1.Rows.Count - 1]["DWG Layer"] = poly_int_3d.Layer;
                                                dt1.Rows[dt1.Rows.Count - 1]["X"] = col_int[i].X;
                                                dt1.Rows[dt1.Rows.Count - 1]["Y"] = col_int[i].Y;

                                                Point3d pt1 = poly2d.GetClosestPointTo(col_int[i], Vector3d.ZAxis, false);
                                                double param1 = poly2d.GetParameterAtPoint(pt1);
                                                if (param1 > poly3d.EndParam) param1 = poly3d.EndParam;
                                                dt1.Rows[dt1.Rows.Count - 1]["Z on CL"] = poly3d.GetPointAtParameter(param1).Z;
                                                dt1.Rows[dt1.Rows.Count - 1]["Sta"] = poly3d.GetDistanceAtParameter(param1);

                                                Point3d p2 = poly1.GetClosestPointTo(col_int[i], Vector3d.ZAxis, false);
                                                double param2 = poly1.GetParameterAtPoint(p2);
                                                if (param2 > poly_int_3d.EndParam) param2 = poly_int_3d.EndParam;

                                                double z_obj = poly_int_3d.GetPointAtParameter(param2).Z;

                                                dt1.Rows[dt1.Rows.Count - 1]["Z on object"] = z_obj;
                                                Functions.add_object_data_to_datatable(dt1, Tables1, poly_int_3d.ObjectId);
                                            }
                                        }
                                    }

                                    if (line_int != null)
                                    {
                                        Polyline poly1 = new Polyline();
                                        poly1.AddVertexAt(0, new Point2d(line_int.StartPoint.X, line_int.StartPoint.Y), 0, 0, 0);
                                        poly1.AddVertexAt(1, new Point2d(line_int.EndPoint.X, line_int.EndPoint.Y), 0, 0, 0);

                                        poly1.Elevation = poly2d.Elevation;



                                        Point3dCollection col_int = Functions.Intersect_on_both_operands(poly2d, poly1);
                                        if (col_int.Count > 0)
                                        {
                                            for (int i = 0; i < col_int.Count; ++i)
                                            {
                                                dt1.Rows.Add();
                                                dt1.Rows[dt1.Rows.Count - 1]["Type of object"] = "Line";
                                                dt1.Rows[dt1.Rows.Count - 1]["DWG Layer"] = line_int.Layer;
                                                dt1.Rows[dt1.Rows.Count - 1]["X"] = col_int[i].X;
                                                dt1.Rows[dt1.Rows.Count - 1]["Y"] = col_int[i].Y;

                                                Point3d pt1 = poly2d.GetClosestPointTo(col_int[i], Vector3d.ZAxis, false);
                                                double param1 = poly2d.GetParameterAtPoint(pt1);
                                                if (param1 > poly3d.EndParam) param1 = poly3d.EndParam;
                                                dt1.Rows[dt1.Rows.Count - 1]["Z on CL"] = poly3d.GetPointAtParameter(param1).Z;
                                                dt1.Rows[dt1.Rows.Count - 1]["Sta"] = poly3d.GetDistanceAtParameter(param1);


                                                System.Data.DataTable dt3 = new System.Data.DataTable();
                                                dt3.Columns.Add("X", typeof(double));
                                                dt3.Columns.Add("Y", typeof(double));
                                                dt3.Columns.Add("Z", typeof(double));


                                                dt3.Rows.Add();
                                                dt3.Rows[0]["X"] = line_int.StartPoint.X;
                                                dt3.Rows[0]["Y"] = line_int.StartPoint.Y;
                                                dt3.Rows[0]["Z"] = line_int.StartPoint.Z;
                                                dt3.Rows.Add();
                                                dt3.Rows[1]["X"] = line_int.EndPoint.X;
                                                dt3.Rows[1]["Y"] = line_int.EndPoint.Y;
                                                dt3.Rows[1]["Z"] = line_int.EndPoint.Z;



                                                Polyline3d poly_line_3d = Functions.Build_3d_poly_for_scanning(dt3);
                                                Point3d p2 = poly1.GetClosestPointTo(col_int[i], Vector3d.ZAxis, false);
                                                double param2 = poly1.GetParameterAtPoint(p2);
                                                if (param2 > poly_line_3d.EndParam) param2 = poly_line_3d.EndParam;

                                                double z_obj = poly_line_3d.GetPointAtParameter(param2).Z;

                                                dt1.Rows[dt1.Rows.Count - 1]["Z on object"] = z_obj;
                                                Functions.add_object_data_to_datatable(dt1, Tables1, line_int.ObjectId);
                                                poly_line_3d.Erase();
                                            }
                                        }
                                    }

                                    if (dbpt1 != null)
                                    {
                                        dt1.Rows.Add();
                                        dt1.Rows[dt1.Rows.Count - 1]["Type of object"] = "Point";
                                        dt1.Rows[dt1.Rows.Count - 1]["DWG Layer"] = dbpt1.Layer;
                                        dt1.Rows[dt1.Rows.Count - 1]["X"] = dbpt1.Position.X;
                                        dt1.Rows[dt1.Rows.Count - 1]["Y"] = dbpt1.Position.Y;
                                        dt1.Rows[dt1.Rows.Count - 1]["Z on object"] = dbpt1.Position.Z;
                                        Point3d pt1 = new Point3d();
                                        double param1 = -1;
                                        double x1 = 0;
                                        double y1 = 0;
                                        double x2 = 0;
                                        double y2 = 0;
                                        double x3 = 0;
                                        double y3 = 0;
                                        if (original_poly2d != null)
                                        {
                                            pt1 = original_poly2d.GetClosestPointTo(dbpt1.Position, Vector3d.ZAxis, false);
                                            dt1.Rows[dt1.Rows.Count - 1]["Sta"] = original_poly2d.GetDistAtPoint(pt1);
                                            dt1.Rows[dt1.Rows.Count - 1]["Z on CL"] = original_poly2d.Elevation;
                                            param1 = original_poly2d.GetParameterAtPoint(pt1);
                                            if (param1 + 1 <= original_poly2d.EndParam)
                                            {
                                                x1 = dbpt1.Position.X;
                                                y1 = dbpt1.Position.Y;
                                                x2 = pt1.X;
                                                y1 = pt1.Y;
                                                x3 = original_poly2d.GetPointAtParameter(param1 + 1).X;
                                                y3 = original_poly2d.GetPointAtParameter(param1 + 1).Y;
                                            }
                                            else
                                            {
                                                x3 = dbpt1.Position.X;
                                                y3 = dbpt1.Position.Y;
                                                x2 = pt1.X;
                                                y1 = pt1.Y;
                                                x1 = original_poly2d.GetPointAtParameter(param1 - 1).X;
                                                y1 = original_poly2d.GetPointAtParameter(param1 - 1).Y;
                                            }

                                        }
                                        else
                                        {
                                            pt1 = poly2d.GetClosestPointTo(dbpt1.Position, Vector3d.ZAxis, false);
                                            param1 = poly2d.GetParameterAtPoint(pt1);
                                            if (param1 + 1 <= poly2d.EndParam)
                                            {
                                                x1 = dbpt1.Position.X;
                                                y1 = dbpt1.Position.Y;
                                                x2 = pt1.X;
                                                y1 = pt1.Y;
                                                x3 = poly2d.GetPointAtParameter(param1 + 1).X;
                                                y3 = poly2d.GetPointAtParameter(param1 + 1).Y;
                                            }
                                            else
                                            {
                                                x3 = dbpt1.Position.X;
                                                y3 = dbpt1.Position.Y;
                                                x2 = pt1.X;
                                                y1 = pt1.Y;
                                                x1 = poly2d.GetPointAtParameter(param1 - 1).X;
                                                y1 = poly2d.GetPointAtParameter(param1 - 1).Y;
                                            }

                                            if (param1 > poly3d.EndParam) param1 = poly3d.EndParam;
                                            dt1.Rows[dt1.Rows.Count - 1]["Sta"] = poly3d.GetDistanceAtParameter(param1);
                                            dt1.Rows[dt1.Rows.Count - 1]["Z on CL"] = poly3d.GetPointAtParameter(param1).Z;


                                        }
                                        if (dt1.Columns.Contains("Offset") == false)
                                        {
                                            dt1.Columns.Add("Offset", typeof(double));
                                        }

                                        if (dt1.Columns.Contains("Side") == false)
                                        {
                                            dt1.Columns.Add("Side", typeof(string));
                                        }

                                        double dist = Math.Pow(Math.Pow(pt1.X - dbpt1.Position.X, 2) + Math.Pow(pt1.Y - dbpt1.Position.Y, 2), 0.5);


                                        string lr = Functions.Get_deflection_side(x1, y1, x2, y2, x3, y3);

                                        dt1.Rows[dt1.Rows.Count - 1]["Offset"] = dist;
                                        dt1.Rows[dt1.Rows.Count - 1]["Side"] = lr;


                                        Functions.add_object_data_to_datatable(dt1, Tables1, dbpt1.ObjectId);



                                    }

                                    if (block1 != null)
                                    {
                                        dt1.Rows.Add();
                                        dt1.Rows[dt1.Rows.Count - 1]["Type of object"] = "Block Reference";
                                        dt1.Rows[dt1.Rows.Count - 1]["Block Name"] = Functions.get_block_name(block1);
                                        dt1.Rows[dt1.Rows.Count - 1]["DWG Layer"] = block1.Layer;
                                        dt1.Rows[dt1.Rows.Count - 1]["X"] = block1.Position.X;
                                        dt1.Rows[dt1.Rows.Count - 1]["Y"] = block1.Position.Y;
                                        dt1.Rows[dt1.Rows.Count - 1]["Z on object"] = block1.Position.Z;
                                        Point3d pt1 = new Point3d();
                                        double param1 = -1;
                                        double x1 = 0;
                                        double y1 = 0;
                                        double x2 = 0;
                                        double y2 = 0;
                                        double x3 = 0;
                                        double y3 = 0;
                                        if (original_poly2d != null)
                                        {
                                            pt1 = original_poly2d.GetClosestPointTo(block1.Position, Vector3d.ZAxis, false);
                                            dt1.Rows[dt1.Rows.Count - 1]["Sta"] = original_poly2d.GetDistAtPoint(pt1);
                                            dt1.Rows[dt1.Rows.Count - 1]["Z on CL"] = original_poly2d.Elevation;
                                            param1 = original_poly2d.GetParameterAtPoint(pt1);
                                            if (param1 + 1 <= original_poly2d.EndParam)
                                            {
                                                x1 = block1.Position.X;
                                                y1 = block1.Position.Y;
                                                x2 = pt1.X;
                                                y1 = pt1.Y;
                                                x3 = original_poly2d.GetPointAtParameter(param1 + 1).X;
                                                y3 = original_poly2d.GetPointAtParameter(param1 + 1).Y;
                                            }
                                            else
                                            {
                                                x3 = block1.Position.X;
                                                y3 = block1.Position.Y;
                                                x2 = pt1.X;
                                                y1 = pt1.Y;
                                                x1 = original_poly2d.GetPointAtParameter(param1 - 1).X;
                                                y1 = original_poly2d.GetPointAtParameter(param1 - 1).Y;
                                            }

                                        }
                                        else
                                        {
                                            pt1 = poly2d.GetClosestPointTo(block1.Position, Vector3d.ZAxis, false);
                                            param1 = poly2d.GetParameterAtPoint(pt1);
                                            if (param1 + 1 <= poly2d.EndParam)
                                            {
                                                x1 = block1.Position.X;
                                                y1 = block1.Position.Y;
                                                x2 = pt1.X;
                                                y1 = pt1.Y;
                                                x3 = poly2d.GetPointAtParameter(param1 + 1).X;
                                                y3 = poly2d.GetPointAtParameter(param1 + 1).Y;
                                            }
                                            else
                                            {
                                                x3 = block1.Position.X;
                                                y3 = block1.Position.Y;
                                                x2 = pt1.X;
                                                y1 = pt1.Y;
                                                x1 = poly2d.GetPointAtParameter(param1 - 1).X;
                                                y1 = poly2d.GetPointAtParameter(param1 - 1).Y;
                                            }

                                            if (param1 > poly3d.EndParam) param1 = poly3d.EndParam;
                                            dt1.Rows[dt1.Rows.Count - 1]["Sta"] = poly3d.GetDistanceAtParameter(param1);
                                            dt1.Rows[dt1.Rows.Count - 1]["Z on CL"] = poly3d.GetPointAtParameter(param1).Z;


                                        }
                                        if (dt1.Columns.Contains("Offset") == false)
                                        {
                                            dt1.Columns.Add("Offset", typeof(double));
                                        }

                                        if (dt1.Columns.Contains("Side") == false)
                                        {
                                            dt1.Columns.Add("Side", typeof(string));
                                        }

                                        double dist = Math.Pow(Math.Pow(pt1.X - block1.Position.X, 2) + Math.Pow(pt1.Y - block1.Position.Y, 2), 0.5);


                                        string lr = Functions.Get_deflection_side(x1, y1, x2, y2, x3, y3);

                                        dt1.Rows[dt1.Rows.Count - 1]["Offset"] = dist;
                                        dt1.Rows[dt1.Rows.Count - 1]["Side"] = lr;

                                        if (block1.AttributeCollection.Count > 0)
                                        {
                                            foreach (ObjectId id2 in block1.AttributeCollection)
                                            {
                                                AttributeReference atr1 = Trans1.GetObject(id2, OpenMode.ForRead) as AttributeReference;
                                                if (atr1 != null)
                                                {
                                                    if (dt1.Columns.Contains("BlockAttribute: " + atr1.Tag) == false)
                                                    {
                                                        dt1.Columns.Add("BlockAttribute: " + atr1.Tag, typeof(string));
                                                    }
                                                    dt1.Rows[dt1.Rows.Count - 1]["BlockAttribute: " + atr1.Tag] = atr1.TextString;
                                                }
                                            }
                                        }


                                        Functions.add_object_data_to_datatable(dt1, Tables1, block1.ObjectId);
                                    }

                                }



                            }




                        }


                        Functions.Transfer_datatable_to_new_excel_spreadsheet(dt1, Convert.ToString(DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + "_" + DateTime.Now.Hour + "hr" + DateTime.Now.Minute + "min" + DateTime.Now.Second) + "sec");



                        if (delete_poly3d == true)
                        {
                            poly3d.Erase();
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
        }


        [CommandMethod("scan_cover")]
        public void scan_cover()
        {

            Editor editor1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;

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

                        Autodesk.AutoCAD.DatabaseServices.LayerTable layer_table = (Autodesk.AutoCAD.DatabaseServices.LayerTable)Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

                        Matrix3d CurentUCSmatrix = Editor1.CurrentUserCoordinateSystem;

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                        Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the reference polyline:");
                        Prompt_centerline.SetRejectMessage("\nSelect a polyline!");
                        Prompt_centerline.AllowNone = true;
                        Prompt_centerline.AddAllowedClass(typeof(Polyline), false);
                        Rezultat_centerline = ThisDrawing.Editor.GetEntity(Prompt_centerline);

                        if (Rezultat_centerline.Status != PromptStatus.OK)
                        {
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Polyline poly2d = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForRead) as Polyline;

                        System.Data.DataTable dt1 = new System.Data.DataTable();
                        dt1.Columns.Add("Type of object", typeof(string));
                        dt1.Columns.Add("BlockName", typeof(string));
                        dt1.Columns.Add("Layer", typeof(string));
                        dt1.Columns.Add("X", typeof(double));
                        dt1.Columns.Add("Y", typeof(double));
                        dt1.Columns.Add("Distance", typeof(double));


                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez.MessageForAdding = "\nSelect the objects:";
                        Prompt_rez.SingleOnly = false;
                        Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        for (int i = 0; i < Rezultat1.Value.Count; i++)
                        {
                            BlockReference block1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as BlockReference;

                            if (block1 != null)
                            {
                                double x1 = block1.Position.X;
                                double y1 = block1.Position.Y;
                                {
                                    dt1.Rows.Add();
                                    dt1.Rows[dt1.Rows.Count - 1]["Type of object"] = "BLock reference";
                                    dt1.Rows[dt1.Rows.Count - 1]["Layer"] = block1.Layer;
                                    dt1.Rows[dt1.Rows.Count - 1]["X"] = x1;
                                    dt1.Rows[dt1.Rows.Count - 1]["Y"] = y1;
                                    dt1.Rows[dt1.Rows.Count - 1]["BlockName"] = Functions.get_block_name(block1);

                                    Xline xline1 = new Xline();
                                    xline1.BasePoint = new Point3d(x1, y1, poly2d.Elevation);
                                    xline1.SecondPoint = new Point3d(x1, y1 + 10, poly2d.Elevation);

                                    Point3dCollection col_int = Functions.Intersect_on_both_operands(xline1, poly2d);

                                    if (col_int.Count > 0)
                                    {
                                        Point3d pt1 = col_int[0];
                                        dt1.Rows[dt1.Rows.Count - 1]["Distance"] = pt1.Y - y1;
                                    }

                                    if (block1.AttributeCollection.Count > 0)
                                    {
                                        foreach (ObjectId id1 in block1.AttributeCollection)
                                        {
                                            AttributeReference atr1 = Trans1.GetObject(id1, OpenMode.ForRead) as AttributeReference;
                                            if (atr1 != null)
                                            {
                                                if (dt1.Columns.Contains("BlockAttribute: " + atr1.Tag) == false)
                                                {
                                                    dt1.Columns.Add("BlockAttribute: " + atr1.Tag, typeof(string));
                                                }
                                                dt1.Rows[dt1.Rows.Count - 1]["BlockAttribute: " + atr1.Tag] = atr1.TextString;
                                            }
                                        }
                                    }

                                    Functions.add_object_data_to_datatable(dt1, Tables1, block1.ObjectId);
                                }
                            }
                        }
                        Functions.Transfer_datatable_to_new_excel_spreadsheet(dt1, Convert.ToString(DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + "_" + DateTime.Now.Hour + "hr" + DateTime.Now.Minute + "min" + DateTime.Now.Second) + "sec");
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

    }
}
