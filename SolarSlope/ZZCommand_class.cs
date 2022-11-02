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
using Autodesk.Civil.DatabaseServices;

namespace Alignment_mdi
{
    public class ZZCommand_class
    {
        public bool isSECURE()
        {

            string number_drive = GetHDDSerialNumber("C");

            switch (number_drive)
            {
                case "8CDA6CE3":
                    return true;
                case "36D79DE5":
                    return true;
                case "FEA3192C":
                    return true;
                case "B454BD5B":
                    return true;
                case "6E40460D":
                    return true;
                case "0892E01D":
                    return true;
                case "4ED21ABF":
                    return true;
                case "56766C69":
                    return true;
                case "DA214366":
                    return true;
                case "3CF68AF2":
                    return true;
                case "389A2249":
                    return true;
                case "AED6B68E":
                    return true;
                case "8C040338":
                    return true;
                case "8CD08F48":
                    return true;
                case "0E26E402":
                    return true;
                case "4A123A50":
                    return true;

                case "98D9B617":
                    return true;
                case "B838FEB4":
                    return true;
                case "1AE1721C":
                    return true;
                case "CA9E6FFE":
                    return true;
                case "DE281128":
                    return true;
                case "FC7C4F1":
                    return true;
                case "B67EC134":
                    return true;
                case "E64DBF0A":
                    return true;
                case "561F1509":
                    return true;

                case "120E4B54":
                    return true;
                case "F6633173":
                    return true;
                case "40D6BDCB":
                    return true;
                case "18399D24":
                    return true;


                case "B63AD3F6":
                    return true;
                default:
                    try
                    {
                        string UserDNS = Environment.GetEnvironmentVariable("USERDNSDOMAIN");
                        if (UserDNS.ToUpper() == "HMMG.CC" | UserDNS.ToLower() == "mottmac.group.int")
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


        public string GetHDDSerialNumber(string drive)
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


        //[CommandMethod("SLOPEZ")]
        public void Show_wgen_form()
        {
            if (isSECURE() == true)
            {
                foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
                {
                    if (Forma1 is Alignment_mdi.Solar_main_form)
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
                    Alignment_mdi.Solar_main_form forma2 = new Alignment_mdi.Solar_main_form();
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


        [CommandMethod("profiler")]
        public void Show_profiler_form()
        {
            if (isSECURE() == true)
            {
                foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
                {
                    if (Forma1 is Alignment_mdi.Profiler_main)
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
                    Alignment_mdi.Profiler_main forma2 = new Alignment_mdi.Profiler_main();
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

        [CommandMethod("zid")]
        public void alignment_id()
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
                        BlockTableRecord BTrecord_MS = Trans1.GetObject(BlockTable1[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;
                        BlockTableRecord BTrecord_PS = Trans1.GetObject(BlockTable1[BlockTableRecord.PaperSpace], OpenMode.ForRead) as BlockTableRecord;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;
                        LayerTable LayerTable1 = Trans1.GetObject(ThisDrawing.Database.LayerTableId, OpenMode.ForRead) as LayerTable;
                        TextStyleTable Text_style_table1 = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;


                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                        Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the alignment:");
                        Prompt_centerline.SetRejectMessage("\nSelect an alignment!");
                        Prompt_centerline.AllowNone = true;
                        Prompt_centerline.AddAllowedClass(typeof(Autodesk.Civil.DatabaseServices.Alignment), false);
                        Rezultat_centerline = ThisDrawing.Editor.GetEntity(Prompt_centerline);

                        if (Rezultat_centerline.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Alignment al1 = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForRead) as Alignment;
                        if (al1 != null)
                        {
                            ThisDrawing.Editor.WriteMessage("\n" + al1.StyleName);
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


        [CommandMethod("zzz")]
        public void profile_3000()
        {
            if (isSECURE() == true)
            {

                ObjectId[] Empty_array = null;
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.Civil.ApplicationServices.CivilDocument CivilDrawing = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument;

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

                            Point3d ptins = new Point3d(2958776.6390, 10598887.8716, 0);
                            LayerTable LayerTable1 = Trans1.GetObject(ThisDrawing.Database.LayerTableId, OpenMode.ForRead) as LayerTable;
                            string layer1 = "C-PV-Grade-PROF_Design";
                            string layer2 = "C-PV-Grade-PROF_EG";

                            Functions.Creaza_layer(layer1, 1, true);
                            Functions.Creaza_layer(layer2, 2, true);

                            ObjectIdCollection col1 = CivilDrawing.GetAlignmentIds();

                            if (col1.Count > 0)
                            {
                                for (int i = 0; i < col1.Count; i++)
                                {
                                    Alignment align1 = Trans1.GetObject(col1[i], OpenMode.ForRead) as Alignment;



                                    if (align1 != null && align1.Name == "test1")
                                    {
                                        ObjectId layerid1 = LayerTable1[layer1];
                                        ObjectId layerid2 = LayerTable1[layer2];

                                        ObjectId styleid1 = CivilDrawing.Styles.ProfileStyles["Design Profile"];
                                        ObjectId styleid2 = CivilDrawing.Styles.ProfileStyles["Existing Ground Profile"];

                                        ObjectId labelsetid = CivilDrawing.Styles.LabelSetStyles.ProfileLabelSetStyles[0];

                                        ObjectIdCollection col_surf = CivilDrawing.GetSurfaceIds();

                                        ObjectId surfaceid1 = ObjectId.Null;
                                        ObjectId surfaceid2 = ObjectId.Null;
                                        string surf1_name = "";
                                        string surf2_name = "";

                                        for (int j = 0; j < col_surf.Count; ++j)
                                        {
                                            Autodesk.Civil.DatabaseServices.Surface surf1 = Trans1.GetObject(col_surf[j], OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Surface;
                                            if (surf1 != null)
                                            {
                                                if (surf1.Name == "Grizzly Ridge Surface")
                                                {
                                                    surfaceid2 = col_surf[j];
                                                    surf1_name = surf1.Name.Replace(" ", "_");
                                                }
                                                if (surf1.Name == "EG_Rough_Grade")
                                                {
                                                    surfaceid1 = col_surf[j];
                                                    surf2_name = surf1.Name.Replace(" ", "_");
                                                }
                                            }
                                        }



                                        if (surfaceid1 != ObjectId.Null)
                                        {
                                            ObjectId profid1 = Profile.CreateFromSurface("prof_surf_" + surf1_name + "_align_" + align1.Name, align1.ObjectId, surfaceid1, layerid1, styleid1, labelsetid);
                                        }

                                        if (surfaceid2 != ObjectId.Null)
                                        {
                                            ObjectId profid2 = Profile.CreateFromSurface("prof_surf_" + surf2_name + "_align_" + align1.Name, align1.ObjectId, surfaceid2, layerid2, styleid2, labelsetid);
                                        }

                                        ThisDrawing.Editor.WriteMessage("\nprofile view band set styles:");

                                        Autodesk.Civil.DatabaseServices.Styles.ProfileViewBandSetStyleCollection col2 = CivilDrawing.Styles.ProfileViewBandSetStyles;
                                        ObjectId profviewsbandtyle = col2[0];
                                        IEnumerator<ObjectId> enum2 = col2.GetEnumerator();

                                        while (enum2.MoveNext())
                                        {
                                            Autodesk.Civil.DatabaseServices.Styles.ProfileViewBandSetStyle tst2 = Trans1.GetObject(enum2.Current, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.ProfileViewBandSetStyle;

                                            if (tst2.Name == "_No Bands")
                                            {
                                                profviewsbandtyle = tst2.ObjectId;
                                            }
                                            ThisDrawing.Editor.WriteMessage("\n" + tst2.Name);
                                        }


                                        ThisDrawing.Editor.WriteMessage("\nprofile view styles:");

                                        Autodesk.Civil.DatabaseServices.Styles.ProfileViewStyleCollection col3 = CivilDrawing.Styles.ProfileViewStyles;
                                        ObjectId profviewstyle = CivilDrawing.Styles.ProfileViewStyles[0];
                                        IEnumerator<ObjectId> enum3 = col3.GetEnumerator();

                                        double vexag = 1;

                                        while (enum3.MoveNext())
                                        {
                                            Autodesk.Civil.DatabaseServices.Styles.ProfileViewStyle tst3 = Trans1.GetObject(enum3.Current, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.ProfileViewStyle;
                                            ThisDrawing.Editor.WriteMessage("\n" + tst3.Name);
                                            if (tst3.Name == "MM_Sample")
                                            {
                                                profviewstyle = tst3.ObjectId;
                                                Autodesk.Civil.DatabaseServices.Styles.GraphStyle graphstyle3 = tst3.GraphStyle;
                                                vexag = graphstyle3.VerticalExaggeration;
                                            }
                                        }




                                        ObjectId profview1 = ProfileView.Create(align1.ObjectId, ptins, "pview_" + align1.Name, profviewsbandtyle, profviewstyle);

                                        ProfileView pview1 = Trans1.GetObject(profview1, OpenMode.ForRead) as ProfileView;



                                        if (pview1 != null)
                                        {

                                            Point3d pt1 = pview1.StartPoint;

                                            Extents3d ext1 = pview1.GeometricExtents;
                                            double height1 = ext1.MaxPoint.Y - ext1.MinPoint.Y;

                                            ptins = new Point3d(ptins.X, ptins.Y + height1 + 100, 0);
                                            ThisDrawing.Editor.WriteMessage("\nIns=" + pt1.ToString());
                                            ThisDrawing.Editor.WriteMessage("\nH=" + height1.ToString());
                                        }
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
            else
            {
                return;
            }



        }


        [CommandMethod("TOP2XL1")]
        public void SELECT_MLEADER_COGOX2()
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
                    dt1.Columns.Add("Station", typeof(string));
                    dt1.Columns.Add("Ground", typeof(double));
                    dt1.Columns.Add("Top of Pipe", typeof(double));
                    bool run1 = true;
                    do
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                            BlockTableRecord BTrecord_MS = Trans1.GetObject(BlockTable1[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;
                            BlockTableRecord BTrecord_PS = Trans1.GetObject(BlockTable1[BlockTableRecord.PaperSpace], OpenMode.ForRead) as BlockTableRecord;
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;



                            Autodesk.AutoCAD.EditorInput.PromptEntityResult rez_ml1;
                            Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_ml;
                            Prompt_ml = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the mleader with station:");
                            Prompt_ml.SetRejectMessage("\nSelect a mleader!");
                            Prompt_ml.AllowNone = true;
                            Prompt_ml.AddAllowedClass(typeof(MLeader), false);
                            rez_ml1 = ThisDrawing.Editor.GetEntity(Prompt_ml);

                            if (rez_ml1.Status != PromptStatus.OK)
                            {

                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                run1 = false;
                                Trans1.Commit();
                            }
                            else
                            {
                                MLeader ml1 = Trans1.GetObject(rez_ml1.ObjectId, OpenMode.ForWrite) as MLeader;
                                if (ml1 != null)
                                {
                                    dt1.Rows.Add();
                                    dt1.Rows[dt1.Rows.Count - 1][0] = ml1.MText.Contents.Replace("+", "");

                                    ml1.ColorIndex = 1;

                                    Editor1.WriteMessage("\n" + ml1.MText.Contents);

                                    Autodesk.AutoCAD.EditorInput.PromptEntityResult rez_ground;
                                    Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_ground;
                                    Prompt_ground = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the ground cogopoint:");
                                    Prompt_ground.SetRejectMessage("\nSelect a cogopoint!");
                                    Prompt_ground.AllowNone = true;
                                    Prompt_ground.AddAllowedClass(typeof(CogoPoint), false);
                                    rez_ground = ThisDrawing.Editor.GetEntity(Prompt_ground);
                                    if (rez_ground.Status == PromptStatus.OK)
                                    {
                                        CogoPoint cogo1 = Trans1.GetObject(rez_ground.ObjectId, OpenMode.ForRead) as CogoPoint;
                                        if (cogo1 != null)
                                        {
                                            dt1.Rows[dt1.Rows.Count - 1][1] = cogo1.Elevation;
                                            Editor1.WriteMessage("\n" + Convert.ToString(cogo1.Elevation));
                                        }

                                    }

                                    Autodesk.AutoCAD.EditorInput.PromptEntityResult rez_top;
                                    Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_top;
                                    Prompt_top = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the TOP cogopoint:");
                                    Prompt_top.SetRejectMessage("\nSelect a cogopoint!");
                                    Prompt_top.AllowNone = true;
                                    Prompt_top.AddAllowedClass(typeof(CogoPoint), false);
                                    rez_top = ThisDrawing.Editor.GetEntity(Prompt_top);
                                    if (rez_top.Status == PromptStatus.OK)
                                    {
                                        CogoPoint cogo1 = Trans1.GetObject(rez_top.ObjectId, OpenMode.ForRead) as CogoPoint;
                                        if (cogo1 != null)
                                        {
                                            dt1.Rows[dt1.Rows.Count - 1][2] = cogo1.Elevation;
                                            Editor1.WriteMessage("\n" + Convert.ToString(cogo1.Elevation));
                                        }

                                    }

                                }
                                else
                                {
                                    run1 = false;
                                    Trans1.Commit();
                                }
                            }

                            if (run1 == true) Trans1.Commit();


                        }
                    } while (run1 == true);

                    Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt1);
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
