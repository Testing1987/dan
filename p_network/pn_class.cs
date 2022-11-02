using Autodesk.AutoCAD.EditorInput;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.Settings;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace p_network
{
    public class pn_class
    {
        [Autodesk.AutoCAD.Runtime.CommandMethod("PN_LIST")]
        public void pn_LIST()
        {
            CivilDocument doc1 = CivilApplication.ActiveDocument;
            Editor editor1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            SettingsPipeNetwork set1 = doc1.Settings.GetSettings<SettingsPipeNetwork>() as SettingsPipeNetwork;



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


                        //Vous devez ajouter une référence à AecBaseMgd.dll(dans le répertoire d'installation).
                        ObjectId dsvid = Autodesk.Aec.ApplicationServices.DrawingSetupVariables.GetInstance(ThisDrawing.Database, false);
                        Autodesk.Aec.ApplicationServices.DrawingSetupVariables dsv = Trans1.GetObject(dsvid, OpenMode.ForRead) as Autodesk.Aec.ApplicationServices.DrawingSetupVariables;


                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                        Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the C3D pipe:");
                        Prompt_centerline.SetRejectMessage("\nSelect a polyline!");
                        Prompt_centerline.AllowNone = true;
                        Prompt_centerline.AddAllowedClass(typeof(Autodesk.Civil.DatabaseServices.Pipe), false);
                        Rezultat_centerline = ThisDrawing.Editor.GetEntity(Prompt_centerline);

                        if (Rezultat_centerline.Status != PromptStatus.OK)
                        {
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Autodesk.Civil.DatabaseServices.Pipe pipe1 = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Pipe;

                        if (pipe1 != null)
                        {
                            Point3d pt1 = pipe1.StartPoint;
                            Point3d pt2 = pipe1.EndPoint;

                            string msg = "_dwgunits linear = " + dsv.LinearUnit.ToString() +"\r\n" +
                                             "Start point = " +   pt1.ToString() + "\r\n" +
                                             "End point = " + pt2.ToString() + "\r\n" +
                                             "inner diam = " + pipe1.InnerDiameterOrWidth + "\r\n" +
                                             "outer diam = " + pipe1.OuterDiameterOrWidth + "\r\n" +
                                             "style name = " + pipe1.StyleName + "\r\n" +
                                             "description = " + pipe1.Description;
                            MessageBox.Show(msg);

                            //my tests indicate is z minus half of inner diameter gets you to the bop...... why is not outer diam?
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

        [Autodesk.AutoCAD.Runtime.CommandMethod("PN_ID")]
        public void pn_ID()
        {
            CivilDocument doc1 = CivilApplication.ActiveDocument;
            Editor editor1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            SettingsPipeNetwork set1 = doc1.Settings.GetSettings<SettingsPipeNetwork>() as SettingsPipeNetwork;



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



                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                        PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\npoint:");
                        PP1.AllowNone = false;
                        Point_res1 = Editor1.GetPoint(PP1);

                        if (Point_res1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Point3d pt1 = Point_res1.Value;
                        Matrix3d CurentUCSmatrix = Editor1.CurrentUserCoordinateSystem;
                        Matrix3d CurentUCSmatrix1 = Editor1.CurrentUserCoordinateSystem.Inverse();

                        Point3d pt2 = pt1.TransformBy(CurentUCSmatrix);
                        Point3d pt3 = pt1.TransformBy(CurentUCSmatrix1);

                        MessageBox.Show(pt1.ToString() + "\r\n" + pt2.ToString() + "\r\n" + pt3.ToString());

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
