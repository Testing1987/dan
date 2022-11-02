#region Namespaces

using System;
using System.Text;
using System.Linq;
using System.Xml;
using System.Reflection;
using System.ComponentModel;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows;
using System.Windows.Forms;
using System.Drawing;
using System.IO;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Windows;

using MgdAcApplication = Autodesk.AutoCAD.ApplicationServices.Application;
using MgdAcDocument = Autodesk.AutoCAD.ApplicationServices.Document;
using AcWindowsNS = Autodesk.AutoCAD.Windows;

#endregion

namespace NESW1111
{
    public class Mtextjig_X : EntityJig
    {
        #region Fields

        public int mCurJigFactorIndex = 1;  // Jig Factor Index

        public Autodesk.AutoCAD.Geometry.Point3d mArrowLocation; // Jig Factor #1
        public Autodesk.AutoCAD.Geometry.Point3d mTextLocation; // Jig Factor #2
        public string _content; // Jig Factor #3
        double _y = 0;

        #endregion

        #region Constructors

        public Mtextjig_X(MText mtext1, string content1, double rot1, double height1, AttachmentPoint align1, double y) : base(mtext1)
        {
            Mt.SetDatabaseDefaults();

            Mt.Contents = content1;
            Mt.Rotation = rot1;
            Mt.TextHeight = height1;
            Mt.Attachment = align1;
            _y = y;
        }

        #endregion

        #region Properties

        private Editor Editor
        {
            get
            {
                return MgdAcApplication.DocumentManager.MdiActiveDocument.Editor;
            }
        }

        private Matrix3d UCS
        {
            get
            {
                return MgdAcApplication.DocumentManager.MdiActiveDocument.Editor.CurrentUserCoordinateSystem;
            }
        }

        #endregion

        #region Overrides

        public new MText Mt  // Overload the Entity property for convenience.
        {
            get
            {
                return base.Entity as MText;
            }
        }

        protected override bool Update()
        {

            switch (mCurJigFactorIndex)
            {
                case 1:
                    Mt.Location = new Point3d(mArrowLocation.X, _y, 0);
                    break;
                default:
                    return false;
            }

            return true;
        }

        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            switch (mCurJigFactorIndex)
            {
                case 1:
                    JigPromptPointOptions prOptions1 = new JigPromptPointOptions("\nArrow Location:");
                    // Set properties such as UseBasePoint and BasePoint of the prompt options object if necessary here.
                    prOptions1.UserInputControls = UserInputControls.Accept3dCoordinates | UserInputControls.GovernedByOrthoMode | UserInputControls.GovernedByUCSDetect | UserInputControls.UseBasePointElevation;
                    PromptPointResult prResult1 = prompts.AcquirePoint(prOptions1);
                    if (prResult1.Status == PromptStatus.Cancel && prResult1.Status == PromptStatus.Error)
                        return SamplerStatus.Cancel;

                    if (prResult1.Value.Equals(mArrowLocation))  //Use better comparison method if necessary.
                    {
                        return SamplerStatus.NoChange;
                    }
                    else
                    {
                        mArrowLocation = new Point3d(prResult1.Value.X, -_y, 0);
                        return SamplerStatus.OK;
                    }


                default:
                    break;
            }

            return SamplerStatus.OK;
        }



        #endregion

        #region Methods to Call

        public static MText Jig(string content1, double rot1, double height1, AttachmentPoint align1, double y)
        {
            Mtextjig_X jigger = null;
            try
            {
                jigger = new Mtextjig_X(new MText(), content1, rot1, height1, align1, y);
                PromptResult pr;
                do
                {
                    pr = MgdAcApplication.DocumentManager.MdiActiveDocument.Editor.Drag(jigger);

                    jigger.mCurJigFactorIndex++;


                } while (pr.Status != PromptStatus.Cancel && pr.Status != PromptStatus.Error && jigger.mCurJigFactorIndex <= 1);

                if (pr.Status == PromptStatus.Cancel || pr.Status == PromptStatus.Error)
                {
                    if (jigger != null && jigger.Mt != null) jigger.Mt.Dispose();
                    return null;
                }
                else
                {
                    return jigger.Mt;
                }
            }
            catch
            {
                if (jigger != null && jigger.Mt != null) jigger.Mt.Dispose();
                return null;
            }
        }

        #endregion

        #region Test Commands

        [CommandMethod("zz")]
        public static void TestMLeaderJigger_Method()
        {
            try
            {
                Entity new_mtext = Mtextjig_X.Jig("345", 0, 0.1, AttachmentPoint.BottomLeft, 10);
                if (new_mtext != null)
                {
                    Database db = HostApplicationServices.WorkingDatabase;
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                        btr.AppendEntity(new_mtext);
                        tr.AddNewlyCreatedDBObject(new_mtext, true);
                        tr.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MgdAcApplication.DocumentManager.MdiActiveDocument.Editor.WriteMessage(ex.ToString());
            }
        }

        #endregion
    }
}
