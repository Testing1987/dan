using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;
using System.Collections.Generic;

using Autodesk.AutoCAD.Runtime;

using System;
using System.Text;
using System.Linq;
using System.Xml;
using System.Reflection;
using System.ComponentModel;
using System.Collections;

using System.Diagnostics;
using System.Windows;
using System.Windows.Forms;
using System.Drawing;
using System.IO;
using Autodesk.AutoCAD.Windows;

using MgdAcApplication = Autodesk.AutoCAD.ApplicationServices.Application;
using MgdAcDocument = Autodesk.AutoCAD.ApplicationServices.Document;
using AcWindowsNS = Autodesk.AutoCAD.Windows;

namespace NESW
{
    public class DragMove
    {
        private Document _dwg;
        private SelectionSet _sset;

        private Point3d _basePoint = new Point3d(0.0, 0.0, 0.0);
        private bool _useBasePoint = false;
        private Line _rubberLine = null;

        public DragMove(Document dwg, SelectionSet ss)
        {
            _dwg = dwg;
            _sset = ss;
        }

        #region public properties

        public bool UseBasePoint
        {
            set { _useBasePoint = value; }
            get { return _useBasePoint; }
        }

        public Point3d BasePoint
        {
            set { _basePoint = value; }
            get { return _basePoint; }
        }

        #endregion

        #region public methods

        public void DoDrag()
        {
            if (_sset.Count == 0) return;

            _rubberLine = null;

            using (Transaction tran =
                _dwg.Database.TransactionManager.StartTransaction())
            {
                //Highlight entities
                SetHighlight(true, tran);

                if (!_useBasePoint)
                {
                    _basePoint = GetDefaultBasePoint();
                }
                else
                {
                    Point3d pt;
                    if (!PickBasePoint(_dwg.Editor, out pt)) return;

                    _basePoint = pt;
                    _useBasePoint = true;
                }

                PromptPointResult ptRes = _dwg.Editor.Drag
                    (_sset, "\nPick point to move to: ",
                        delegate (Point3d pt, ref Matrix3d mat)
                        {
                            if (pt == _basePoint)
                            {
                                return SamplerStatus.NoChange;
                            }
                            else
                            {
                                if (_useBasePoint)
                                {
                                    if (_rubberLine == null)
                                    {
                                        _rubberLine = new Line(_basePoint, pt);
                                        _rubberLine.SetDatabaseDefaults(_dwg.Database);

                                        IntegerCollection intCol;

                                        //Create transient graphics: rubberband line
                                        intCol = new IntegerCollection();
                                        TransientManager.CurrentTransientManager.
                                            AddTransient(_rubberLine,
                                            TransientDrawingMode.DirectShortTerm, 128, intCol);
                                    }
                                    else
                                    {
                                        _rubberLine.EndPoint = pt;

                                        //Update the transient graphics
                                        IntegerCollection intCol = new IntegerCollection();
                                        TransientManager.CurrentTransientManager.
                                            UpdateTransient(_rubberLine, intCol);
                                    }
                                }

                                mat = Matrix3d.Displacement(_basePoint.GetVectorTo(pt));
                                return SamplerStatus.OK;
                            }
                        }
                    );

                if (_rubberLine != null)
                {
                    //Clear transient graphics
                    IntegerCollection intCol = new IntegerCollection();
                    TransientManager.CurrentTransientManager.
                        EraseTransient(_rubberLine, intCol);

                    _rubberLine.Dispose();
                    _rubberLine = null;
                }

                if (ptRes.Status == PromptStatus.OK)
                {
                    MoveObjects(ptRes.Value, tran);
                }

                //Unhighlight entities
                SetHighlight(false, tran);

                tran.Commit();
            }
        }

        public void DoDrag(bool pickBasePt)
        {
            _useBasePoint = pickBasePt;

            DoDrag();
        }

        public void DoDrag(Point3d basePt)
        {
            _basePoint = basePt;
            _useBasePoint = true;

            DoDrag();
        }

        #endregion

        #region private methods

        private Point3d GetDefaultBasePoint()
        {
            Point3d pt = new Point3d();

            Extents3d exts = new Extents3d(
                new Point3d(0.0, 0.0, 0.0), new Point3d(0.0, 0.0, 0.0));

            using (Transaction tran =
                _dwg.Database.TransactionManager.StartTransaction())
            {
                ObjectId[] ids = _sset.GetObjectIds();
                for (int i = 0; i < ids.Length; i++)
                {
                    ObjectId id = ids[i];
                    Entity ent = (Entity)tran.GetObject(id, OpenMode.ForRead);

                    if (i == 0)
                    {
                        exts = ent.GeometricExtents;
                    }
                    else
                    {
                        Extents3d ext = ent.GeometricExtents;
                        exts.AddExtents(ext);
                    }
                }

                tran.Commit();
            }

            return exts.MinPoint;
        }

        private bool PickBasePoint(Editor ed, out Point3d pt)
        {
            pt = new Point3d();

            PromptPointOptions opt = new PromptPointOptions("Pick base point: ");
            PromptPointResult res = ed.GetPoint(opt);
            if (res.Status == PromptStatus.OK)
            {
                pt = res.Value;
                return true;
            }
            else
            {
                return false;
            }
        }

        private void MoveObjects(Point3d pt, Transaction tran)
        {
            Matrix3d mat = Matrix3d.Displacement(_basePoint.GetVectorTo(pt));
            foreach (ObjectId id in _sset.GetObjectIds())
            {
                Entity ent = (Entity)tran.GetObject(id, OpenMode.ForWrite);
                ent.TransformBy(mat);
            }
        }

        private void SetHighlight(bool highlight, Transaction tran)
        {
            foreach (ObjectId id in _sset.GetObjectIds())
            {
                Entity ent = (Entity)tran.GetObject(id, OpenMode.ForWrite);

                if (highlight)
                    ent.Highlight();
                else
                    ent.Unhighlight();
            }
        }

        #endregion
    }




}


namespace NESW
{
    public class MyCommand
    {
        [CommandMethod("DragMove", CommandFlags.UsePickSet)]
        public static void DoMove()
        {
            Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.
                Application.DocumentManager.MdiActiveDocument;

            Editor ed = ThisDrawing.Editor;

            //Get PickFirst SelectionSet
            PromptSelectionResult setRes = ed.SelectImplied();

            if (setRes.Status != PromptStatus.OK)
            {
                //if not PickFirst set, ask user to pick:
                ObjectId[] ids = GetUserPickedObjects(ThisDrawing);
                if (ids.Length == 0)
                {
                    ed.WriteMessage("\n*Cancelled*");
                    ed.WriteMessage("\n*Cancelled*");
                    return;
                }

                ed.SetImpliedSelection(ids);
                setRes = ed.SelectImplied();

                if (setRes.Status != PromptStatus.OK)
                {
                    ed.Regen();
                    return;
                }
            }

            //Do the dragging with the selectionSet
            DragMove drag = new DragMove(ThisDrawing, setRes.Value);
            drag.DoDrag(true);
        }

        private static ObjectId[] GetUserPickedObjects(Document dwg)
        {
            List<ObjectId> ids = new List<ObjectId>();
            using (Transaction tran =
                dwg.Database.TransactionManager.StartTransaction())
            {
                bool go = true;
                while (go)
                {
                    go = false;

                    PromptEntityOptions opt =
                        new PromptEntityOptions("\nPick an entity: ");
                    PromptEntityResult res = dwg.Editor.GetEntity(opt);

                    if (res.Status == PromptStatus.OK)
                    {
                        bool exists = false;
                        foreach (ObjectId id in ids)
                        {
                            if (id == res.ObjectId)
                            {
                                exists = true;
                                break;
                            }
                        }

                        if (!exists)
                        {
                            //Highlight
                            Entity ent = (Entity)tran.GetObject(
                                res.ObjectId, OpenMode.ForWrite);

                            ent.Highlight();

                            ids.Add(res.ObjectId);
                            go = true;
                        }
                    }
                }

                tran.Commit();
            }

            return ids.ToArray();
        }
    }
}

namespace NESW
{
    public class Move_mtext_Jig_X : EntityJig
    {
        Point3d newlocation;
        Point3d oldlocation;

        public Move_mtext_Jig_X(MText mtext1)
          : base(mtext1.Clone() as MText)
        {
            newlocation = mtext1.Location;
            oldlocation = mtext1.Location;
        }

        protected override bool Update()
        {
            // adjust the entity on screen
            ((MText)Entity).Location = new Point3d(newlocation.X, oldlocation.Y, 0);
            return true;
        }

        protected override SamplerStatus Sampler(
          JigPrompts prompts)
        {
            // get the new location (on cursor move)
            JigPromptPointOptions jppo =
              new JigPromptPointOptions("Specify X position: ");
            PromptPointResult pdr = prompts.AcquirePoint(jppo);

            // check if is OK
            if (pdr.Status != PromptStatus.OK)
                return SamplerStatus.Cancel;
            if (pdr.Value.DistanceTo(newlocation)
              < Tolerance.Global.EqualPoint)
                return SamplerStatus.NoChange;

            // and store the new valuye
            newlocation = new Point3d(pdr.Value.X, oldlocation.Y, 0);
            return SamplerStatus.OK;
        }

        /// <summary>
        /// Return the new location, during/after JIG
        /// </summary>
        public Point3d Location { get { return newlocation; } }

        [CommandMethod("moveMTXT")]
        public static void CmdMoveLabel()
        {
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
              .MdiActiveDocument.Editor;

            // select the label
            PromptEntityOptions peo =
              new PromptEntityOptions("\nSelect a label to move: ");
            peo.SetRejectMessage("\nLabels only");
            peo.AddAllowedClass(typeof(MText), false);
            PromptEntityResult per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            // start a transaction
            Database db = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
              .MdiActiveDocument.Database;
            using (Transaction trans = db.
              TransactionManager.StartTransaction())
            {
                // open the label object
                MText lbl = trans.GetObject(per.ObjectId, OpenMode.ForRead) as MText;

                // start the JIG drag
                Move_mtext_Jig_X jig_X = new Move_mtext_Jig_X(lbl);
                PromptResult pr = ed.Drag(jig_X);

                // if not OK, clean return
                if (pr.Status != PromptStatus.OK) return;

                // everything ok, update the location
                lbl.UpgradeOpen();
                lbl.Location = jig_X.Location;

                trans.Commit();
            }
        }
    }


    public class Move_mtext_Jig_Y : EntityJig
    {
        Point3d newlocation;
        Point3d oldlocation;

        public Move_mtext_Jig_Y(MText mtext1)
          : base(mtext1.Clone() as MText)
        {
            newlocation = mtext1.Location;
            oldlocation = mtext1.Location;
        }

        protected override bool Update()
        {
            // adjust the entity on screen
            ((MText)Entity).Location = new Point3d(oldlocation.X, newlocation.Y, 0);
            return true;
        }

        protected override SamplerStatus Sampler(
          JigPrompts prompts)
        {
            // get the new location (on cursor move)
            JigPromptPointOptions jppo =
              new JigPromptPointOptions("Specify Y position: ");
            PromptPointResult pdr = prompts.AcquirePoint(jppo);

            // check if is OK
            if (pdr.Status != PromptStatus.OK)
                return SamplerStatus.Cancel;
            if (pdr.Value.DistanceTo(newlocation)
              < Tolerance.Global.EqualPoint)
                return SamplerStatus.NoChange;

            // and store the new valuye
            newlocation = new Point3d(oldlocation.X, pdr.Value.Y, 0);
            return SamplerStatus.OK;
        }

        /// <summary>
        /// Return the new location, during/after JIG
        /// </summary>
        public Point3d Location { get { return newlocation; } }


    }

}

namespace NESW
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
                    JigPromptPointOptions prOptions1 = new JigPromptPointOptions("\nX Location:");
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

namespace NESW
{
    public class Mtextjig_Y : EntityJig
    {
        #region Fields

        public int mCurJigFactorIndex = 1;  // Jig Factor Index

        public Autodesk.AutoCAD.Geometry.Point3d mArrowLocation; // Jig Factor #1
        public Autodesk.AutoCAD.Geometry.Point3d mTextLocation; // Jig Factor #2
        public string _content; // Jig Factor #3
        double _x = 0;

        #endregion

        #region Constructors

        public Mtextjig_Y(MText mtext1, string content1, double rot1, double height1, AttachmentPoint align1, double x) : base(mtext1)
        {
            Mt.SetDatabaseDefaults();

            Mt.Contents = content1;
            Mt.Rotation = rot1;
            Mt.TextHeight = height1;
            Mt.Attachment = align1;
            _x = x;
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
                    Mt.Location = new Point3d(_x, mArrowLocation.Y, 0);
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
                    JigPromptPointOptions prOptions1 = new JigPromptPointOptions("\nY location:");
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
                        mArrowLocation = new Point3d(_x, prResult1.Value.Y, 0);
                        return SamplerStatus.OK;
                    }


                default:
                    break;
            }

            return SamplerStatus.OK;
        }



        #endregion

        #region Methods to Call

        public static MText Jig(string content1, double rot1, double height1, AttachmentPoint align1, double X)
        {
            Mtextjig_Y jigger = null;
            try
            {
                jigger = new Mtextjig_Y(new MText(), content1, rot1, height1, align1, X);
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


    }
}
