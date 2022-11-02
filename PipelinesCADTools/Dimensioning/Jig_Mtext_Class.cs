

using System;
using System.Text;
using System.Linq;
using System.Xml;
using System.Reflection;
using System.ComponentModel;
using System.Collections;
using System.Collections.Generic;
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




namespace Jig1
{


    public class jig_Mtext_class : EntityJig
    {
        MText Mtext1;
        Point3d Location_point = new Point3d();
        JigPromptPointOptions Jig_options = new JigPromptPointOptions();

        public jig_Mtext_class(MText Mtext_entity, Double TextH, double Textrot, String Content, ObjectId txtSid)
            : base(Mtext_entity)
        {
            Mtext1 = Mtext_entity;
            Mtext1.Contents = Content;
            Mtext1.Rotation = Textrot;
            Mtext1.TextHeight = TextH;
            Mtext1.Attachment = AttachmentPoint.MiddleCenter;
            Mtext1.TextStyleId = txtSid;
        }

        public PromptPointResult BeginJig()
        {
            if (Jig_options == null)
            {
                Jig_options = new JigPromptPointOptions();
                Jig_options.Message = "/nSpecify insertion point:";
            }
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            PromptResult PromptResult1 = Editor1.Drag(new jig_Mtext_class(Mtext1, Mtext1.TextHeight, Mtext1.Rotation, Mtext1.Contents, Mtext1.TextStyleId));

            while (PromptResult1.Status != PromptStatus.Cancel)
            {
                switch (PromptResult1.Status)
                {
                    case PromptStatus.OK:
                        return (PromptPointResult)PromptResult1;

                    case PromptStatus.None:
                        return (PromptPointResult)PromptResult1;
                    case PromptStatus.Other:
                        return (PromptPointResult)PromptResult1;
                }

            }

            return null;
        }


        protected override bool Update()
        {
            Mtext1.Location = Location_point;

            return false;
        }

        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            PromptPointResult Result1 = prompts.AcquirePoint(Jig_options);
            if (Result1.Value.IsEqualTo(Location_point) == false)
            {
                Location_point = Result1.Value;
            }
            else
            {
                return SamplerStatus.NoChange;
            }
            if (Result1.Status == PromptStatus.Cancel)
            {
                return SamplerStatus.Cancel;
                Jig_options.Message = "/nCanceled";
            }

            else
            {
                return SamplerStatus.OK;
            }
        }



    }



}

public class RotateJig : EntityJig
{
    #region Fields

    public int mCurJigFactorIndex = 2;
    private Point3d mBasePoint = Point3d.Origin;
    private double mLastAngle = 0;
    private double mLastScaleFactor = 1;

    private Point3d mPosition;                  //Factor #1
    private double mRotationAngle;              //Factor #2
    private double mScaleFactor;                //Factor #3

    #endregion

    #region Constructors

    public RotateJig(Entity ent, Point3d basePoint)
        : base(ent)
    {
        mBasePoint = basePoint;

    }

    #endregion

    #region Properties

    private static Editor Editor
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
            return Editor.CurrentUserCoordinateSystem;
        }
    }

    #endregion

    #region Overrides

    protected override bool Update()
    {
        switch (mCurJigFactorIndex)
        {
            case 1:
                Matrix3d mat1 = Matrix3d.Displacement(mBasePoint.GetVectorTo(mPosition).TransformBy(UCS));
                Entity.TransformBy(mat1);
                mBasePoint = mPosition;
                break;
            case 2:
                Matrix3d mat2 = Matrix3d.Rotation(mRotationAngle - mLastAngle, Vector3d.ZAxis.TransformBy(UCS), mBasePoint.TransformBy(UCS));
                Entity.TransformBy(mat2);
                mLastAngle = mRotationAngle;
                break;
            case 3:
                Matrix3d mat3 = Matrix3d.Scaling(mScaleFactor / mLastScaleFactor, mBasePoint.TransformBy(UCS));
                Entity.TransformBy(mat3);
                mLastScaleFactor = mScaleFactor;
                break;
            default:
                break;
        }

        return true;
    }

    protected override SamplerStatus Sampler(JigPrompts prompts)
    {
        switch (mCurJigFactorIndex)
        {
            case 1:
                JigPromptPointOptions prOptions1 = new JigPromptPointOptions("\nNew location:");
                PromptPointResult prResult1 = prompts.AcquirePoint(prOptions1);
                if (prResult1.Status == PromptStatus.Cancel) return SamplerStatus.Cancel;

                Point3d tempPt = prResult1.Value.TransformBy(UCS.Inverse());
                if (tempPt.Equals(mPosition))
                {
                    return SamplerStatus.NoChange;
                }
                else
                {
                    mPosition = tempPt;
                    return SamplerStatus.OK;
                }
            case 2:
                JigPromptAngleOptions prOptions2 = new JigPromptAngleOptions("\nRotation angle:");
                prOptions2.BasePoint = mBasePoint.TransformBy(UCS);
                prOptions2.UseBasePoint = true;
                PromptDoubleResult prResult2 = prompts.AcquireAngle(prOptions2);

                if (prResult2.Status == PromptStatus.Cancel) return SamplerStatus.Cancel;

                if (prResult2.Value.Equals(mRotationAngle))
                {
                    return SamplerStatus.NoChange;
                }
                else
                {
                    mRotationAngle = prResult2.Value;
                    return SamplerStatus.OK;
                }
            case 3:
                JigPromptDistanceOptions prOptions3 = new JigPromptDistanceOptions("\nScale factor:");
                prOptions3.BasePoint = mBasePoint.TransformBy(UCS);
                prOptions3.UseBasePoint = true;
                PromptDoubleResult prResult3 = prompts.AcquireDistance(prOptions3);

                if (prResult3.Status == PromptStatus.Cancel) return SamplerStatus.Cancel;

                if (prResult3.Value.Equals(mScaleFactor))
                {
                    return SamplerStatus.NoChange;
                }
                else
                {
                    mScaleFactor = prResult3.Value;
                    return SamplerStatus.OK;
                }
            default:
                break;
        }

        return SamplerStatus.OK;
    }

    #endregion

    #region Methods to Call

    public static bool Jig(Entity ent, Point3d basePt)
    {
        RotateJig jigger = null;
        try
        {
            jigger = new RotateJig(ent, basePt);
            PromptResult pr;

            do
            {
                pr = Editor.Drag(jigger);
               jigger.mCurJigFactorIndex++;

            } while (pr.Status != PromptStatus.Cancel && pr.Status != PromptStatus.Error && jigger.mCurJigFactorIndex <2);


            if (pr.Status == PromptStatus.Cancel || pr.Status == PromptStatus.Error)
                return CleanUp(jigger);
            else
                return true;



        }
        catch
        {
            return CleanUp(jigger);
        }
    }

    private static bool CleanUp(RotateJig jigger)
    {
        if (jigger != null && jigger.Entity != null)
            jigger.Entity.Dispose();

        return false;
    }

    #endregion

    #region Test Commands

    [CommandMethod("rj")]
    public static void RotateJig_Method()
    {
        Database db = HostApplicationServices.WorkingDatabase;
        try
        {
            PromptEntityResult selRes = Editor.GetEntity("\nPick an entity:");
            PromptPointResult ptRes = Editor.GetPoint("\nBase Point: ");
            if (selRes.Status == PromptStatus.OK && ptRes.Status == PromptStatus.OK)
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    Entity ent = tr.GetObject(selRes.ObjectId, OpenMode.ForWrite) as Entity;
                    if (ent != null)
                    {
                        using (Transaction tr1 = db.TransactionManager.StartTransaction())
                        {
                            ent.Highlight();
                            tr1.Commit();
                        }

                        if (RotateJig.Jig(ent, ptRes.Value))
                            tr.Commit();
                    }
                }
            }
        }
        catch (System.Exception ex)
        {
            Editor.WriteMessage(ex.ToString());
        }
    }



    #endregion
}

public class MoveRotateScaleJig : EntityJig
{
    #region Fields

    public int mCurJigFactorIndex = 1;
    private Point3d mBasePoint = Point3d.Origin;
    private double mLastAngle = 0;
    private double mLastScaleFactor = 1;

    private Point3d mPosition;                  //Factor #1
    private double mRotationAngle;              //Factor #2
    private double mScaleFactor;                //Factor #3

    #endregion

    #region Constructors

    public MoveRotateScaleJig(Entity ent, Point3d basePoint)
        : base(ent)
    {
        mBasePoint = basePoint;

    }

    #endregion

    #region Properties

    private static Editor Editor
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
            return Editor.CurrentUserCoordinateSystem;
        }
    }

    #endregion

    #region Overrides

    protected override bool Update()
    {
        switch (mCurJigFactorIndex)
        {
            case 1:
                Matrix3d mat1 = Matrix3d.Displacement(mBasePoint.GetVectorTo(mPosition).TransformBy(UCS));
                Entity.TransformBy(mat1);
                mBasePoint = mPosition;
                break;
            case 2:
                Matrix3d mat2 = Matrix3d.Rotation(mRotationAngle - mLastAngle, Vector3d.ZAxis.TransformBy(UCS), mBasePoint.TransformBy(UCS));
                Entity.TransformBy(mat2);
                mLastAngle = mRotationAngle;
                break;
            case 3:
                Matrix3d mat3 = Matrix3d.Scaling(mScaleFactor / mLastScaleFactor, mBasePoint.TransformBy(UCS));
                Entity.TransformBy(mat3);
                mLastScaleFactor = mScaleFactor;
                break;
            default:
                break;
        }

        return true;
    }

    protected override SamplerStatus Sampler(JigPrompts prompts)
    {
        switch (mCurJigFactorIndex)
        {
            case 1:
                JigPromptPointOptions prOptions1 = new JigPromptPointOptions("\nNew location:");
                PromptPointResult prResult1 = prompts.AcquirePoint(prOptions1);
                if (prResult1.Status == PromptStatus.Cancel) return SamplerStatus.Cancel;

                Point3d tempPt = prResult1.Value.TransformBy(UCS.Inverse());
                if (tempPt.Equals(mPosition))
                {
                    return SamplerStatus.NoChange;
                }
                else
                {
                    mPosition = tempPt;
                    return SamplerStatus.OK;
                }
            case 2:
                JigPromptAngleOptions prOptions2 = new JigPromptAngleOptions("\nRotation angle:");
                prOptions2.BasePoint = mBasePoint.TransformBy(UCS);
                prOptions2.UseBasePoint = true;
                PromptDoubleResult prResult2 = prompts.AcquireAngle(prOptions2);

                if (prResult2.Status == PromptStatus.Cancel) return SamplerStatus.Cancel;

                if (prResult2.Value.Equals(mRotationAngle))
                {
                    return SamplerStatus.NoChange;
                }
                else
                {
                    mRotationAngle = prResult2.Value;
                    return SamplerStatus.OK;
                }
            case 3:
                JigPromptDistanceOptions prOptions3 = new JigPromptDistanceOptions("\nScale factor:");
                prOptions3.BasePoint = mBasePoint.TransformBy(UCS);
                prOptions3.UseBasePoint = true;
                PromptDoubleResult prResult3 = prompts.AcquireDistance(prOptions3);

                if (prResult3.Status == PromptStatus.Cancel) return SamplerStatus.Cancel;

                if (prResult3.Value.Equals(mScaleFactor))
                {
                    return SamplerStatus.NoChange;
                }
                else
                {
                    mScaleFactor = prResult3.Value;
                    return SamplerStatus.OK;
                }
            default:
                break;
        }

        return SamplerStatus.OK;
    }

    #endregion

    #region Methods to Call

    public static bool Jig(Entity ent, Point3d basePt)
    {
        MoveRotateScaleJig jigger = null;
        try
        {
            jigger = new MoveRotateScaleJig(ent, basePt);
            PromptResult pr;
           

            do
            {
                pr = Editor.Drag(jigger);

              

                    
                    jigger.mCurJigFactorIndex++;
                




            } while (pr.Status != PromptStatus.Cancel && pr.Status != PromptStatus.Error && jigger.mCurJigFactorIndex <= 3);


            if (pr.Status == PromptStatus.Cancel || pr.Status == PromptStatus.Error)
                return CleanUp(jigger);
            else
                return true;



        }
        catch
        {
            return CleanUp(jigger);
        }
    }

    private static bool CleanUp(MoveRotateScaleJig jigger)
    {
        if (jigger != null && jigger.Entity != null)
            jigger.Entity.Dispose();

        return false;
    }

    #endregion

    #region Test Commands

    [CommandMethod("MRSJ")]
    public static void MoveRotateScaleJig_Method()
    {
        Database db = HostApplicationServices.WorkingDatabase;
        try
        {
            PromptEntityResult selRes = Editor.GetEntity("\nPick an entity:");
            PromptPointResult ptRes = Editor.GetPoint("\nBase Point: ");
            if (selRes.Status == PromptStatus.OK && ptRes.Status == PromptStatus.OK)
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    Entity ent = tr.GetObject(selRes.ObjectId, OpenMode.ForWrite) as Entity;
                    if (ent != null)
                    {
                        using (Transaction tr1 = db.TransactionManager.StartTransaction())
                        {
                            ent.Highlight();
                            tr1.Commit();
                        }

                        if (MoveRotateScaleJig.Jig(ent, ptRes.Value))
                            tr.Commit();
                    }
                }
            }
        }
        catch (System.Exception ex)
        {
            Editor.WriteMessage(ex.ToString());
        }
    }



    #endregion
}
