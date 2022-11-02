using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.ApplicationServices;

namespace Dimensioning
{





    class Acholade_Jig_Class : DrawJig
    {
        Point3d Point1;
        Point3d Point2;
        Point3d Point3;

        Double Arr_len;
        Double Arr_width;
        Double rad_small;
        PromptPointResult PromptPointResult1;

        public PromptPointResult StartJig(Point3d Punct1, Point3d Punct2, Double Arrow_length, Double Arrow_width, Double small_radius)
        {
            Point1 = Punct1;
            Point2 = Punct2;

            Arr_len = Arrow_length;
            Arr_width = Arrow_width;
            rad_small = small_radius;


            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            PromptPointResult1 = (PromptPointResult)ed.Drag(this);
            do
            {
                if (PromptPointResult1.Status == PromptStatus.OK)
                {
                    return PromptPointResult1;

                }
            } while (PromptPointResult1.Status != PromptStatus.Cancel);

            return PromptPointResult1;
        }

        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            PromptPointResult1 = prompts.AcquirePoint("\nPick label location:");

            if (PromptPointResult1.Value.IsEqualTo(Point3))
            {
                return SamplerStatus.NoChange;
            }
            else
            {
                Point3 = PromptPointResult1.Value;
                return SamplerStatus.OK;
            }

        }

        protected override bool WorldDraw(Autodesk.AutoCAD.GraphicsInterface.WorldDraw draw)
        {
            CircularArc3d Arc3d = new CircularArc3d(Point1, Point2, Point3);
            Circle Circle1 = new Circle(Arc3d.Center, Vector3d.ZAxis, Point1.GetVectorTo(Arc3d.Center).Length);


            Line Line1 = new Line(Circle1.Center, Point3);
            Double Scale_f = (Math.Pow(((Circle1.Radius + rad_small) * (Circle1.Radius + rad_small) - rad_small * rad_small), 0.5)) / Circle1.Radius;
            Line1.TransformBy(Matrix3d.Scaling(Scale_f, Circle1.Center));



            Point3d PointT = new Point3d();
            PointT = Line1.EndPoint;
            Point3d PointA = new Point3d();
            PointA = Line1.GetPointAtDist(Line1.Length - rad_small);
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


                    Point3d PointB3 = new Point3d();
                    Point3d PointB4 = new Point3d();
                    PointB3 = Point1;
                    PointB4 = Point2;


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
                        Double b01 = Math.Tan(Arc31.TotalAngle / 4);




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
                            Double b41 = Math.Tan(Arc41.TotalAngle / 4);
                            Poly123.AddVertexAt(4, new Point2d(PtI1.X, PtI1.Y), b41, 0, 0);



                            AngleStart = Functions.GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PointB4.X, PointB4.Y);
                            AngleEnd = Functions.GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PtArr4.X, PtArr4.Y);
                            Arc Arc42 = new Arc(Circle1.Center, Circle1.Radius, AngleEnd, AngleStart);
                            double b42 = Math.Tan(Arc42.TotalAngle / 4);



                            Poly123.AddVertexAt(5, new Point2d(PtArr4.X, PtArr4.Y), b42, Arr_width, 0);



                            Poly123.AddVertexAt(6, new Point2d(PointB4.X, PointB4.Y), 0, 0, 0);



                            draw.Geometry.Polyline(Poly123, 0, 6);
                        }
                    }
                }
            }




            return true;
        }

    }
}
namespace Dimensioning
{

    public class jig_Mtext_at_acholade_class : EntityJig
    {
        MText Mtext1;
        Point3d Location_point = new Point3d();
        JigPromptPointOptions Jig_options = new JigPromptPointOptions();

        public jig_Mtext_at_acholade_class(MText Mtext_entity, double TextH, double Textrot, string Content)
            : base(Mtext_entity)
        {
            Mtext1 = Mtext_entity;
            Mtext1.Contents = Content;
            Mtext1.Rotation = Textrot;
            Mtext1.TextHeight = TextH;
            Mtext1.Attachment = AttachmentPoint.MiddleCenter;
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
            PromptResult PromptResult1 = Editor1.Drag(new jig_Mtext_at_acholade_class(Mtext1, Mtext1.TextHeight, Mtext1.Rotation, Mtext1.Contents));

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
