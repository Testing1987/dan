using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

namespace Alignment_mdi
{

    class Jig_rectangle_viewport_along2D_manual_pt2 : DrawJig
    {
        Polyline Poly_cl2d;

        Point3d Punct1 = new Point3d();
        Point3d Punct2 = new Point3d();


        double Vw_height;
        double Vw_scale;

        double Max_dist1;
        double Min_dist1;

        PromptPointResult PromptPointResult1;


        public PromptPointResult StartJig(double VVw_Scale, double VVw_width, double VVw_height, Polyline Polyline_2d, Point3d Pt1, double min_dist)
        {

            Vw_height = VVw_height;
            Vw_scale = VVw_Scale;
            Punct1 = Pt1;
            Poly_cl2d = Polyline_2d;
            Max_dist1 = VVw_width;
            Min_dist1 = min_dist;
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
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
            PromptPointResult1 = prompts.AcquirePoint("\nPlease pick end location:");

            if (PromptPointResult1.Value.IsEqualTo(Punct2))
            {
                return SamplerStatus.NoChange;
            }
            else
            {
                Punct2 = PromptPointResult1.Value;
                return SamplerStatus.OK;
            }

        }

        protected override bool WorldDraw(Autodesk.AutoCAD.GraphicsInterface.WorldDraw draw)
        {

            Point3d Pt_on_poly1;
            Point3d Pt_on_poly2;

            Pt_on_poly1 = Poly_cl2d.GetClosestPointTo(Punct1, Vector3d.ZAxis, false);
            Pt_on_poly2 = Poly_cl2d.GetClosestPointTo(Punct2, Vector3d.ZAxis, false);


            //double dist1 = Pt_on_poly1.DistanceTo(Pt_on_poly2);
            double dist1 = Poly_cl2d.GetDistAtPoint(Pt_on_poly2) - Poly_cl2d.GetDistAtPoint(Pt_on_poly1);

            if (dist1 > Min_dist1 & dist1 <= Max_dist1)
            {
                Polyline Poly1 = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                Poly1 = creaza_rectangle_viewport(Pt_on_poly1, Pt_on_poly2);

                draw.Geometry.Polyline(Poly1, 0, 4);
            }

            return true;
        }

        private Polyline creaza_rectangle_viewport(Point3d Point1, Point3d Point2)
        {
            Line Line1R = new Line(Point1, Point2);
            Point3d Point_distR = new Point3d();
            if (Line1R.Length > Vw_height / Vw_scale)
            {
                Point_distR = Line1R.GetPointAtDist(Vw_height / Vw_scale);
                Line1R.EndPoint = Point_distR;
            }
            else
            {
                Line1R.TransformBy(Matrix3d.Scaling((Vw_height / Vw_scale) / Line1R.Length, Line1R.StartPoint));
            }

            Line1R.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, Point1));
            Point3d Point_middler = new Point3d((Point1.X + Line1R.EndPoint.X) / 2, (Point1.Y + Line1R.EndPoint.Y) / 2, 0);

            Line1R.TransformBy(Matrix3d.Displacement(Point_middler.GetVectorTo(Point1)));
            Point3d Pt1r = new Point3d();
            Pt1r = Line1R.StartPoint;
            Point3d Pt2r = new Point3d();
            Pt2r = Line1R.EndPoint;
            Line1R.TransformBy(Matrix3d.Displacement(Point1.GetVectorTo(Point2)));

            Point3d Pt4r = new Point3d();
            Pt4r = Line1R.StartPoint;
            Point3d Pt3r = new Point3d();
            Pt3r = Line1R.EndPoint;

            Polyline Poly1r = new Autodesk.AutoCAD.DatabaseServices.Polyline();
            Poly1r.AddVertexAt(0, new Point2d(Pt1r.X, Pt1r.Y), 0, 0, 0);
            Poly1r.AddVertexAt(1, new Point2d(Pt2r.X, Pt2r.Y), 0, 0, 0);
            Poly1r.AddVertexAt(2, new Point2d(Pt3r.X, Pt3r.Y), 0, 0, 0);
            Poly1r.AddVertexAt(3, new Point2d(Pt4r.X, Pt4r.Y), 0, 0, 0);
            Poly1r.AddVertexAt(4, new Point2d(Pt1r.X, Pt1r.Y), 0, 0, 0);



            return Poly1r;
        }




    }

    class Jig_rectangle_viewport_SHEET_CUTTER_manual_up_down : DrawJig
    {



        Point3d Base_point;
        Point3d New_position;
        private List<Entity> mEntities = new List<Entity>();

        Line Linie1;


        public Jig_rectangle_viewport_SHEET_CUTTER_manual_up_down(Point3d basePt, Line Line1)
        {
            Base_point = basePt.TransformBy(UCS);
            mEntities = new List<Entity>();
            Linie1 = Line1;
        }


        public Point3d Base
        {
            get
            {
                return New_position;
            }
            set
            {
                New_position = value;
            }
        }


        public Point3d Location
        {
            get
            {
                return New_position;
            }
            set
            {
                New_position = value;
            }
        }

        public Matrix3d UCS
        {
            get
            {
                return Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.CurrentUserCoordinateSystem;
            }
        }

        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            JigPromptPointOptions prOptions1 = new JigPromptPointOptions("\nNew location:");
            prOptions1.UseBasePoint = false;

            PromptPointResult prResult1 = prompts.AcquirePoint(prOptions1);
            if (prResult1.Status == PromptStatus.Cancel && prResult1.Status == PromptStatus.Error)
            {
                return SamplerStatus.Cancel;
            }

            if (New_position.IsEqualTo(prResult1.Value, new Tolerance(0.000000001, 0.000000001)))
            {

                return SamplerStatus.NoChange;

            }
            else
            {
                New_position = prResult1.Value;
                return SamplerStatus.OK;
            }
        }

        public void AddEntity(Entity ent)
        {
            mEntities.Add(ent);
        }

        public void TransformEntities()
        {
            Matrix3d Move_matrix = Matrix3d.Displacement(Base_point.GetVectorTo(Linie1.GetClosestPointTo(New_position, Vector3d.ZAxis, true)));

            foreach (Entity ent in mEntities)
            {
                ent.TransformBy(Move_matrix);
            }
        }

        protected override bool WorldDraw(Autodesk.AutoCAD.GraphicsInterface.WorldDraw draw)
        {
            Matrix3d Move_matrix = Matrix3d.Displacement(Base_point.GetVectorTo(Linie1.GetClosestPointTo(New_position, Vector3d.ZAxis, true)));

            Autodesk.AutoCAD.GraphicsInterface.WorldGeometry geo = draw.Geometry;
            if (geo != null)
            {
                geo.PushModelTransform(Move_matrix);

                foreach (Entity ent in mEntities)
                {
                    geo.Draw(ent);
                }

                geo.PopModelTransform();
            }

            return true;
        }




    }

    class Jig_rectangle_viewport_along3D_manual_pt2 : DrawJig
    {
        Polyline3d Poly_cl3d;
        Polyline Poly_cl2d;

        Point3d Punct1 = new Point3d();
        Point3d Punct2 = new Point3d();


        double Vw_height;
        double Vw_scale;

        double Max_dist1;
        double Min_dist1;

        PromptPointResult PromptPointResult1;


        public PromptPointResult StartJig(double VVw_Scale, double VVw_width, double VVw_height, Polyline3d Polyline_3d, Polyline Polyline_2d, Point3d Pt1, double min_dist)
        {

            Vw_height = VVw_height;
            Vw_scale = VVw_Scale;
            Punct1 = Pt1;
            Poly_cl3d = Polyline_3d;
            Poly_cl2d = Polyline_2d;

            Max_dist1 = VVw_width;


            Min_dist1 = min_dist;
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
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
            PromptPointResult1 = prompts.AcquirePoint("\nPlease pick end location:");

            if (PromptPointResult1.Value.IsEqualTo(Punct2))
            {
                return SamplerStatus.NoChange;
            }
            else
            {
                Punct2 = PromptPointResult1.Value;
                return SamplerStatus.OK;
            }

        }

        protected override bool WorldDraw(Autodesk.AutoCAD.GraphicsInterface.WorldDraw draw)
        {

            Point3d Pt_on_poly1;
            Point3d Pt_on_poly2;

            Pt_on_poly1 = Poly_cl2d.GetClosestPointTo(Punct1, Vector3d.ZAxis, false);
            Pt_on_poly2 = Poly_cl2d.GetClosestPointTo(Punct2, Vector3d.ZAxis, false);
            double Param1 = Poly_cl2d.GetParameterAtPoint(Pt_on_poly1);
            double Param2 = Poly_cl2d.GetParameterAtPoint(Pt_on_poly2);
            Point3d Pt_on_poly11 = Poly_cl3d.GetPointAtParameter(Param1);
            Point3d Pt_on_poly12 = Poly_cl3d.GetPointAtParameter(Param2);

            double dist1 = Poly_cl3d.GetDistAtPoint(Pt_on_poly12) - Poly_cl3d.GetDistAtPoint(Pt_on_poly11);

            if (dist1 > Min_dist1 & dist1 <= Max_dist1)
            {
                Polyline Poly1 = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                Poly1 = creaza_rectangle_viewport(Pt_on_poly1, Pt_on_poly2);

                draw.Geometry.Polyline(Poly1, 0, 4);
            }

            return true;
        }

        private Polyline creaza_rectangle_viewport(Point3d Point1, Point3d Point2)
        {
            Line Line1R = new Line(Point1, Point2);
            Point3d Point_distR = new Point3d();
            if (Line1R.Length > Vw_height / Vw_scale)
            {
                Point_distR = Line1R.GetPointAtDist(Vw_height / Vw_scale);
                Line1R.EndPoint = Point_distR;
            }
            else
            {
                Line1R.TransformBy(Matrix3d.Scaling((Vw_height / Vw_scale) / Line1R.Length, Line1R.StartPoint));
            }

            Line1R.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, Point1));
            Point3d Point_middler = new Point3d((Point1.X + Line1R.EndPoint.X) / 2, (Point1.Y + Line1R.EndPoint.Y) / 2, 0);

            Line1R.TransformBy(Matrix3d.Displacement(Point_middler.GetVectorTo(Point1)));
            Point3d Pt1r = new Point3d();
            Pt1r = Line1R.StartPoint;
            Point3d Pt2r = new Point3d();
            Pt2r = Line1R.EndPoint;
            Line1R.TransformBy(Matrix3d.Displacement(Point1.GetVectorTo(Point2)));

            Point3d Pt4r = new Point3d();
            Pt4r = Line1R.StartPoint;
            Point3d Pt3r = new Point3d();
            Pt3r = Line1R.EndPoint;

            Polyline Poly1r = new Autodesk.AutoCAD.DatabaseServices.Polyline();
            Poly1r.AddVertexAt(0, new Point2d(Pt1r.X, Pt1r.Y), 0, 0, 0);
            Poly1r.AddVertexAt(1, new Point2d(Pt2r.X, Pt2r.Y), 0, 0, 0);
            Poly1r.AddVertexAt(2, new Point2d(Pt3r.X, Pt3r.Y), 0, 0, 0);
            Poly1r.AddVertexAt(3, new Point2d(Pt4r.X, Pt4r.Y), 0, 0, 0);
            Poly1r.AddVertexAt(4, new Point2d(Pt1r.X, Pt1r.Y), 0, 0, 0);



            return Poly1r;
        }




    }

    class Jig_rectangle_viewport_manual : DrawJig
    {


        Point3d Punct1 = new Point3d();
        Point3d Punct2 = new Point3d();


        double Vw_height;
        double Vw_scale;


        double Min_dist1;

        PromptPointResult PromptPointResult1;


        public PromptPointResult StartJig(double VVw_Scale, double VVw_height, Point3d Pt1, double min_dist)
        {

            Vw_height = VVw_height;
            Vw_scale = VVw_Scale;
            Punct1 = Pt1;

            Min_dist1 = min_dist;
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
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
            PromptPointResult1 = prompts.AcquirePoint("\nPlease pick end location:");

            if (PromptPointResult1.Value.IsEqualTo(Punct2))
            {
                return SamplerStatus.NoChange;
            }
            else
            {
                Punct2 = PromptPointResult1.Value;
                return SamplerStatus.OK;
            }

        }

        protected override bool WorldDraw(Autodesk.AutoCAD.GraphicsInterface.WorldDraw draw)
        {

            double dist1 = Punct1.DistanceTo(Punct2);

            if (dist1 > Min_dist1)
            {
                Polyline Poly1 = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                Poly1 = creaza_rectangle_viewport(Punct1, Punct2);

                draw.Geometry.Polyline(Poly1, 0, 4);
            }

            return true;
        }

        private Polyline creaza_rectangle_viewport(Point3d Point1, Point3d Point2)
        {
            Line Line1R = new Line(Point1, Point2);


            Point3d Point_distR = new Point3d();
            if (Line1R.Length > Vw_height / Vw_scale)
            {
                Point_distR = Line1R.GetPointAtDist(Vw_height / Vw_scale);
                Line1R.EndPoint = Point_distR;
            }
            else
            {
                Line1R.TransformBy(Matrix3d.Scaling((Vw_height / Vw_scale) / Line1R.Length, Line1R.StartPoint));
            }

            Line1R.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, Point1));
            Point3d Point_middler = new Point3d((Point1.X + Line1R.EndPoint.X) / 2, (Point1.Y + Line1R.EndPoint.Y) / 2, 0);

            Line1R.TransformBy(Matrix3d.Displacement(Point_middler.GetVectorTo(Point1)));
            Point3d Pt1r = new Point3d();
            Pt1r = Line1R.StartPoint;
            Point3d Pt2r = new Point3d();
            Pt2r = Line1R.EndPoint;
            Line1R.TransformBy(Matrix3d.Displacement(Point1.GetVectorTo(Point2)));

            Point3d Pt4r = new Point3d();
            Pt4r = Line1R.StartPoint;
            Point3d Pt3r = new Point3d();
            Pt3r = Line1R.EndPoint;

            Polyline Poly1r = new Autodesk.AutoCAD.DatabaseServices.Polyline();
            Poly1r.AddVertexAt(0, new Point2d(Pt1r.X, Pt1r.Y), 0, 0, 0);
            Poly1r.AddVertexAt(1, new Point2d(Pt2r.X, Pt2r.Y), 0, 0, 0);
            Poly1r.AddVertexAt(2, new Point2d(Pt3r.X, Pt3r.Y), 0, 0, 0);
            Poly1r.AddVertexAt(3, new Point2d(Pt4r.X, Pt4r.Y), 0, 0, 0);
            Poly1r.AddVertexAt(4, new Point2d(Pt1r.X, Pt1r.Y), 0, 0, 0);



            return Poly1r;
        }




    }


    class Jig_rectangle_viewport_pick_points : DrawJig
    {

        Point3d Punct1 = new Point3d();
        Point3d Punct2 = new Point3d();
        string mesaj_1 = "";


        double Min_dist1;

        PromptPointResult PromptPointResult1;



        public PromptPointResult StartJig(Point3d Basept, double min_dist, string mesaj)
        {


            mesaj_1 = mesaj;
            Punct1 = Basept;
            Min_dist1 = min_dist;
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
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
            PromptPointResult1 = prompts.AcquirePoint(mesaj_1);

            if (PromptPointResult1.Value.IsEqualTo(Punct2))
            {
                return SamplerStatus.NoChange;
            }
            else
            {
                Punct2 = PromptPointResult1.Value;
                return SamplerStatus.OK;
            }

        }

        protected override bool WorldDraw(Autodesk.AutoCAD.GraphicsInterface.WorldDraw draw)
        {




            //double dist1 = Pt_on_poly1.DistanceTo(Pt_on_poly2);
            double dist1 = Punct1.DistanceTo(Punct2);

            if (dist1 > Min_dist1)
            {
                Polyline Poly1 = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                Poly1.AddVertexAt(0, new Point2d(Punct1.X, Punct1.Y), 0, 0, 0);
                Poly1.AddVertexAt(1, new Point2d(Punct1.X, Punct2.Y), 0, 0, 0);
                Poly1.AddVertexAt(2, new Point2d(Punct2.X, Punct2.Y), 0, 0, 0);
                Poly1.AddVertexAt(3, new Point2d(Punct2.X, Punct1.Y), 0, 0, 0);
                Poly1.AddVertexAt(4, new Point2d(Punct1.X, Punct1.Y), 0, 0, 0);

                draw.Geometry.Polyline(Poly1, 0, 4);
            }

            return true;
        }





    }

    class Jig_rectangle_viewport_pick_points_on_line : DrawJig
    {

        Point3d Punct1 = new Point3d();
        Point3d Punct2 = new Point3d();
        string mesaj_1 = "";


        double Min_dist1;

        PromptPointResult PromptPointResult1;



        public PromptPointResult StartJig(Point3d Basept, double min_dist, string mesaj)
        {

            mesaj_1 = mesaj;

            Punct1 = Basept;
            Min_dist1 = min_dist;
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
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
            PromptPointResult1 = prompts.AcquirePoint(mesaj_1);

            if (PromptPointResult1.Value.IsEqualTo(Punct2))
            {
                return SamplerStatus.NoChange;
            }
            else
            {
                Punct2 = PromptPointResult1.Value;
                return SamplerStatus.OK;
            }

        }

        protected override bool WorldDraw(Autodesk.AutoCAD.GraphicsInterface.WorldDraw draw)
        {




            //double dist1 = Pt_on_poly1.DistanceTo(Pt_on_poly2);
            double dist1 = Punct1.DistanceTo(Punct2);

            if (dist1 > Min_dist1)
            {
                Polyline Poly1 = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                Poly1.AddVertexAt(0, new Point2d(Punct1.X, Punct1.Y), 0, 0, 0);
                Poly1.AddVertexAt(1, new Point2d(Punct2.X, Punct2.Y), 0, 0, 0);
                draw.Geometry.Polyline(Poly1, 0, 1);
            }

            return true;
        }





    }

}
