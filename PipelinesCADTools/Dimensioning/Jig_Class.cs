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

    class Jig_Class1 : DrawJig
    {
        Point3d Basept;
        Point3d Base_start_point;

        Double Inaltimea1;
        PromptPointResult PromptPointResult1;


        public PromptPointResult StartJig(Point3d Punct1,  Double Scalefactor)
        {
            Base_start_point = Punct1;
            Inaltimea1 = Scalefactor;
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
            PromptPointResult1 = prompts.AcquirePoint("\nPick second point : ");

            if (PromptPointResult1.Value.IsEqualTo(Basept))
            {
                return SamplerStatus.NoChange;
            }
            else
            {
                Basept = PromptPointResult1.Value;
                return SamplerStatus.OK;
            }

        }

        protected override bool WorldDraw(Autodesk.AutoCAD.GraphicsInterface.WorldDraw draw)
        {

            Double x0 = Base_start_point.X;
            Double y0 = Base_start_point.Y;
            Double x = Basept.X;
            Double y = Basept.Y;



            Double Length1 = (Inaltimea1) / 500;
            Double Wdth1 = (Inaltimea1) / 1250;
            Double Dist1 = Math.Pow(Math.Pow(x0 - x, 2) + Math.Pow(y0 - y, 2), 0.5);




            Double xm = 0;
            Double ym = 0;
            Double x1, y1;

            Double Bear1 = Functions.GET_Bearing_rad(x0, y0, x, y);

            if (Bear1 < Math.PI / 2)
            {
                x1 = Length1 * Math.Cos(Bear1);
                y1 = Length1 * Math.Sin(Bear1);
                xm = x0 + x1;
                ym = y0 + y1;
            }

            if (Bear1 == Math.PI / 2)
            {
                x1 = 0;
                y1 = Length1;
                xm = x0;
                ym = y0 + y1;
            }

            if (Bear1 > Math.PI / 2 & Bear1 < Math.PI)
            {
                x1 = Length1 * Math.Cos(Math.PI - Bear1);
                y1 = Length1 * Math.Sin(Math.PI - Bear1);
                xm = x0 - x1;
                ym = y0 + y1;
            }

            if (Bear1 == Math.PI)
            {
                x1 = Length1;
                y1 = 0;
                xm = x0 - x1;
                ym = y0;
            }

            if (Bear1 > Math.PI & Bear1 < 3 * Math.PI / 2)
            {
                x1 = Length1 * Math.Cos(Bear1 - Math.PI);
                y1 = Length1 * Math.Sin(Bear1 - Math.PI);
                xm = x0 - x1;
                ym = y0 - y1;
            }

            if (Bear1 == 3 * Math.PI / 2)
            {
                x1 = 0;
                y1 = Length1;
                xm = x0;
                ym = y0 - y1;
            }

            if (Bear1 > 3 * Math.PI / 2 & Bear1 < 2 * Math.PI)
            {
                x1 = Length1 * Math.Cos(2 * Math.PI - Bear1);
                y1 = Length1 * Math.Sin(2 * Math.PI - Bear1);
                xm = x0 + x1;
                ym = y0 - y1;
            }

            if (Bear1 == 2 * Math.PI | Bear1 == 0)
            {
                x1 = Length1;
                y1 = 0;
                xm = x0 + x1;
                ym = y0;
            }





            Autodesk.AutoCAD.DatabaseServices.Polyline Poly1 = new Autodesk.AutoCAD.DatabaseServices.Polyline();
            Poly1.AddVertexAt(0, new Point2d(x0, y0), 0, 0, Wdth1);
            Poly1.AddVertexAt(1, new Point2d(xm, ym), 0, 0, 0);
            Poly1.AddVertexAt(2, new Point2d(x, y), 0, 0, 0);
            draw.Geometry.Polyline(Poly1, 0, 2);
            return true;
        }
        public Double GET_distanta_Double(Double x1, Double y1, Double x2, Double y2)
        {
            return Math.Pow(Math.Pow(x1 - x2, 2) + Math.Pow(y1 - y2, 2), 0.5);
        }


       


    }

    class Jig_Class2 : DrawJig
    {
        Point3d Basept;
        Point3d Base_start_point;
        Point3d Base_mid_point;
        Double Inaltimea1;
        PromptPointResult PromptPointResult1;

        public PromptPointResult StartJig(Point3d Punct1, Point3d Punct2, Double Scalefactor)
        {
            Base_start_point = Punct1;
            Base_mid_point = Punct2;
            Inaltimea1 = Scalefactor;
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
            PromptPointResult1 = prompts.AcquirePoint("\nPick third point : ");

            if (PromptPointResult1.Value.IsEqualTo(Basept))
            {
                return SamplerStatus.NoChange;
            }
            else
            {
                Basept = PromptPointResult1.Value;
                return SamplerStatus.OK;
            }

        }

        protected override bool WorldDraw(Autodesk.AutoCAD.GraphicsInterface.WorldDraw draw)
        {
            Double x0 = Base_start_point.X;
            Double y0 = Base_start_point.Y;

            Double x2 = Base_mid_point.X;
            Double y2 = Base_mid_point.Y;

            Double x3 = Basept.X;
            Double y3 = Basept.Y;


            Double Length1 = (Inaltimea1) / 500;
            Double Wdth1 = (Inaltimea1) / 1250;
            Double Dist1 = Math.Pow(Math.Pow(x0 - x2, 2) + Math.Pow(y0 - y2, 2), 0.5);


            Double x1 = X_Y_for_arc_leader(x0, y0, x2, y2, Inaltimea1, "X");
            Double y1 = X_Y_for_arc_leader(x0, y0, x2, y2, Inaltimea1, "Y");

            Double Bulge1 = Bulge_for_arc_leader(x0, y0, x2, y2, x3, y3, Inaltimea1);



            Autodesk.AutoCAD.DatabaseServices.Polyline Poly1 = new Autodesk.AutoCAD.DatabaseServices.Polyline();
            Poly1.AddVertexAt(0, new Point2d(x0, y0), 0, 0, Wdth1);
            Poly1.AddVertexAt(1, new Point2d(x1, y1), Bulge1, 0, 0);
            Poly1.AddVertexAt(2, new Point2d(x3, y3), 0, 0, 0);
            draw.Geometry.Polyline(Poly1, 0, 2);
            return true;
        }
        public Double GET_distanta_Double(Double x1, Double y1, Double x2, Double y2)
        {
            return Math.Pow(Math.Pow(x1 - x2, 2) + Math.Pow(y1 - y2, 2), 0.5);
        }


        public Double X_Y_for_arc_leader(Double x0, Double y0, Double x2, Double y2, Double Inaltimea, String X_or_Y)
        {
            Double Length1 = Inaltimea / 500;
            Double Dist1 = GET_distanta_Double(x0, y0, x2, y2);


            Double x1 = 0;
            Double y1 = 0;
            Double Dx;
            Double Dy;

            Double Bear1 = Functions.GET_Bearing_rad(x0, y0, x2, y2);

            if (Bear1 < Math.PI / 2)
            {
                Dx = Length1 * Math.Cos(Bear1);
                Dy = Length1 * Math.Sin(Bear1);
                x1 = x0 + Dx;
                y1 = y0 + Dy;
            }




            if (Bear1 == Math.PI / 2)
            {
                Dx = 0;
                Dy = Length1;
                x1 = x0;
                y1 = y0 + Dy;
            }

            if (Bear1 > Math.PI / 2 & Bear1 < Math.PI)
            {
                Dx = Length1 * Math.Cos(Math.PI - Bear1);
                Dy = Length1 * Math.Sin(Math.PI - Bear1);
                x1 = x0 - Dx;
                y1 = y0 + Dy;
            }

            if (Bear1 == Math.PI)
            {
                Dx = Length1;
                Dy = 0;
                x1 = x0 - Dx;
                y1 = y0;
            }

            if (Bear1 > Math.PI & Bear1 < 3 * Math.PI / 2)
            {
                Dx = Length1 * Math.Cos(Bear1 - Math.PI);
                Dy = Length1 * Math.Sin(Bear1 - Math.PI);
                x1 = x0 - Dx;
                y1 = y0 - Dy;
            }

            if (Bear1 == 3 * Math.PI / 2)
            {
                Dx = 0;
                Dy = Length1;
                x1 = x0;
                y1 = y0 - Dy;
            }

            if (Bear1 > 3 * Math.PI / 2 & Bear1 < 2 * Math.PI)
            {
                Dx = Length1 * Math.Cos(2 * Math.PI - Bear1);
                Dy = Length1 * Math.Sin(2 * Math.PI - Bear1);
                x1 = x0 + Dx;
                y1 = y0 - Dy;
            }



            if (Bear1 == 2 * Math.PI | Bear1 == 0)
            {
                Dx = Length1;
                Dy = 0;
                x1 = x0 + Dx;
                y1 = y0;
            }



            Double Val = x1;

            if (X_or_Y.ToUpper() == "Y") Val = y1;

            return Val;

        }



        public Double Bulge_for_arc_leader(Double x0, Double y0, Double x2, Double y2, Double x3, Double y3, Double Inaltimea)
        {

            Double x1 = X_Y_for_arc_leader(x0, y0, x2, y2, Inaltimea, "X");
            Double y1 = X_Y_for_arc_leader(x0, y0, x2, y2, Inaltimea, "Y");

            Double B1 = Functions.GET_Bearing_rad(x1, y1, x0, y0);

            Double B2 = Functions.GET_Bearing_rad(x1, y1, x3, y3);

            Double Angle1 = 0;

            if (B1 <= Math.PI / 2)
            {
                if (B2 >= B1 & B2 <= B1 + Math.PI)
                {
                    Angle1 = B2 - B1;
                    // - bulge
                }
                else
                {
                    Angle1 = 2 * Math.PI - (B2 - B1);
                    // + bulge
                }
            }

            if (B1 > Math.PI / 2 & B1 <= Math.PI)
            {
                if (B2 >= B1 & B2 <= B1 + Math.PI)
                {
                    Angle1 = B2 - B1;
                    // - bulge
                }
                else
                {
                    Angle1 = 2 * Math.PI - (B2 - B1);
                    // + bulge
                }
            }

            if (B1 > Math.PI & B1 <= 3 * Math.PI / 2)
            {
                if (B2 <= B1 & B2 >= B1 - Math.PI)
                {
                    Angle1 = -(B2 - B1);
                    // + bulge
                }
                else
                {
                    Angle1 = 2 * Math.PI + (B2 - B1);
                    // - bulge
                }
            }

            if (B1 > 3 * Math.PI / 2 & B1 <= 2 * Math.PI)
            {
                if (B2 <= B1 & B2 >= B1 - Math.PI)
                {
                    Angle1 = -(B2 - B1);
                    // + bulge
                }
                else
                {
                    Angle1 = 2 * Math.PI + (B2 - B1);
                    // - bulge
                }
            }

            if (Angle1 >= 2 * Math.PI) Angle1 = Angle1 - 2 * Math.PI;


            Double a2;
            Double Unghi_la_centru;


            if (Angle1 >= Math.PI / 2)
            {
                a2 = Angle1 - Math.PI / 2;
                Unghi_la_centru = Math.PI - 2 * a2;
            }
            else
            {
                a2 = Math.PI / 2 - Angle1;
                Unghi_la_centru = 2 * Math.PI - (Math.PI - 2 * a2);
            }


            Double Bulge1 = Math.Tan(Unghi_la_centru / 4);

            if (B1 <= Math.PI / 2)
            {
                if (B2 >= B1 & B2 <= B1 + Math.PI)
                {
                    Bulge1 = -Bulge1;
                }
            }

            if (B1 > Math.PI / 2 & B1 <= Math.PI)
            {
                if (B2 >= B1 & B2 <= B1 + Math.PI)
                {
                    Bulge1 = -Bulge1;
                }
            }

            if (B1 > Math.PI & B1 <= 3 * Math.PI / 2)
            {
                if (B2 <= B1 & B2 >= B1 - Math.PI)
                {
                }
                else
                {
                    Bulge1 = -Bulge1;
                }
            }

            if (B1 > 3 * Math.PI / 2 & B1 <= 2 * Math.PI)
            {

                if (B2 <= B1 & B2 >= B1 - Math.PI)
                {
                }
                else
                {
                    Bulge1 = -Bulge1;
                }
            }

            return Bulge1;

        }


    }

    class Jig_Class3 : DrawJig
    {
        Point3d Basept;
        String Continut;

        Double Inaltimea1;
        PromptPointResult PromptPointResult1;


        public PromptPointResult StartJig( Double TextH, string Content)
        {
            Continut = Content;
            Inaltimea1 = TextH;
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
            PromptPointResult1 = prompts.AcquirePoint("\nText insertion: ");

            if (PromptPointResult1.Value.IsEqualTo(Basept))
            {
                return SamplerStatus.NoChange;
            }
            else
            {
                Basept = PromptPointResult1.Value;
                return SamplerStatus.OK;
            }

        }

        protected override bool WorldDraw(Autodesk.AutoCAD.GraphicsInterface.WorldDraw draw)
        {
            Autodesk.AutoCAD.GraphicsInterface.TextStyle  TS1 =new Autodesk.AutoCAD.GraphicsInterface.TextStyle();
            TS1.TextSize = Inaltimea1;
            
            draw.Geometry.Text(Basept, Vector3d.ZAxis, Vector3d.XAxis,Continut,true,TS1);


            
            return true;
        }
        public Double GET_distanta_Double(Double x1, Double y1, Double x2, Double y2)
        {
            return Math.Pow(Math.Pow(x1 - x2, 2) + Math.Pow(y1 - y2, 2), 0.5);
        }





    }


}
