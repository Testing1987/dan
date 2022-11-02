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

    class Jig_draw_polyline : DrawJig
    {

        List<Point3d> lista_pt = new List<Point3d>();

        Point3d Punct2 = new Point3d();
        string mesaj_1 = "";

        PromptPointResult PromptPointResult1;

        public PromptPointResult StartJig(List<Point3d> _lista_pt, string mesaj)
        {

            mesaj_1 = mesaj;

            lista_pt = _lista_pt;

            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            PromptPointResult1 = (PromptPointResult)ed.Drag(this);
            do
            {
                if (PromptPointResult1.Status == PromptStatus.OK)
                {
                    return PromptPointResult1;
                }
            } while (PromptPointResult1.Status != PromptStatus.Cancel && PromptPointResult1.Status != PromptStatus.None);

            return PromptPointResult1;
        }


        protected override SamplerStatus Sampler(JigPrompts prompts)
        {

            JigPromptPointOptions prompt2 = new JigPromptPointOptions(mesaj_1);
            prompt2.UserInputControls = UserInputControls.NullResponseAccepted;

            PromptPointResult1 = prompts.AcquirePoint(prompt2);



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

            Polyline Poly1 = new Autodesk.AutoCAD.DatabaseServices.Polyline();


            for (int i = 0; i < lista_pt.Count; i++)
            {
                Poly1.AddVertexAt(i, new Point2d(lista_pt[i].X, lista_pt[i].Y), 0, 0, 0);
            }

            Poly1.AddVertexAt(lista_pt.Count, new Point2d(Punct2.X, Punct2.Y), 0, 0, 0);
            draw.Geometry.Polyline(Poly1, 0, lista_pt.Count);


            return true;
        }





    }

}
