using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;



namespace SampleTesting
{
    public class MovelLabelJig : EntityJig
    {
        Point3d newlocation;
        Point3d oldlocation;

        public MovelLabelJig(MText mtext1)
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
              new JigPromptPointOptions("Move label: ");
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

        [CommandMethod("moveLabel")]
        public static void CmdMoveLabel()
        {
            Editor ed = Application.DocumentManager
              .MdiActiveDocument.Editor;

            // select the label
            PromptEntityOptions peo =
              new PromptEntityOptions("\nSelect a label to move: ");
            peo.SetRejectMessage("\nLabels only");
            peo.AddAllowedClass(typeof(MText), false);
            PromptEntityResult per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            // start a transaction
            Database db = Application.DocumentManager
              .MdiActiveDocument.Database;
            using (Transaction trans = db.
              TransactionManager.StartTransaction())
            {
                // open the label object
                MText lbl = trans.GetObject(per.ObjectId, OpenMode.ForRead) as MText;

                // start the JIG drag
                MovelLabelJig jig = new MovelLabelJig(lbl);
                PromptResult pr = ed.Drag(jig);

                // if not OK, clean return
                if (pr.Status != PromptStatus.OK) return;

                // everything ok, update the location
                lbl.UpgradeOpen();
                lbl.Location = jig.Location;

                trans.Commit();
            }
        }
    }
}
