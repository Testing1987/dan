using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System.Collections.Generic;



namespace Alignment_mdi
{

    class AttInfo
    {
        private Point3d _pos;
        private Point3d _aln;
        private bool _aligned;

        public AttInfo(Point3d pos, Point3d aln, bool aligned)

        {
            _pos = pos;
            _aln = aln;
            _aligned = aligned;
        }

        public Point3d Position
        {
            set { _pos = value; }
            get { return _pos; }
        }


        public Point3d Alignment
        {
            set { _aln = value; }
            get { return _aln; }
        }

        public bool IsAligned
        {
            set { _aligned = value; }
            get { return _aligned; }
        }
    }



    class BlockJig_no_attrib : EntityJig
    {
        private Point3d _pos;
        private Dictionary<ObjectId, AttInfo> _attInfo;
        private Transaction _tr;
        private Polyline _poly2d;

        public BlockJig_no_attrib(Transaction tr, BlockReference br, Dictionary<ObjectId, AttInfo> attInfo, Polyline poly2d) : base(br)
        {
            _poly2d = poly2d;
            _pos = br.Position;
            _attInfo = attInfo;
            _tr = tr;

        }



        protected override bool Update()
        {
            BlockReference br = Entity as BlockReference;

            if (_poly2d != null)
            {
                br.Position = _poly2d.GetClosestPointTo(_pos, Vector3d.ZAxis, false);
            }
            else
            {
                br.Position = _pos;
            }

            if (br.AttributeCollection.Count != 0)
            {
                foreach (ObjectId id in br.AttributeCollection)
                {
                    DBObject obj = _tr.GetObject(id, OpenMode.ForRead);
                    AttributeReference attrib1 = obj as AttributeReference;

                    // Apply block transform to att def position
                    if (attrib1 != null)
                    {
                        attrib1.UpgradeOpen();
                        AttInfo ai = _attInfo[attrib1.ObjectId];
                        attrib1.Position = ai.Position.TransformBy(br.BlockTransform);

                        if (ai.IsAligned)
                        {
                            attrib1.AlignmentPoint = ai.Alignment.TransformBy(br.BlockTransform);
                        }

                        if (attrib1.IsMTextAttribute)
                        {
                            attrib1.UpdateMTextAttribute();
                        }
                    }
                }
            }
            return true;
        }



        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            JigPromptPointOptions opts = new JigPromptPointOptions("\nSelect insertion point:");
            opts.BasePoint = new Point3d(0, 0, 0);
            opts.UserInputControls = UserInputControls.NoZeroResponseAccepted;

            PromptPointResult ppr = prompts.AcquirePoint(opts);

            if (_pos == ppr.Value)
            {
                return SamplerStatus.NoChange;
            }

            _pos = ppr.Value;

            return SamplerStatus.OK;
        }



        public PromptStatus Run()
        {
            Document ThisDrawing = Application.DocumentManager.MdiActiveDocument;
            Editor Editor1 = ThisDrawing.Editor;
            PromptResult promptResult = Editor1.Drag(this);
            return promptResult.Status;
        }
    }



    public class Commands

    {
        [CommandMethod("BJ")]
        static public void BlockJigCmd()
        {
            Document ThisDrawing = Application.DocumentManager.MdiActiveDocument;
            Database db = ThisDrawing.Database;
            Editor ed = ThisDrawing.Editor;
            PromptStringOptions pso = new PromptStringOptions("\nEnter block name: ");

            PromptResult pr = ed.GetString(pso);
            if (pr.Status != PromptStatus.OK) return;

            Transaction trans1 = ThisDrawing.TransactionManager.StartTransaction();

            using (trans1)
            {
                BlockTable blockTable1 = (BlockTable)trans1.GetObject(db.BlockTableId, OpenMode.ForRead);
                if (!blockTable1.Has(pr.StringResult))
                {
                    ed.WriteMessage("\nBlock \"" + pr.StringResult + "\" not found.");
                    return;
                }

                BlockTableRecord btRecord = (BlockTableRecord)trans1.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                BlockTableRecord btr = (BlockTableRecord)trans1.GetObject(blockTable1[pr.StringResult], OpenMode.ForRead);

                // Block needs to be inserted to current space before
                // being able to append attribute to it

                BlockReference br = new BlockReference(new Point3d(), btr.ObjectId);
                btRecord.AppendEntity(br);
                trans1.AddNewlyCreatedDBObject(br, true);

                Dictionary<ObjectId, AttInfo> attInfo = new Dictionary<ObjectId, AttInfo>();

                if (btr.HasAttributeDefinitions)
                {
                    foreach (ObjectId id in btr)
                    {
                        DBObject obj = trans1.GetObject(id, OpenMode.ForRead);

                        AttributeDefinition ad = obj as AttributeDefinition;

                        if (ad != null && !ad.Constant)
                        {
                            AttributeReference ar = new AttributeReference();

                            ar.SetAttributeFromBlock(ad, br.BlockTransform);

                            ar.Position = ad.Position.TransformBy(br.BlockTransform);

                            if (ad.Justify != AttachmentPoint.BaseLeft)
                            {
                                ar.AlignmentPoint = ad.AlignmentPoint.TransformBy(br.BlockTransform);
                            }
                            if (ar.IsMTextAttribute)
                            {
                                ar.UpdateMTextAttribute();
                            }

                            ar.TextString = ad.TextString;

                            ObjectId arId = br.AttributeCollection.AppendAttribute(ar);
                            trans1.AddNewlyCreatedDBObject(ar, true);

                            // Initialize our dictionary with the ObjectId of
                            // the attribute reference + attribute definition info
                            attInfo.Add(arId, new AttInfo(ad.Position, ad.AlignmentPoint, ad.Justify != AttachmentPoint.BaseLeft));
                        }
                    }
                }

                // Run the jig
                BlockJig_no_attrib myJig = new BlockJig_no_attrib(trans1, br, attInfo, null);

                if (myJig.Run() != PromptStatus.OK)
                {
                    return;
                }

                // Commit changes if user accepted, otherwise discard
                trans1.Commit();
            }
        }
    }






}
