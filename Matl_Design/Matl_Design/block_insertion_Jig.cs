using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace Alignment_mdi
{
    public class BlockJig : EntityJig
    {
        protected Point3d position;
        protected Polyline _poly2d;
        protected BlockReference _br;
        protected string message;

        public BlockJig(BlockReference br, string message, Polyline poly1) : base(br)
        {
            if (br == null) throw new ArgumentNullException("br");
            _br = br;
            _poly2d = poly1;
            position = br.Position;
            this.message = message;
        }

        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            var jppo = new JigPromptPointOptions(message);
            jppo.UserInputControls = UserInputControls.Accept3dCoordinates;
            var ppr = prompts.AcquirePoint(jppo);
            if (position.DistanceTo(ppr.Value) < Tolerance.Global.EqualPoint)
                return SamplerStatus.NoChange;
            position = ppr.Value;
            return SamplerStatus.OK;
        }

        protected override bool Update()
        {
            if (_poly2d != null)
            {
                _br.Position = _poly2d.GetClosestPointTo(position, Vector3d.ZAxis, false);
            }
            else
            {
                _br.Position = position;
            }
            return true;
        }
    }

    public class BlockAttribJig : BlockJig
    {
        private Dictionary<AttributeReference, TextInfo> attRefInfos;
        private Transaction Trans1;
        public BlockAttribJig(BlockReference br, string message, Polyline poly1, Dictionary<AttributeReference, TextInfo> attInfos) : base(br, message, poly1)
        {
            attRefInfos = attInfos;
            Trans1 = br.Database.TransactionManager.TopTransaction;
        }

        protected override bool Update()
        {
            base.Update();
            foreach (var entry in attRefInfos)
            {
                var att = entry.Key;
                var info = entry.Value;
                att.Position = info.Position.TransformBy(_br.BlockTransform);
                if (info.IsAligned)
                {
                    att.AlignmentPoint = info.Alignment.TransformBy(_br.BlockTransform);
                    att.AdjustAlignment(_br.Database);
                }
                if (att.IsMTextAttribute)
                {
                    att.UpdateMTextAttribute();
                }
            }
            return true;
        }
    }

    public class TextInfo
    {
        public Point3d Position { get; private set; }
        public Point3d Alignment { get; private set; }
        public bool IsAligned { get; private set; }
        public double Rotation { get; private set; }
        public TextInfo(DBText text)
        {
            Position = text.Position;
            IsAligned = text.Justify != AttachmentPoint.BaseLeft;
            Alignment = text.AlignmentPoint;
            Rotation = text.Rotation;
        }
    }


    public class jig_actions
    {
        static public void insert_block(ref BlockReference br, string bl_name, string layer1, Polyline poly1)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Application.DocumentManager.MdiActiveDocument;
            using (Transaction Trans1 = ThisDrawing.Database.TransactionManager.StartTransaction())
            {
                BlockTable bt = Trans1.GetObject(ThisDrawing.Database.BlockTableId, OpenMode.ForRead) as BlockTable;
                
                if (bt.Has(bl_name) == false)
                {
                    Autodesk.AutoCAD.ApplicationServices.Core.Application.ShowAlertDialog("Block " + bl_name + " not found");
                    return;
                }

                Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                // create a block reference and add it to the current space
                br = new BlockReference(Point3d.Origin, bt[bl_name]);
                br.TransformBy(ThisDrawing.Editor.CurrentUserCoordinateSystem);
                br.Layer = layer1;
                br.ColorIndex = 256;
                BTrecord.AppendEntity(br);
                Trans1.AddNewlyCreatedDBObject(br, true);

                // add the attribute references to the block reference and fill a dictionary with attributes text infos
                var btr = (BlockTableRecord)Trans1.GetObject(bt[bl_name], OpenMode.ForRead);
                var attInfos = new Dictionary<AttributeReference, TextInfo>();
                foreach (ObjectId id in btr)
                {
                    if (id.ObjectClass.Name == "AcDbAttributeDefinition")
                    {
                        AttributeDefinition attDef = (AttributeDefinition)Trans1.GetObject(id, OpenMode.ForRead);
                        AttributeReference attRef = new AttributeReference();
                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                        br.AttributeCollection.AppendAttribute(attRef);
                        Trans1.AddNewlyCreatedDBObject(attRef, true);
                        attInfos.Add(attRef, new TextInfo(attDef));
                    }
                }

                BlockAttribJig jig = new BlockAttribJig(br, "\nInsertion point: ", poly1, attInfos);
                PromptResult result = ThisDrawing.Editor.Drag(jig);
                // erase the block reference if the user cancelled
                if (result.Status == PromptStatus.Cancel)
                {
                    br.Erase();
                }
                Trans1.Commit();
            }
        }
    }

}
