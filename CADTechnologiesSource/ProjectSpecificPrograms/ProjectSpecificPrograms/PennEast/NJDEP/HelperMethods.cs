using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjectSpecificPrograms.PennEast.NJDEP
{
    public class HelperMethods
    {
        public void FindBlock(BlockReference blockReference, string blockName, string tag)
        {
            Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor Editor1 = ThisDrawing.Editor;

            using (DocumentLock Lock1 = ThisDrawing.LockDocument())
            {
                using (Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    BlockTable BlockTable = Trans1.GetObject(ThisDrawing.Database.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTR = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;

                    foreach (ObjectId id1 in BTR)
                    {
                        blockReference = Trans1.GetObject(id1, OpenMode.ForRead) as BlockReference;
                        if (blockReference != null && blockReference.Name == blockName)
                        {
                            AttributeCollection AtrCol1 = blockReference.AttributeCollection;
                            foreach (ObjectId AtrId1 in AtrCol1)
                            {
                                AttributeReference AtrRef1 = Trans1.GetObject(AtrId1, OpenMode.ForRead) as AttributeReference;
                                if (AtrRef1 != null && AtrRef1.Tag == tag)
                                {
                                    BlockAligner.tag_FHA = AtrRef1.Tag;
                                    BlockAligner.value_FHA = AtrRef1.TextString;
                                    BlockAligner.FHA_Header_Position = blockReference.Position;
                                    BlockAligner.Last_Placed_FHA = blockReference.Position;
                                }
                            }
                        }
                    }
                    Trans1.Commit();
                }
            }
        }
    }
}
