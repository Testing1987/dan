using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

namespace CADTechnologiesSource.All.Commands
{
    public class LeftJustified
    {
        [CommandMethod("LEFTJUSTIFYTEXT")]
        public void LeftJustifyTextAndMText()
        {
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database thisDatabase = thisDrawing.Database;
            Editor thisEditor = thisDrawing.Editor;

            using (DocumentLock docLock = thisDrawing.LockDocument())
            {
                try
                {
                    using (Transaction thisTransaction = thisDatabase.TransactionManager.StartTransaction())
                    {
                        //model space
                        #region Model Space
                        BlockTable blockTable = thisTransaction.GetObject(thisDatabase.BlockTableId, OpenMode.ForRead) as BlockTable;
                        BlockTableRecord modelSpace = thisTransaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

                        //MText
                        foreach (ObjectId objectId in modelSpace)
                        {
                            MText mText = thisTransaction.GetObject(objectId, OpenMode.ForRead) as MText;
                            if (mText != null && mText.Attachment != AttachmentPoint.TopLeft)
                            {
                                mText.UpgradeOpen();
                                mText.Attachment = AttachmentPoint.TopLeft;
                            }
                        }
                        //dbText
                        foreach (ObjectId objectId in modelSpace)
                        {
                            DBText dBext = thisTransaction.GetObject(objectId, OpenMode.ForRead) as DBText;
                            if (dBext != null && dBext.Justify != AttachmentPoint.MiddleLeft)
                            {
                                dBext.UpgradeOpen();
                                dBext.Justify = AttachmentPoint.MiddleLeft;
                            }
                        }
                        #endregion


                        #region Paper Space
                        //paper space (all layouts)
                        DBDictionary dBDictionary = thisTransaction.GetObject(thisDatabase.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                        foreach (DBDictionaryEntry dBDictionaryEntry in dBDictionary)
                        {
                            Layout layout = dBDictionaryEntry.Value.GetObject(OpenMode.ForRead) as Layout;
                            if (layout != null && layout.LayoutName != "Model" && layout.TabOrder > 0)
                            {
                                BlockTableRecord paperSpace = thisTransaction.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite) as BlockTableRecord;

                                //MText
                                foreach (ObjectId objectId in paperSpace)
                                {
                                    MText mText = thisTransaction.GetObject(objectId, OpenMode.ForRead) as MText;
                                    if (mText != null && mText.Attachment != AttachmentPoint.TopLeft)
                                    {
                                        mText.UpgradeOpen();
                                        mText.Attachment = AttachmentPoint.TopLeft;
                                    }
                                }

                                //dbText 
                                foreach (ObjectId objectId in paperSpace)
                                {
                                    DBText dBext = thisTransaction.GetObject(objectId, OpenMode.ForRead) as DBText;
                                    if (dBext != null && dBext.Justify != AttachmentPoint.MiddleLeft)
                                    {
                                        dBext.UpgradeOpen();
                                        dBext.Justify = AttachmentPoint.MiddleLeft;
                                    }
                                }
                            }
                        } 
                        #endregion


                        thisTransaction.Commit();
                    }
                }
                catch (System.Exception ex)
                {
                    throw;
                }
            }
        }
    }
}
