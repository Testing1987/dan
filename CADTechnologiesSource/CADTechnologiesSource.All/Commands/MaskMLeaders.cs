using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using System.Windows;

namespace CADTechnologiesSource.All.Commands
{
    public class MaskMLeaders
    {
        [CommandMethod("MASKMLEADERS")]
        public void MaskMultiLeaders()
        {
            Document thisDocument = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database thisDatabase = thisDocument.Database;
            try
            {
                using (Transaction transaction = thisDocument.Database.TransactionManager.StartTransaction())
                {
                    BlockTable blockTable = transaction.GetObject(thisDatabase.BlockTableId, OpenMode.ForWrite) as BlockTable;
                    BlockTableRecord blockTableRecord = transaction.GetObject(thisDatabase.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                    foreach (ObjectId objectId in blockTableRecord)
                    {
                        MLeader mLeader = transaction.GetObject(objectId, OpenMode.ForWrite) as MLeader;
                        if (mLeader != null)
                        {
                            MText mText = mLeader.MText;
                            MText newMText = mText;

                            if (newMText.BackgroundFill == false)
                            {
                                newMText.BackgroundFill = true;
                                newMText.UseBackgroundColor = true;
                                newMText.BackgroundScaleFactor = 1.5;
                                mLeader.MText = newMText;
                            }
                        }
                    }
                    transaction.Commit();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
