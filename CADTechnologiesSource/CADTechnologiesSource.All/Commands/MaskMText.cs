using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using System.Windows;

namespace CADTechnologiesSource.All.Commands
{
    public class MaskMText
    {
        [CommandMethod("MASKMTEXT")]
        public void MaskMTexts()
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
                        MText mText = transaction.GetObject(objectId, OpenMode.ForRead) as MText;
                        if (mText != null)
                        {
                            if (mText.BackgroundFill == false)
                            {
                                mText.UpgradeOpen();
                                mText.BackgroundFill = true;
                                mText.UseBackgroundColor = true;
                                mText.BackgroundScaleFactor = 1.5;
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

        [CommandMethod("UNMASKMTEXT")]
        public void UnMaskMTexts()
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
                        MText mText = transaction.GetObject(objectId, OpenMode.ForRead) as MText;
                        if (mText != null)
                        {
                            if (mText.BackgroundFill == true)
                            {
                                mText.UpgradeOpen();
                                mText.BackgroundFill = false;
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
