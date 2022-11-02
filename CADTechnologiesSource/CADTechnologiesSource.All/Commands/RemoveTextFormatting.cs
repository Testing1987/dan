using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

namespace CADTechnologiesSource.All.Commands
{
    public class RemoveTextFormatting
    {
        [CommandMethod("CLEARTEXTFORMATTING")]
        public void ClearTxtFormatting()
        {
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database thisDatabase = thisDrawing.Database;
            Editor thisEditor = thisDrawing.Editor;

            using (DocumentLock docLock = thisDrawing.LockDocument())
            {
                PromptEntityOptions peo = new PromptEntityOptions("\nSelect an MTEXT object");
                peo.SetRejectMessage("\r\nSelect an MTEXT.");
                peo.AddAllowedClass(typeof(MText), true);

                PromptEntityResult per = thisEditor.GetEntity(peo);

                if (per.Status != PromptStatus.OK)
                {
                    return;
                }
                else
                {
                    try
                    {
                        using (Transaction thisTransaction = thisDatabase.TransactionManager.StartTransaction())
                        {
                            MText mText = thisTransaction.GetObject(per.ObjectId, OpenMode.ForRead) as MText;
                            if (mText != null)
                            {
                                TextEditor textEditor = TextEditor.CreateTextEditor(mText);
                                if (textEditor != null)
                                {
                                    mText.UpgradeOpen();
                                    textEditor.SelectAll();
                                    textEditor.Selection.RemoveAllFormatting();
                                    textEditor.Close(TextEditor.ExitStatus.ExitSave);
                                    thisTransaction.Commit();
                                }
                            }
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
}
