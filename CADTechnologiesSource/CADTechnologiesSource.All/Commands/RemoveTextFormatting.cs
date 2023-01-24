using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System.Windows;

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
                PromptSelectionResult psr = thisEditor.GetSelection();
                if (psr.Status != PromptStatus.OK)
                {
                    return;
                }
                else
                {
                    SelectionSet selectedItems = psr.Value;
                    try
                    {
                        using (Transaction thisTransaction = thisDatabase.TransactionManager.StartTransaction())
                        {
                            foreach (SelectedObject item in selectedItems)
                            {
                                if (item != null)
                                {
                                    RXClass mtextClass = RXObject.GetClass(typeof(MText));
                                    ObjectId objectId = item.ObjectId;
                                    if (objectId.ObjectClass == mtextClass)
                                    {
                                        MText mText = thisTransaction.GetObject(item.ObjectId, OpenMode.ForRead) as MText;
                                        if (mText != null)
                                        {
                                            TextEditor textEditor = TextEditor.CreateTextEditor(mText);
                                            if (textEditor != null)
                                            {
                                                mText.UpgradeOpen();
                                                textEditor.SelectAll();
                                                textEditor.Selection.RemoveAllFormatting();
                                                textEditor.Close(TextEditor.ExitStatus.ExitSave);
                                            }
                                        }
                                    }
                                }
                            }
                            thisTransaction.Commit();
                        }
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
        }
    }
}
