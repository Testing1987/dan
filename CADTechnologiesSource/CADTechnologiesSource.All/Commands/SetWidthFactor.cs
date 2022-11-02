using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Windows;

namespace CADTechnologiesSource.All.Commands
{
    public class SetWidthFactor
    {
        [CommandMethod("SetWidthFactor")]
        public void SetWidthFactorCommand()
        {
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database thisDatabase = thisDrawing.Database;
            Editor thisEditor = thisDrawing.Editor;
            using (Transaction thisTransaction = thisDrawing.TransactionManager.StartTransaction())
            {
                try
                {
                    PromptEntityOptions promptEntityOptions = new PromptEntityOptions(Environment.NewLine + "Select the MText you want to change.");
                    promptEntityOptions.SetRejectMessage(Environment.NewLine + "You did not select an MText.");
                    promptEntityOptions.AddAllowedClass(typeof(MText), true);

                    PromptEntityResult promptEntityResult = thisEditor.GetEntity(promptEntityOptions);
                    if (promptEntityResult.Status != PromptStatus.OK)
                    {
                        return;
                    }

                    MText mText = thisTransaction.GetObject(promptEntityResult.ObjectId, OpenMode.ForWrite) as MText;
                    string oldContents = mText.Contents;
                    string newContents = mText.Text;
                    //string newContents = "{\\W1;" + mText.Text + "}";
                    mText.Contents = newContents;
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                thisTransaction.Commit();
            }
        }
    }
}
