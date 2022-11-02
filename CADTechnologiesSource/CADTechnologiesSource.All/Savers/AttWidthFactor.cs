using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;

namespace CADTechnologiesSource.All.Savers
{
    public class AttWidthFactor
    {
        [CommandMethod("BlockAttWidth")]
        public void SetMLineBlockWidthFactor()
        {
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor editor = thisDrawing.Editor;
            using (DocumentLock lock1 = thisDrawing.LockDocument())
            {
                using (Transaction trans1 = thisDrawing.TransactionManager.StartTransaction())
                {
                    PromptEntityOptions peo = new PromptEntityOptions("\nSelect a multiline block attribute");
                    peo.SetRejectMessage("Please select a multiline block attribute.");
                    peo.AddAllowedClass(typeof(BlockReference), true);

                    PromptEntityResult per = editor.GetEntity(peo);

                    if (per.Status == PromptStatus.OK)
                    {
                        try
                        {
                            BlockReference block = trans1.GetObject(per.ObjectId, OpenMode.ForWrite) as BlockReference;
                            AttributeCollection attributes = block.AttributeCollection;

                            foreach (ObjectId attribute in attributes)
                            {
                                AttributeReference atr1 = trans1.GetObject(attribute, OpenMode.ForWrite) as AttributeReference;
                                if (atr1 != null && atr1.Tag == "DESCR")
                                {
                                    atr1.TextString = "{\\W0.65;" + atr1.MTextAttribute.Contents + "}";
                                }
                            }
                        }
                        catch (System.Exception)
                        {
                            MessageBox.Show("That object can't be changed.", "Unable to change width", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                    trans1.Commit();
                }
            }
        }
    }
}
