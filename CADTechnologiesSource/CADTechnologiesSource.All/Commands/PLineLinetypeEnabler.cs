using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System.Windows;

namespace CADTechnologiesSource.All.Commands
{
    public class PLineLinetypeEnabler
    {
        [CommandMethod("ENABLEPLINELINETYPEGENERATION")]
        public void EnablePolylineLinetypeGeneration()
        {
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database thisDatabase = thisDrawing.Database;
            Editor thisEditor = thisDrawing.Editor;
            using (DocumentLock documentLock = thisDrawing.LockDocument())
            {
                #region Model Space
                try
                {
                    using (Transaction trans1 = thisDrawing.TransactionManager.StartTransaction())
                    {
                        int count = 0;
                        BlockTable blockTable = trans1.GetObject(thisDatabase.BlockTableId, OpenMode.ForRead) as BlockTable;
                        BlockTableRecord blockTableRecord = trans1.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

                        foreach (ObjectId objectId in blockTableRecord)
                        {
                            Polyline polyline = trans1.GetObject(objectId, OpenMode.ForRead) as Polyline;
                            if (polyline != null)
                            {
                                switch (polyline.Plinegen)
                                {
                                    case false:
                                        polyline.UpgradeOpen();
                                        polyline.Plinegen = true;
                                        count++;
                                        break;

                                    default:
                                        break;
                                }
                            }

                            Polyline2d polyline2d = trans1.GetObject(objectId, OpenMode.ForRead) as Polyline2d;
                            if (polyline2d != null)
                            {
                                switch (polyline2d.LinetypeGenerationOn)
                                {
                                    case false:
                                        polyline2d.UpgradeOpen();
                                        polyline2d.LinetypeGenerationOn = true;
                                        break;

                                    default:
                                        break;
                                }
                            }
                        }
                        trans1.Commit();
                        if (count == 1)
                        {
                            thisEditor.WriteMessage($"\r\nLinetype Generation set to true for {count} Polyline in Model Space.");
                        }
                        else
                        {
                            thisEditor.WriteMessage($"\r\nLinetype Generation set to true for {count} Polylines in Model Space.");
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                #endregion

                #region Paper Space
                try
                {
                    using (Transaction trans1 = thisDrawing.TransactionManager.StartTransaction())
                    {
                        int count = 0;
                        DBDictionary dBDictionary = trans1.GetObject(thisDatabase.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                        foreach (DBDictionaryEntry dBDictionaryEntry in dBDictionary)
                        {
                            Layout layout = dBDictionaryEntry.Value.GetObject(OpenMode.ForRead) as Layout;
                            if (layout != null)
                            {
                                BlockTableRecord blockTableRecord = trans1.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite) as BlockTableRecord;
                                foreach (ObjectId objectId in blockTableRecord)
                                {
                                    Polyline polyline = trans1.GetObject(objectId, OpenMode.ForRead) as Polyline;
                                    if (polyline != null)
                                    {
                                        switch (polyline.Plinegen)
                                        {
                                            case false:
                                                polyline.UpgradeOpen();
                                                polyline.Plinegen = true;
                                                count++;
                                                break;

                                            default:
                                                break;
                                        }
                                    }

                                    Polyline2d polyline2d = trans1.GetObject(objectId, OpenMode.ForRead) as Polyline2d;
                                    if (polyline2d != null)
                                    {
                                        switch (polyline2d.LinetypeGenerationOn)
                                        {
                                            case false:
                                                polyline2d.UpgradeOpen();
                                                polyline2d.LinetypeGenerationOn = true;
                                                break;

                                            default:
                                                break;
                                        }
                                    }
                                }
                            }
                        }
                        trans1.Commit();
                        if (count == 1)
                        {
                            thisEditor.WriteMessage($"\r\nLinetype Generation set to true for {count} Polyline in Paper Space.");
                        }
                        else
                        {
                            thisEditor.WriteMessage($"\r\nLinetype Generation set to true for {count} Polylines in Paper Space.");
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                #endregion
            }
        }
    }
}
