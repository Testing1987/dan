using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Threading;
using System.Threading.Tasks;

namespace CADTechnologiesSource.All.Commands
{
    public class ScrambledEggs
    {
        [CommandMethod("ScrambledEggs")]
        public void ScrambleEggs()
        {
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database thisDatabase = thisDrawing.Database;
            using (Transaction transaction = thisDrawing.TransactionManager.StartTransaction())
            {
                try
                {
                    Random random = new Random();
                    Random ispltrng = new Random();
                    Random isoffrng = new Random();
                    Random isfrzrng = new Random();
                    LayerTable layerTable = AutoCADHelpers.DatabaseHelpers.GetLayerTableForWrite(transaction, thisDatabase);
                    foreach (ObjectId objectId in layerTable)
                    {
                        LayerTableRecord layerTableRecord = objectId.GetObject(OpenMode.ForWrite) as LayerTableRecord;
                        if(layerTableRecord.Id != thisDatabase.Clayer)
                        {
                            layerTableRecord.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, (short)random.Next(1, 257));
                            layerTableRecord.IsOff = isoffrng.Next(0, 2) > 0;
                            layerTableRecord.IsFrozen = isfrzrng.Next(0, 2) > 0;
                            layerTableRecord.IsPlottable = ispltrng.Next(0, 2) > 0;
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                transaction.Commit();
            }
        }
    }
}
