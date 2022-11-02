using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlockManager.Core.CoreLogic
{
    class _123
    {
        public void testmethod()
        {
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.CurrentDocument;
            Document destinationDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.CurrentDocument;
            Database thisDatabase = thisDrawing.Database;
            Database destinationDatabase = destinationDrawing.Database;
            ObjectIdCollection objectIdCollection = new ObjectIdCollection();
            DocumentCollection docManager = Application.DocumentManager;

            using (Transaction destinationTransaction = destinationDatabase.TransactionManager.StartTransaction())
            {
                // Open the Block table for read
                BlockTable destinationBlockTable;
                destinationBlockTable = destinationTransaction.GetObject(destinationDatabase.BlockTableId, OpenMode.ForRead) as BlockTable;

                // Open the Block table record Model space for read
                BlockTableRecord destinationModelSpace;
                destinationModelSpace = destinationTransaction.GetObject(destinationBlockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;
                                                   


                // Clone the objects to the new database
                IdMapping map = new IdMapping();
                thisDatabase.WblockCloneObjects(objectIdCollection, destinationModelSpace.ObjectId, map, DuplicateRecordCloning.Ignore, false);
                                      
                // Save the copied objects to the database
                destinationTransaction.Commit();
            }
        }
    }
}
