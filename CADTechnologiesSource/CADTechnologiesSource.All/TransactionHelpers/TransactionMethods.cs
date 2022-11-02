using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Text;

namespace CADTechnologiesSource.All.TransactionHelpers
{
    public class TransactionMethods
    {
        /// <summary>
        /// Opens a transaction in the current drawing and allows you to pass in an <see cref="Action"/> that will modify the block table or a record.
        /// </summary>
        /// <param name="action">The calling method</param>
        public static void ModifyDrawingWithinTransaction(Action<Transaction, Database, BlockTable, BlockTableRecord> action)
        {
            var document = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            var database = document.Database;
            using (var transaction = document.TransactionManager.StartTransaction())
            {
                BlockTable blocktable = transaction.GetObject(database.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord blockTableRecord = transaction.GetObject(blocktable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                action(transaction, database, blocktable, blockTableRecord);
                transaction.Commit();
            }
        }

        /// <summary>
        /// Opens a transaction in the current drawing and allows you to execute a generic action that does not modify the block table.
        /// </summary>
        /// <param name="action">The calling method</param>
        public static void DoActionWithinTransaction(Action<Transaction, Database> action)
        {
            var document = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            var database = document.Database;
            using (var transaction = document.TransactionManager.StartTransaction())
            {
                action(transaction, database);
                transaction.Commit();
            }
        }
    }
}
