using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Text;

namespace CADTechnologiesSource.All.AnnotationHelpers
{
    #region Delegates

    public delegate MText MTextDelegate(Transaction transaction, Database database,
                 BlockTable blocktable, BlockTableRecord blocktablerecord);

    #endregion
    public class AnnotationDelegates
    {
        /// <summary>
        /// Returns an MText entity that can be used in a lambda method.
        /// </summary>
        /// <param name="mtextdelegate"></param>
        /// <returns></returns>
        public static MText CreateMTextDelegate(MTextDelegate mtextdelegate)
        {
            MText mt;
            var document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var database = document.Database;
            using (var transaction = document.TransactionManager.StartTransaction())
            {
                BlockTable blocktable = transaction.GetObject(database.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord blockTableRecord = transaction.GetObject(blocktable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                mt = mtextdelegate(transaction, database, blocktable, blockTableRecord);
            }
            return mt;
        }
    }
}
