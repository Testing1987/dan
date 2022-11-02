﻿using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Text;

namespace CADTechnologiesSource.All.AutoCADHelpers
{
    public static class SynchronizeBlockAttributes
    {
        public static void SynchronizeAttributes(this BlockTableRecord target)
        {
            if (target == null)
                throw new ArgumentNullException("btr");

            Transaction tr = target.Database.TransactionManager.TopTransaction;
            //if (tr == null)
            //    throw new AcRx.Exception(ErrorStatus.NoActiveTransactions);


            RXClass attDefClass = RXClass.GetClass(typeof(AttributeDefinition));
            List<AttributeDefinition> attDefs = new List<AttributeDefinition>();
            foreach (ObjectId id in target)
            {
                if (id.ObjectClass == attDefClass)
                {
                    AttributeDefinition attDef = (AttributeDefinition)tr.GetObject(id, OpenMode.ForRead);
                    attDefs.Add(attDef);
                }
            }
            foreach (ObjectId id in target.GetBlockReferenceIds(true, false))
            {
                BlockReference br = (BlockReference)tr.GetObject(id, OpenMode.ForWrite);
                br.ResetAttributes(attDefs);
            }
            if (target.IsDynamicBlock)
            {
                foreach (ObjectId id in target.GetAnonymousBlockIds())
                {
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(id, OpenMode.ForRead);
                    foreach (ObjectId brId in btr.GetBlockReferenceIds(true, false))
                    {
                        BlockReference br = (BlockReference)tr.GetObject(brId, OpenMode.ForWrite);
                        br.ResetAttributes(attDefs);
                    }
                }
            }
        }

        private static void ResetAttributes(this BlockReference br, List<AttributeDefinition> attDefs)
        {
            TransactionManager tm = br.Database.TransactionManager;
            Dictionary<string, string> attValues = new Dictionary<string, string>();
            foreach (ObjectId id in br.AttributeCollection)
            {
                if (!id.IsErased)
                {
                    AttributeReference attRef = (AttributeReference)tm.GetObject(id, OpenMode.ForWrite);
                    attValues.Add(attRef.Tag, attRef.TextString);
                    attRef.Erase();
                }
            }
            foreach (AttributeDefinition attDef in attDefs)
            {
                AttributeReference attRef = new AttributeReference();
                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                if (attValues.ContainsKey(attDef.Tag))
                {
                    attRef.TextString = attValues[attDef.Tag.ToUpper()];
                }
                br.AttributeCollection.AppendAttribute(attRef);
                tm.AddNewlyCreatedDBObject(attRef, true);
            }
        }
    }
}
