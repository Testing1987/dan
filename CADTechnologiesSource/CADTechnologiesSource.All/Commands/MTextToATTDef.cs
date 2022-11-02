using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;

namespace CADTechnologiesSource.All.Commands
{
    public class MTextToATTDef
    {
        [CommandMethod("ConvertSurveyStation")]
        public void ConvertMTextToATTDef()
        {
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database thisDatabase = thisDrawing.Database;
            Editor thisEditor = thisDrawing.Editor;
            using (Transaction transaction = thisDrawing.TransactionManager.StartTransaction())
            {
                try
                {
                    // Create a new PromptSelectionOptions...
                    PromptSelectionOptions pso = new PromptSelectionOptions();

                    // Ask the user to select items
                    PromptSelectionResult psr = thisDrawing.Editor.GetSelection(pso);
                    if (psr.Status == PromptStatus.OK)
                    {
                        BlockTable blockTable = transaction.GetObject(thisDrawing.Database.BlockTableId, OpenMode.ForWrite) as BlockTable;

                        // Create a new selection set and assign the user's selection to it
                        SelectionSet selectionSet = psr.Value;

                        // iterate through each object in the selection and...
                        foreach (SelectedObject selectedObject in selectionSet)
                        {
                            if (selectedObject != null)
                            {
                                // create an entity for each one
                                Entity entity = transaction.GetObject(selectedObject.ObjectId, OpenMode.ForWrite) as Entity;
                                if (entity != null)
                                {

                                    // Check to make sure the entity is an MText
                                    if (entity is MText)
                                    {
                                        // make a nex MText out of the selected MText's values
                                        MText selectedMText = (MText)transaction.GetObject(entity.ObjectId, OpenMode.ForWrite);

                                        // Check to make sure we have the block named SURVEY_STATION_BLOCK
                                        if (blockTable.Has("SURVEY_STATION_BLOCK"))
                                        {
                                            // create a new blockTableRecord in model space based on the SURVEY_STATION_BLOCK
                                            BlockTableRecord modelSpace = (BlockTableRecord)transaction.GetObject(thisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite);
                                            BlockTableRecord blockTableRecord = (BlockTableRecord)transaction.GetObject(blockTable["SURVEY_STATION_BLOCK"], OpenMode.ForWrite);

                                            // Create a block reference of SURVEY_STATION_BLOCK and assign it the same location as the selected MText it is going to replace,
                                            // and assign it the object id of the blockTableRecord created for it
                                            BlockReference surveyStationBlock = new BlockReference(selectedMText.Location, blockTableRecord.ObjectId);

                                            // Append the new block into model space
                                            modelSpace.AppendEntity(surveyStationBlock);

                                            //Add the new block reference to the transaction
                                            transaction.AddNewlyCreatedDBObject(surveyStationBlock, true);

                                            // Get the attribute collection of the new block reference
                                            AttributeCollection attributeCollection = surveyStationBlock.AttributeCollection;

                                            // Get a new enumerator to enumerate through the blockTableRecord
                                            BlockTableRecordEnumerator btr_enum = blockTableRecord.GetEnumerator();

                                            // Look through the block table record for an attribute definition and assign it to the block referernce
                                            while (btr_enum.MoveNext())
                                            {
                                                Entity entity1 = (Entity)transaction.GetObject(btr_enum.Current, OpenMode.ForWrite);
                                                if (entity1 is AttributeDefinition)
                                                {
                                                    AttributeDefinition attributeDefinition = (AttributeDefinition)entity1;
                                                    AttributeReference attributeReference = new AttributeReference();
                                                    attributeReference.SetAttributeFromBlock(attributeDefinition, surveyStationBlock.BlockTransform);

                                                    if (attributeReference != null)
                                                    {
                                                        attributeCollection.AppendAttribute(attributeReference);
                                                        transaction.AddNewlyCreatedDBObject(attributeReference, true);
                                                    }
                                                }
                                            }

                                            // Find the newly created attribute reference and assign it the value of the MText which is being replaced
                                            foreach (ObjectId attID in attributeCollection)
                                            {
                                                AttributeReference attributeReference = (AttributeReference)transaction.GetObject(attID, OpenMode.ForWrite);
                                                if (attributeReference.Tag == "SURVEY_STATION")
                                                {
                                                    attributeReference.TextString = selectedMText.Text;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            MessageBox.Show("You Don't Have The block named SURVEY_STATION_BLOCK in this drawing.");
                                        }
                                    }

                                    //if the object is DBText
                                    if (entity is DBText)
                                    {
                                        DBText selectedDBText = (DBText)transaction.GetObject(entity.ObjectId, OpenMode.ForWrite);
                                        if (blockTable.Has("SURVEY_STATION_BLOCK"))
                                        {
                                            BlockTableRecord modelSpace = (BlockTableRecord)transaction.GetObject(thisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite);
                                            BlockTableRecord blockTableRecord = (BlockTableRecord)transaction.GetObject(blockTable["SURVEY_STATION_BLOCK"], OpenMode.ForWrite);

                                            BlockReference surveyStationBlock = new BlockReference(selectedDBText.AlignmentPoint, blockTableRecord.ObjectId);
                                            modelSpace.AppendEntity(surveyStationBlock);
                                            transaction.AddNewlyCreatedDBObject(surveyStationBlock, true);

                                            AttributeCollection attributeCollection = surveyStationBlock.AttributeCollection;
                                            BlockTableRecordEnumerator btr_enum = blockTableRecord.GetEnumerator();
                                            while (btr_enum.MoveNext())
                                            {
                                                Entity entity1 = (Entity)transaction.GetObject(btr_enum.Current, OpenMode.ForWrite);
                                                if (entity1 is AttributeDefinition)
                                                {
                                                    AttributeDefinition attributeDefinition = (AttributeDefinition)entity1;
                                                    AttributeReference attributeReference = new AttributeReference();
                                                    attributeReference.SetAttributeFromBlock(attributeDefinition, surveyStationBlock.BlockTransform);

                                                    if (attributeReference != null)
                                                    {
                                                        attributeCollection.AppendAttribute(attributeReference);
                                                        transaction.AddNewlyCreatedDBObject(attributeReference, true);
                                                    }
                                                }
                                            }

                                            foreach (ObjectId attID in attributeCollection)
                                            {
                                                AttributeReference attributeReference = (AttributeReference)transaction.GetObject(attID, OpenMode.ForWrite);
                                                if (attributeReference.Tag == "SURVEY_STATION")
                                                {
                                                    attributeReference.TextString = selectedDBText.TextString;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            MessageBox.Show("You Don't Have The block named SURVEY_STATION_BLOCK in this drawing.");
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                transaction.Commit();
                // Set the user's screen focus to AutoCAD and access the needed database information
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            }
        }
    }
}
