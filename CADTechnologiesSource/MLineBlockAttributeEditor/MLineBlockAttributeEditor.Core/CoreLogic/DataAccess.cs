using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using CADTechnologiesSource.All.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;

namespace MLineBlockAttributeEditor.Core.CoreLogic
{
    public class DataAccess
    {
        #region Blocks
        
        /// <summary>
        /// Create a <see cref="BlockModel"/> that inherits the properties of a block selected by the user in AutoCAD.
        /// </summary>
        /// <returns></returns>
        public ObservableCollection<BlockAttributeModel> BuildBlockData(PromptEntityResult per, Document thisDrawing, Database thisDatabase)
        {
            // Set up a BlockModel to populate from a user selection
            ObservableCollection<BlockAttributeModel> output = new ObservableCollection<BlockAttributeModel>();

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            // Ensure the selection is okay
            if (per.Status == PromptStatus.OK)
            {
                // Lock the drawing
                using (DocumentLock docLock = thisDrawing.LockDocument())
                {
                    try
                    {
                        // Start a transaction
                        using (Transaction transaction = thisDrawing.TransactionManager.StartTransaction())
                        {
                            // Open the Layer Table
                            LayerTable layerTable = transaction.GetObject(thisDatabase.LayerTableId, OpenMode.ForRead) as LayerTable;

                            // Create a block out of the user's selection
                            BlockReference selectedBlock = transaction.GetObject(per.ObjectId, OpenMode.ForRead) as BlockReference;

                            // Read all attributes in the block...
                            AttributeCollection attributeCollection = selectedBlock.AttributeCollection;

                            // Count how many are found...
                            var attributeQuantity = attributeCollection.Count;
                            var mlineAttributeQuantity = 0;
                            // and set up some containers for single line attributes so we can display to the user which attributes won't be shown in the application
                            var singleLineAttributeQuantity = 0;
                            List<string> singleLineAttributeReferencesList = new List<string>();


                            // and create a reference for each one found
                            foreach (ObjectId attributeID in attributeCollection)
                            {
                                AttributeReference attributeReference = (AttributeReference)transaction.GetObject(attributeID, OpenMode.ForRead);
                                // Make sure we only get multi-line attributes
                                if (attributeReference.IsMTextAttribute == true)
                                {
                                    mlineAttributeQuantity++;
                                    // Get the attribute's layer and linetype
                                    LinetypeTableRecord lineTypeRecord = transaction.GetObject(attributeReference.LinetypeId, OpenMode.ForRead) as LinetypeTableRecord;

                                    // and create a new blockAttributeModel for each attribute found in the selection.
                                    BlockAttributeModel blockAttributeModel = new BlockAttributeModel();
                                    blockAttributeModel.Tag = attributeReference.Tag;
                                    blockAttributeModel.Layer = attributeReference.Layer;
                                    #region Color

                                    //Get the attribute color by creating an autocad color, setting it to the value of the color found on the layer, and converting it as necessary.
                                    Autodesk.AutoCAD.Colors.Color color1 = attributeReference.Color;
                                    string color_string = Convert.ToString(color1);

                                    //Check for TrueColor by looking for a comma in the value and returning an RGB.
                                    if (color_string.ToLower().Contains(",") == true)
                                    {
                                        int idx1 = color_string.IndexOf(",", 0);
                                        byte R = Convert.ToByte(color_string.Substring(0, idx1));
                                        int idx2 = color_string.IndexOf(",", idx1 + 1);
                                        byte G = Convert.ToByte(color_string.Substring(idx1 + 1, idx2 - idx1 - 1));
                                        byte B = Convert.ToByte(color_string.Substring(idx2 + 1, color_string.Length - idx2 - 1));
                                        color1 = Autodesk.AutoCAD.Colors.Color.FromRgb(R, G, B);

                                        //Return the truecolor value.
                                        blockAttributeModel.Color = color1;
                                    }
                                    else
                                    {
                                        //Return the index color.
                                        blockAttributeModel.Color = color1;
                                    }
                                    #endregion
                                    blockAttributeModel.Linetype = lineTypeRecord.Name;
                                    blockAttributeModel.Lineweight = attributeReference.LineWeight;
                                    blockAttributeModel.LineWeightString = attributeReference.LineWeight.ToString();
                                    blockAttributeModel.Justification = attributeReference.MTextAttribute.Attachment;
                                    blockAttributeModel.JustificationString = attributeReference.MTextAttribute.Attachment.ToString();
                                    blockAttributeModel.Height = attributeReference.Height;
                                    blockAttributeModel.Rotation = attributeReference.Rotation;
                                    blockAttributeModel.WidthFactor = attributeReference.WidthFactor;
                                    blockAttributeModel.TextString = attributeReference.TextString;
                                    blockAttributeModel.MtextAttributeContent = attributeReference.MTextAttribute.Contents;
                                    //blockAttributeModel.BackgroundFill = attributeReference.MTextAttribute.BackgroundFill;
                                    //blockAttributeModel.UseBackgroundFill = attributeReference.MTextAttribute.UseBackgroundColor;

                                    ////you cannot get the background fill color if BackgroundFIll is false.
                                    //if (blockAttributeModel.BackgroundFill == true)
                                    //{
                                    //    blockAttributeModel.BackgroundFillColor = attributeReference.MTextAttribute.BackgroundFillColor;
                                    //    blockAttributeModel.BackgroundMaskScaleFactor = attributeReference.MTextAttribute.BackgroundScaleFactor;
                                    //}

                                    // finally, add the model to the collection of models.
                                    output.Add(blockAttributeModel);
                                }
                                else
                                {
                                    // if the attribute is not multiline, it must be single line, so increment the singleLineAttributeQuanity +1...
                                    singleLineAttributeQuantity++;

                                    // and add it's tag to the list of single line references
                                    singleLineAttributeReferencesList.Add(attributeReference.Tag);
                                }
                            }
                            // if attribute count is 0, inform the user that the block they selected does not contain any attributes at all.
                            if (attributeQuantity == 0)
                            {
                                MessageBox.Show("The block you selected does not contain any attributes.");
                            }

                            // if the block contains only single line attributes, inform the user that the block does not contain multiline attributes.
                            if (singleLineAttributeQuantity > 0 && mlineAttributeQuantity == 0)
                            {
                                MessageBox.Show("The block you selected does not contain any multiline block attributes.");
                            }

                            // if the block contains both single line attributes and multiline attributes, inform the user that only multiline attributes will be displayed.
                            if (singleLineAttributeQuantity > 0 && mlineAttributeQuantity > 0)
                            {
                                var singleLineAttributesFoundMessage = string.Join(Environment.NewLine, singleLineAttributeReferencesList);
                                MessageBox.Show("The following attribute references in the selected block are not multiline attributes and were excluded:\r\n\r\n" + singleLineAttributesFoundMessage);
                            }

                            // Return the completed BlockModel to the caller
                            return output;
                        }
                    }
                    catch (Exception ex)
                    {
                        // If an exception is hit display the exception...
                        MessageBox.Show(ex.Message);

                        // and return an empty model
                        return output;
                    }
                }
            }

            // if the prompt status is not okay.
            else
            // Return an empty model
            {
                return output;
            }
        }

        
        
        /// <summary>
        /// Accepts an <see cref="ObjectId"/> of a <see cref="BlockReference"/> and a <see cref="string"/> tag and returns a <see cref="BlockAttributeModel"/> poulated with the data from the <see cref="Attribute"/> found in that block.
        /// </summary>
        /// <param name="selectedBlockID">The <see cref="ObjectId"/> of the <see cref="BlockReference"/> you want to read.</param>
        /// <param name="attributeTag">The tag of the <see cref="Attribute"/> you want to build this <see cref="BlockAttributeModel"/> from.</param>
        /// <returns></returns>
        public BlockAttributeModel ReturnAttributeData(ObjectId selectedBlockID, string attributeTag)
        {
            // set up our attribute model to be returned by the method.
            BlockAttributeModel output = new BlockAttributeModel();

            //access the drawing
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            // Lock the drawing
            using (DocumentLock docLock = thisDrawing.LockDocument())
            {
                try
                {
                    // Start a transaction
                    using (Transaction transaction = thisDrawing.TransactionManager.StartTransaction())
                    {
                        // get the block reference via the passed in ObjectId.
                        BlockReference selectedBlock = transaction.GetObject(selectedBlockID, OpenMode.ForRead) as BlockReference;

                        // Read all attributes in the block and look for the one that matches the passed in tag.
                        AttributeCollection attributeCollection = selectedBlock.AttributeCollection;
                        foreach (ObjectId attributeID in attributeCollection)
                        {
                            AttributeReference attributeReference = (AttributeReference)transaction.GetObject(attributeID, OpenMode.ForRead);
                            // Get the attribute's layer and linetype
                            if (attributeReference.Tag == attributeTag)
                            {
                                // Get the linetype of the attribute reference.
                                LinetypeTableRecord lineTypeRecord = transaction.GetObject(attributeReference.LinetypeId, OpenMode.ForRead) as LinetypeTableRecord;

                                // read the values on the block to return them to output.
                                output.BlockName = selectedBlock.IsDynamicBlock ? ((BlockTableRecord)selectedBlock.DynamicBlockTableRecord.GetObject(OpenMode.ForRead)).Name : selectedBlock.Name;
                                output.Tag = attributeReference.Tag;
                                output.Layer = attributeReference.Layer;
                                #region Color
                                //Get the attribute color by creating an autocad color, setting it to the value of the color found on the layer, and converting it as necessary.
                                Autodesk.AutoCAD.Colors.Color color1 = attributeReference.Color;
                                string color_string = Convert.ToString(color1);
                                //Check for TrueColor by looking for a comma in the value and returning an RGB.
                                if (color_string.ToLower().Contains(",") == true)
                                {
                                    int idx1 = color_string.IndexOf(",", 0);
                                    byte R = Convert.ToByte(color_string.Substring(0, idx1));
                                    int idx2 = color_string.IndexOf(",", idx1 + 1);
                                    byte G = Convert.ToByte(color_string.Substring(idx1 + 1, idx2 - idx1 - 1));
                                    byte B = Convert.ToByte(color_string.Substring(idx2 + 1, color_string.Length - idx2 - 1));
                                    color1 = Autodesk.AutoCAD.Colors.Color.FromRgb(R, G, B);

                                    //Return the truecolor value.
                                    output.Color = color1;
                                }
                                else
                                {
                                    //Return the index color.
                                    output.Color = color1;
                                }
                                #endregion
                                output.Linetype = lineTypeRecord.Name;
                                output.Lineweight = attributeReference.LineWeight;
                                output.LineWeightString = attributeReference.LineWeight.ToString();
                                output.Justification = attributeReference.MTextAttribute.Attachment;
                                output.JustificationString = attributeReference.MTextAttribute.Attachment.ToString();
                                output.Height = attributeReference.Height;
                                output.Rotation = attributeReference.Rotation;
                                output.WidthFactor = attributeReference.WidthFactor;
                                output.TextString = attributeReference.TextString;
                                output.MtextAttributeContent = attributeReference.MTextAttribute.Contents;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    // If an exception is hit display the exception...
                    MessageBox.Show(ex.Message);

                    // and return an empty model
                    return output;
                }
            }
            return output;
        }

        
        
        /// <summary>
        /// Updates the given <see cref="AttributeReference"/> with information passed in by a <see cref="BlockAttributeModel"/>.
        /// </summary>
        /// <param name="blockObjectID"></param>
        /// <param name="blockAttributeModel"></param>
        public void UpdateBlockAttributeReference(ObjectId blockObjectID, BlockAttributeModel blockAttributeModel)
        {
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            using (DocumentLock docLock = thisDrawing.LockDocument())
            {
                try
                {
                    using (Transaction transaction = thisDrawing.TransactionManager.StartTransaction())
                    {
                        if (blockObjectID.IsValid == true)
                        {
                            // get the block reference via the passed in ObjectId.
                            BlockReference selectedBlock = transaction.GetObject(blockObjectID, OpenMode.ForRead) as BlockReference;

                            if (selectedBlock != null)
                            {
                                foreach (ObjectId attributeID in selectedBlock.AttributeCollection)
                                {
                                    AttributeReference attributeReference = attributeID.GetObject(OpenMode.ForWrite) as AttributeReference;

                                    LinetypeTableRecord lineTypeRecord = transaction.GetObject(attributeReference.LinetypeId, OpenMode.ForWrite) as LinetypeTableRecord;

                                    if (attributeReference.Tag == blockAttributeModel.Tag)
                                    {
                                        attributeReference.Tag = blockAttributeModel.Tag;
                                        //attributeReference.Layer = blockAttributeModel.Layer;
                                        //#region Color

                                        //// Create an AutoCAD color by reading the  blockAttributeModel...
                                        //Autodesk.AutoCAD.Colors.Color color1 = blockAttributeModel.Color;
                                        //string color_string = Convert.ToString(color1);

                                        //// Then check to see if we have a true color by looking for a comma in the value...
                                        //if (color_string.ToLower().Contains(",") == true)
                                        //{
                                        //    // if a comma is found, we need to create an R, G, and B byte for each one...
                                        //    int idx1 = color_string.IndexOf(",", 0);
                                        //    byte R = Convert.ToByte(color_string.Substring(0, idx1));
                                        //    int idx2 = color_string.IndexOf(",", idx1 + 1);
                                        //    byte G = Convert.ToByte(color_string.Substring(idx1 + 1, idx2 - idx1 - 1));
                                        //    byte B = Convert.ToByte(color_string.Substring(idx2 + 1, color_string.Length - idx2 - 1));

                                        //    // and set the color to be FromRgb, using those values we just created
                                        //    color1 = Autodesk.AutoCAD.Colors.Color.FromRgb(R, G, B);

                                        //    // finally, return that truecolor value....
                                        //    attributeReference.Color = color1;
                                        //}
                                        //else
                                        //{
                                        //    // Otherwise we have an index color, so we return the index color.
                                        //    attributeReference.Color = color1;
                                        //}

                                        //#endregion
                                        //attributeReference.Linetype = lineTypeRecord.Name;
                                        //attributeReference.LineWeight = blockAttributeModel.Lineweight;
                                        //attributeReference.MTextAttribute.Attachment = blockAttributeModel.Justification;
                                        attributeReference.Height = blockAttributeModel.Height;
                                        attributeReference.Rotation = blockAttributeModel.Rotation;
                                        attributeReference.WidthFactor = blockAttributeModel.WidthFactor;
                                        attributeReference.TextString = blockAttributeModel.TextString;

                                        // First, clear out the old formatting by setting the MTextAttribute.Contents to be equal to the unformatted TextString
                                        attributeReference.MTextAttribute.Contents = blockAttributeModel.TextString;

                                        attributeReference.MTextAttribute.BackgroundFill = blockAttributeModel.BackgroundFill;
                                        attributeReference.MTextAttribute.UseBackgroundColor = blockAttributeModel.UseBackgroundColor;
                                        if (attributeReference.MTextAttribute.BackgroundFill == true)
                                        {
                                            attributeReference.MTextAttribute.BackgroundScaleFactor = blockAttributeModel.BackgroundMaskScaleFactor;
                                        }
                                    }
                                }
                                transaction.Commit();  
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

        }

       
        
        /// <summary>
        /// Updates the given <see cref="AttributeReference"/> text value with information passed in by a <see cref="BlockAttributeModel"/>.
        /// </summary>
        /// <param name="blockObjectID"></param>
        /// <param name="blockAttributeModel"></param>
        public void UpdateBlockAttributeReferenceTextOnly(ObjectId blockObjectID, BlockAttributeModel blockAttributeModel)
        {
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            using (DocumentLock docLock = thisDrawing.LockDocument())
            {
                try
                {
                    using (Transaction transaction = thisDrawing.TransactionManager.StartTransaction())
                    {
                        if (blockObjectID.IsValid == true)
                        {
                            // get the block reference via the passed in ObjectId.
                            BlockReference selectedBlock = transaction.GetObject(blockObjectID, OpenMode.ForRead) as BlockReference;

                            if (selectedBlock != null)
                            {
                                foreach (ObjectId attributeID in selectedBlock.AttributeCollection)
                                {
                                    AttributeReference attributeReference = attributeID.GetObject(OpenMode.ForWrite) as AttributeReference;

                                    LinetypeTableRecord lineTypeRecord = transaction.GetObject(attributeReference.LinetypeId, OpenMode.ForWrite) as LinetypeTableRecord;

                                    if (attributeReference.Tag == blockAttributeModel.Tag)
                                    {
                                        attributeReference.Tag = blockAttributeModel.Tag;
                                        attributeReference.TextString = blockAttributeModel.TextString;

                                        // Clear out the old formatting by setting the MTextAttribute.Contents to be equal to the unformatted TextString
                                        attributeReference.MTextAttribute.Contents = blockAttributeModel.TextString;
                                    }
                                }
                                transaction.Commit();  
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

        }
       
        
        /// <summary>
        /// Updates the given <see cref="AttributeReference"/> text value with information passed in by a <see cref="BlockAttributeModel"/>.
        /// </summary>
        /// <param name="blockObjectID"></param>
        /// <param name="blockAttributeModel"></param>
        public void ApplyBackgroundMask(ObjectId blockObjectID, BlockAttributeModel blockAttributeModel)
        {
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            using (DocumentLock docLock = thisDrawing.LockDocument())
            {
                try
                {
                    using (Transaction transaction = thisDrawing.TransactionManager.StartTransaction())
                    {
                        // get the block reference via the passed in ObjectId.
                        BlockReference selectedBlock = transaction.GetObject(blockObjectID, OpenMode.ForRead) as BlockReference;

                        foreach (ObjectId attributeID in selectedBlock.AttributeCollection)
                        {
                            AttributeReference attributeReference = attributeID.GetObject(OpenMode.ForWrite) as AttributeReference;

                            LinetypeTableRecord lineTypeRecord = transaction.GetObject(attributeReference.LinetypeId, OpenMode.ForWrite) as LinetypeTableRecord;

                            if (attributeReference.Tag == blockAttributeModel.Tag)
                            {
                                if (attributeReference.MTextAttribute.BackgroundFill == false)
                                {
                                    MText mText = attributeReference.MTextAttribute;

                                    mText.BackgroundFill = true;
                                    mText.UseBackgroundColor = true;
                                    mText.BackgroundScaleFactor = 1.50;
                                    attributeReference.MTextAttribute = mText;
                                }
                                else
                                {
                                    MText mText = attributeReference.MTextAttribute;

                                    mText.BackgroundFill = false;
                                    mText.UseBackgroundColor = false;
                                    attributeReference.MTextAttribute = mText;
                                } 
                            }
                        }
                        transaction.Commit();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

        }


        public void UpdateBlockAttributeReferenceWidthFactorViaTextString(ObjectId blockObjectID, BlockAttributeModel blockAttributeModel)
        {
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            using (DocumentLock docLock = thisDrawing.LockDocument())
            {
                try
                {
                    using (Transaction transaction = thisDrawing.TransactionManager.StartTransaction())
                    {
                        if (blockObjectID.IsValid == true)
                        {
                            // get the block reference via the passed in ObjectId.
                            BlockReference selectedBlock = transaction.GetObject(blockObjectID, OpenMode.ForRead) as BlockReference;

                            if (selectedBlock != null)
                            {
                                foreach (ObjectId attributeID in selectedBlock.AttributeCollection)
                                {
                                    AttributeReference attributeReference = attributeID.GetObject(OpenMode.ForWrite) as AttributeReference;

                                    if (attributeReference.Tag == blockAttributeModel.Tag)
                                    {
                                        attributeReference.Tag = blockAttributeModel.Tag;
                                        attributeReference.TextString = blockAttributeModel.TextString;
                                    }
                                }
                                transaction.Commit();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

        }
       
        #endregion

        #region Layers and Linetypes

        /// <summary>
        /// Accesses the <see cref="Database"/> of your current AutoCAD drawing and returns an <see cref="ObservableCollection{string}"/>  of type <see cref="string"/> with all of the layer names.
        /// </summary>
        /// <returns></returns>
        public ObservableCollection<string> GetDrawingLayerNames()
        {
            // set up our output
            ObservableCollection<string> output = new ObservableCollection<string>();

            //Get the drawing and it's database
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database thisDatabase = thisDrawing.Database;

            // Lock the drawing
            using (DocumentLock docLock = thisDrawing.LockDocument())
            {
                try
                {
                    // Start a transaction
                    using (Transaction transaction = thisDrawing.TransactionManager.StartTransaction())
                    {
                        // Open the Layer Table
                        LayerTable layerTable = transaction.GetObject(thisDatabase.LayerTableId, OpenMode.ForRead) as LayerTable;

                        // iterate through the layer IDs...
                        foreach (ObjectId layerID in layerTable)
                        {
                            // and create a layer reference for each layer id found
                            LayerTableRecord layer = transaction.GetObject(layerID, OpenMode.ForRead) as LayerTableRecord;
                            output.Add(layer.Name);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return output;
                }
            }

            return output;
        }
        /// <summary>
        /// Accesses the <see cref="Database"/> of your current AutoCAD drawing and returns an <see cref="ObservableCollection{string}"/> of type <see cref="string"/> of the Linetype names.
        /// </summary>
        /// <returns></returns>
        public ObservableCollection<string> GetDrawingLinetypes()
        {
            // set up our output
            ObservableCollection<string> output = new ObservableCollection<string>();

            //Get the drawing and it's database
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database thisDatabase = thisDrawing.Database;

            // Lock the drawing
            using (DocumentLock docLock = thisDrawing.LockDocument())
            {
                try
                {
                    // Start a transaction
                    using (Transaction transaction = thisDrawing.TransactionManager.StartTransaction())
                    {
                        // Open the Layer Table
                        LinetypeTable lineTypeTable = transaction.GetObject(thisDatabase.LinetypeTableId, OpenMode.ForRead) as LinetypeTable;

                        // iterate through the layer IDs...
                        foreach (ObjectId lineTypeID in lineTypeTable)
                        {
                            // and create a layer reference for each layer id found
                            LinetypeTableRecord linetype = transaction.GetObject(lineTypeID, OpenMode.ForRead) as LinetypeTableRecord;
                            output.Add(linetype.Name);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return output;
                }
            }

            return output;
        }

        #endregion
    }
}
