using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace FastAtt.CoreLogic
{
    public class CADAccess
    {
        public void CreateBlockAttribute(bool Invisible, bool Constant, bool Verify, bool Preset, bool LockPosition, bool MultipleLines, bool SpecifyInsertionInApp, double InsertX, double InsertY, double InsertZ, string Tag, string SelectedTextStyle, string SelectedJustification, bool Annotative, double TextHeight, double Rotation)
        {
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database thisDatabase = thisDrawing.Database;
            Editor thisEditor = thisDrawing.Editor;
            int tagIndex = 1;
        Repeat:
            bool repeat = false;

            using (DocumentLock documentLock = thisDrawing.LockDocument())
            {
                using (Transaction thisTransaction = thisDatabase.TransactionManager.StartTransaction())
                {
                    BlockTable blockTable = thisTransaction.GetObject(thisDatabase.BlockTableId, OpenMode.ForWrite) as BlockTable;
                    BlockTableRecord currentSpace = thisTransaction.GetObject(thisDatabase.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                    AttributeDefinition attributeDefinition = new AttributeDefinition();

                    //read the index of the tag and insert it, then iterate the value.
                    attributeDefinition.Tag = Tag + "_" + tagIndex.ToString();
                    attributeDefinition.Prompt = Tag + "_" + tagIndex.ToString();
                    attributeDefinition.TextString = "-";

                    //populate the attribute definition with the users desired values
                    attributeDefinition.Invisible = Invisible ? true : false;
                    attributeDefinition.Constant = Constant ? true : false;
                    attributeDefinition.Verifiable = Verify ? true : false;
                    attributeDefinition.Preset = Preset ? true : false;
                    attributeDefinition.Invisible = Invisible ? true : false;
                    attributeDefinition.LockPositionInBlock = LockPosition ? true : false;
                    attributeDefinition.IsMTextAttributeDefinition = MultipleLines ? true : false;
                    attributeDefinition.Annotative = Annotative ? AnnotativeStates.True : AnnotativeStates.False;

                    SymbolTable textStyleTable = thisTransaction.GetObject(thisDatabase.TextStyleTableId, OpenMode.ForRead) as SymbolTable;
                    if (textStyleTable != null)
                    {
                        foreach (ObjectId textstyleID in textStyleTable)
                        {
                            TextStyleTableRecord textStyle = thisTransaction.GetObject(textstyleID, OpenMode.ForRead) as TextStyleTableRecord;
                            if (textStyle.Name == SelectedTextStyle)
                            {
                                attributeDefinition.TextStyleId = textStyle.Id;
                            }
                        }

                        foreach (string justification in Enum.GetNames(typeof(AttachmentPoint)))
                        {
                            if (SelectedJustification == justification)
                            {
                                attributeDefinition.Justify = (AttachmentPoint)Enum.Parse(typeof(AttachmentPoint), justification);
                            }
                        }
                        attributeDefinition.Height = TextHeight;
                        attributeDefinition.Rotation = Rotation * Math.PI / 180;

                        try
                        {
                            switch (SpecifyInsertionInApp)
                            {
                                case true:
                                    {
                                        Point3d insertPoint = new Point3d(InsertX, InsertY, InsertZ);
                                        attributeDefinition.AlignmentPoint = insertPoint;

                                        //place it in the drawing
                                        currentSpace.AppendEntity(attributeDefinition);
                                        thisTransaction.AddNewlyCreatedDBObject(attributeDefinition, true);
                                        thisTransaction.Commit();

                                        //repeat makes no sense in this context so don't repeat.
                                        repeat = false;
                                        break;
                                    }
                                case false:
                                    {
                                        //Get the user to select a point
                                        PromptPointOptions promptPoint = new PromptPointOptions("");
                                        promptPoint.Message = "Select insertion point\r\n";
                                        promptPoint.AllowNone = false;
                                        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                                        //Get the result of the users selection
                                        PromptPointResult pointResult = thisEditor.GetPoint(promptPoint);
                                        Point3d insertionPoint = pointResult.Value;
                                        if (pointResult.Status != PromptStatus.Cancel)
                                        {
                                            //incretement the tagindex
                                            tagIndex++;
                                            attributeDefinition.AlignmentPoint = insertionPoint;

                                            //place it in the drawing
                                            currentSpace.AppendEntity(attributeDefinition);
                                            thisTransaction.AddNewlyCreatedDBObject(attributeDefinition, true);
                                            thisTransaction.Commit();

                                            //Since the user is clicking, lets let them repeat
                                            repeat = true;
                                        }
                                        else
                                        {
                                            //if the user has canceled, the operation will stop.
                                            repeat = false;
                                            return;
                                        }
                                        break;
                                    }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }
                    if (repeat == true)
                    {
                        goto Repeat;
                    }
                }
            }
        }
        public List<string> GetTextStylesFromCurrentDrawing()
        {
            List<string> output = new List<string>();
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database thisDatabase = thisDrawing.Database;

            try
            {
                using (Transaction thisTransaction = thisDatabase.TransactionManager.StartTransaction())
                {
                    SymbolTable textStyleTable = thisTransaction.GetObject(thisDatabase.TextStyleTableId, OpenMode.ForRead) as SymbolTable;
                    if (textStyleTable != null)
                    {
                        foreach (ObjectId textstyleID in textStyleTable)
                        {
                            TextStyleTableRecord textStyle = thisTransaction.GetObject(textstyleID, OpenMode.ForRead) as TextStyleTableRecord;
                            output.Add(textStyle.Name);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return output;
        }
    }
}
