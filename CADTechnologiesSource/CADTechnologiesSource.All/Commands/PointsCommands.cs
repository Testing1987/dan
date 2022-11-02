using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System.Windows;

namespace CADTechnologiesSource.All.Commands
{
    public class PointsCommands
    {
        [CommandMethod("POLY_START_END_POINT_CREATOR")]
        public void CreatePointsOnStartAndEndOfPolylines()
        {
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database thisDatabase = thisDrawing.Database;
            Editor thisEditor = thisDrawing.Editor;
            using (Transaction thisTransaction = thisDrawing.TransactionManager.StartTransaction())
            {
                try
                {
                    //set autocad focus
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                    // Create a new PromptSelectionOptions...
                    PromptSelectionOptions pso = new PromptSelectionOptions();

                    // Ask the user to select items
                    PromptSelectionResult psr = thisDrawing.Editor.GetSelection(pso);
                    if (psr.Status == PromptStatus.OK)
                    {
                        BlockTable blockTable = thisTransaction.GetObject(thisDrawing.Database.BlockTableId, OpenMode.ForWrite) as BlockTable;
                        BlockTableRecord modelSpace = thisTransaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                        // Create a new selection set and assign the user's selection to it
                        SelectionSet selectionSet = psr.Value;

                        // iterate through each object in the selection and...
                        foreach (SelectedObject selectedObject in selectionSet)
                        {
                            if (selectedObject != null)
                            {
                                // create an entity for each one
                                Entity entity = thisTransaction.GetObject(selectedObject.ObjectId, OpenMode.ForWrite) as Entity;
                                if (entity != null)
                                {

                                    // Check to make sure the entity is a polyline
                                    if (entity is Polyline)
                                    {
                                        // Read the polylines from the the selected objects
                                        Polyline selectedPolyline = (Polyline)thisTransaction.GetObject(entity.ObjectId, OpenMode.ForRead);

                                        Point3d startPoint =  selectedPolyline.GetPointAtParameter(selectedPolyline.StartParam);
                                        Point3d endPoint = selectedPolyline.GetPointAtParameter(selectedPolyline.EndParam);


                                        DBPoint point1 = new DBPoint(startPoint);
                                        DBPoint point2 = new DBPoint(endPoint);

                                        modelSpace.AppendEntity(point1);
                                        modelSpace.AppendEntity(point2);
                                        thisTransaction.AddNewlyCreatedDBObject(point1, true);
                                        thisTransaction.AddNewlyCreatedDBObject(point2, true);
                                        thisDatabase.Pdmode = 35;
                                        thisDatabase.Pdsize = 15;
                                    }
                                    else
                                    {
                                        MessageBox.Show("No polyline was selected.");
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
                thisTransaction.Commit();
            }
        }
    }
}
