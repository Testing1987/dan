using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Windows;

namespace BoreExcavationBuilder.CoreLogic
{
    public class DataAccess
    {
        public void DrawBoreHole(double length, double width, bool useHatch, string selectedHatch, string BoreHoleLayer, string HatchLayer, double hatchScale, double hatchAngle)
        {
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database thisDatabase = thisDrawing.Database;
            Editor thisEditor = thisDrawing.Editor;
            using (DocumentLock documentLock = thisDrawing.LockDocument())
            {
                using (Transaction thisTransaction = thisDrawing.TransactionManager.StartTransaction())
                {
                    //ObjectId modelSpaceID = SymbolUtilityServices.GetBlockModelSpaceId(thisDatabase);
                    BlockTable blockTable = thisTransaction.GetObject(thisDatabase.BlockTableId, OpenMode.ForWrite) as BlockTable;
                    BlockTableRecord modelSpace = thisTransaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    //Set up the polyline for the bore pit
                    Polyline polyline = new Polyline();
                    polyline.Normal = Vector3d.ZAxis;

                    //Set up the face of the bore pit and a point used to calculate rotation
                    Point3d face_point = new Point3d(0, 0, 0);
                    Point3d rotation_point = new Point3d(0, 0, 0);

                    //Put the user's cursor in modelspace
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                    //Ask the user to get the face of the bore pit
                    PromptPointResult ppr_face;
                    PromptPointOptions ppoptions_face = new PromptPointOptions("");
                    ppoptions_face.Message = "\nSelect the front of the bore pit.";
                    ppoptions_face.AllowNone = false;
                    ppr_face = thisEditor.GetPoint(ppoptions_face);
                    if (ppr_face.Status != PromptStatus.OK)
                    {
                        return;
                    }
                    //set the value of point3D_Face
                    face_point = ppr_face.Value;

                    //Ask the user to select an alignment point for bore pit rotation
                    object currentSnapMode = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("OSMODE");
                    object setSnapMode = 512;
                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", setSnapMode);
                    PromptPointResult ppr_rotation;
                    PromptPointOptions ppoptions_rotation = new PromptPointOptions("");
                    ppoptions_rotation.Message = "\nSelect any point along the bore pit alignment.";
                    ppoptions_rotation.AllowNone = false;
                    ppr_rotation = thisEditor.GetPoint(ppoptions_rotation);
                    if (ppr_rotation.Status != PromptStatus.OK)
                    {
                        return;
                    }
                    //set the value of point3D_Rotation
                    rotation_point = ppr_rotation.Value;

                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", currentSnapMode);

                    //define the polyline
                    //MessageBox.Show(face_point.ToString() + "\r\n" + rotation_point.ToString());
                    polyline.AddVertexAt(0, new Point2d(face_point.X, face_point.Y), 0, 0, 0);
                    polyline.AddVertexAt(1, new Point2d(face_point.X + length, face_point.Y), 0, 0, 0);
                    polyline.AddVertexAt(2, new Point2d(face_point.X + length, face_point.Y + width), 0, 0, 0);
                    polyline.AddVertexAt(3, new Point2d(face_point.X, face_point.Y + width), 0, 0, 0);
                    polyline.Closed = true;

                    //Move the polyline to be centered on the selected alignment
                    Vector3d vector3D = polyline.GetPointAtParameter(0) - polyline.GetPointAtParameter(3);
                    Point3d midPoint = polyline.GetPointAtParameter(0) + vector3D * 0.5;
                    Vector3d moveVector = polyline.GetPointAtParameter(0).GetVectorTo(midPoint);
                    polyline.TransformBy(Matrix3d.Displacement(moveVector));

                    //Make sure the layers we need for the polyline and the hatch are set up.
                    #region Layer Setup
                    try
                    {
                        if (string.IsNullOrEmpty(BoreHoleLayer))
                        {
                            BoreHoleLayer = "EXCAVATION_PIT";
                        }
                        LayerTable layerTable = thisTransaction.GetObject(thisDatabase.LayerTableId, OpenMode.ForRead) as LayerTable;
                        if (layerTable.Has(BoreHoleLayer) == false)
                        {
                            LayerTableRecord layerTableRecord = new LayerTableRecord();
                            layerTableRecord.Color = Color.FromColorIndex(ColorMethod.ByAci, 7);
                            layerTableRecord.Name = BoreHoleLayer;
                            layerTable.UpgradeOpen();
                            layerTable.Add(layerTableRecord);
                            thisTransaction.AddNewlyCreatedDBObject(layerTableRecord, true);
                        }

                        if (useHatch)
                        {
                            //Put the hatch on the correct layer
                            if (string.IsNullOrEmpty(HatchLayer))
                            {
                                HatchLayer = "EXCAVATION_PIT_HATCH";
                            }
                            if (layerTable.Has(HatchLayer) == false)
                            {
                                LayerTableRecord layerTableRecord = new LayerTableRecord();
                                layerTableRecord.Color = Color.FromColorIndex(ColorMethod.ByAci, 7);
                                layerTableRecord.Name = HatchLayer;
                                layerTable.UpgradeOpen();
                                layerTable.Add(layerTableRecord);
                                thisTransaction.AddNewlyCreatedDBObject(layerTableRecord, true);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    #endregion
                    //Put the polyline on the layer.
                    polyline.Layer = BoreHoleLayer;

                    //Set up the rotation of the polyline
                    double rotation = 0;
                    Plane plane = new Plane(new Point3d(0, 0, 0), Vector3d.ZAxis);

                    //Get the vector between the points the user selected as the face point and the rotation point
                    rotation = new Point3d(face_point.X, face_point.Y, 0).GetVectorTo(new Point3d(rotation_point.X, rotation_point.Y, 0)).AngleOnPlane(plane);

                    //rotate the completed polyline based on the user's selection
                    polyline.TransformBy(Matrix3d.Rotation(rotation, Vector3d.ZAxis, face_point));

                    //Append the polyline to the drawing
                    ObjectId polylineId;
                    polylineId = modelSpace.AppendEntity(polyline);
                    thisTransaction.AddNewlyCreatedDBObject(polyline, true);

                    if (useHatch)
                    {
                        try
                        {
                            //Set up an ObjectIdCollection and add the polyline ObjectId to it
                            ObjectIdCollection objectIdCollection = new ObjectIdCollection();
                            objectIdCollection.Add(polylineId);

                            //Create a hatch and set up it's parameters
                            Hatch hatch = new Hatch();
                            Vector3d normal = new Vector3d(0.0, 0.0, 1.0);
                            hatch.Normal = normal;
                            hatch.Elevation = 0.0;
                            hatch.PatternScale = hatchScale;
                            hatch.Layer = HatchLayer;
                            try
                            {
                                hatch.SetHatchPattern(HatchPatternType.PreDefined, selectedHatch);
                                hatch.PatternAngle = hatchAngle / 180 * Math.PI;
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                            //append the hatch to the drawing
                            modelSpace.AppendEntity(hatch);
                            thisTransaction.AddNewlyCreatedDBObject(hatch, true);

                            //Hatch the objects inside the objectIdCollection (which is just the polyline that we created)
                            hatch.Associative = true;
                            hatch.AppendLoop((int)HatchLoopTypes.Default, objectIdCollection);
                            hatch.EvaluateHatch(true);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }

                    //Commit the transaction
                    thisTransaction.Commit();
                }
            }
        }
    }
}
