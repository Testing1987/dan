using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System.Collections.Generic;
using System.Linq;

namespace CADTechnologiesSource.All.Commands
{
    public class OverlapCheck
    {
        [CommandMethod("DELETE_2D_OVERLAPPING_POLYLINES")]
        public void FindOverlaps()
        {
            List<Polyline3d> polyline3Ds = new List<Polyline3d>();
            List<Polyline> polylines = new List<Polyline>();
            int deletedpoly_count = 0;

            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor thisEditor = thisDrawing.Editor;
            Database thisDatabase = thisDrawing.Database;

            using (Transaction transaction = thisDrawing.TransactionManager.StartTransaction())
            {
                BlockTable blockTable = transaction.GetObject(thisDatabase.BlockTableId, OpenMode.ForWrite) as BlockTable;
                BlockTableRecord modelSpace = transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                //now look for 2d/3d overlaps
                foreach (ObjectId objectId in modelSpace)
                {
                    //Find any polyline 3Ds and add them to the list of 3D Polylines
                    DBObject dBObject = transaction.GetObject(objectId, OpenMode.ForRead) as DBObject;
                    if (dBObject is Polyline3d)
                    {
                        Polyline3d polyline3D = transaction.GetObject(objectId, OpenMode.ForRead) as Polyline3d;
                        polyline3Ds.Add(polyline3D);
                    }

                    if (dBObject is Polyline)
                    {
                        //Find any polyline 2Ds and add them to the list of 2D Polylines
                        Polyline polyline = transaction.GetObject(objectId, OpenMode.ForWrite) as Polyline;
                        polylines.Add(polyline);
                    }
                }

                //set up a list of polylines that need to be deleted
                List<Polyline> polyLinesToDelete = new List<Polyline>();

                //compare each polyline to each polyline3d
                //foreach (var poly in polylines)
                //{
                //    foreach (var poly_3d in polyline3Ds)
                //    {
                //        //if the start and end points match, consider it a duplicate
                //        if (poly.StartParam == poly_3d.StartParam && poly.EndParam == poly_3d.EndParam ||
                //            poly.StartParam == poly_3d.EndParam && poly.EndParam == poly_3d.StartParam)
                //            //add the duplicate polylines to the list of polylines to be deleted
                //            polyLinesToDelete.Add(poly);
                //    }
                //}


                //test
                var deleteme = from poly in polylines
                                           from poly_3d in polyline3Ds
                                           where poly.StartPoint == poly_3d.StartPoint && poly.EndPoint == poly_3d.EndPoint ||
                                                      poly.StartPoint == poly_3d.EndPoint && poly.EndPoint == poly_3d.StartPoint
                                                      select poly;

                List<Polyline> newPolystodelete = new List<Polyline>(deleteme);


                try
                {
                    if (newPolystodelete != null && newPolystodelete.Count() >= 0)
                    {
                        //and erase each polyline
                        foreach (var poly in newPolystodelete)
                        {
                            poly.Erase(true);
                            deletedpoly_count++;
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    throw;
                }

                //test


                //make sure the list isn't empty or null
                //try
                //{
                //    if (polyLinesToDelete != null && polyLinesToDelete.Count() >= 0)
                //    {
                //        //and erase each polyline
                //        foreach (var poly in polyLinesToDelete)
                //        {
                //            poly.Erase(true);
                //            deletedpoly_count++;
                //        }
                //    }
                //}
                //catch (System.Exception ex)
                //{
                //    throw;
                //}
                //when done, commit the transaction
                transaction.Commit();
                thisEditor.WriteMessage($"\r\n{deletedpoly_count} 2DPolylines that overlapped with 3DPolylines have been deleted.");
            }
        }
    }
}

