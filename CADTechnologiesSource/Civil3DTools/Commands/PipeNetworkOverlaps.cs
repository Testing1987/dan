using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Civil3DTools.Commands
{
    public class PipeNetworkOverlaps
    {
        [CommandMethod("DELETE_OVERLAPPING_PIPE_NETWORKS")]
        public void FindPipeNetworkOverlaps()
        {
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor thisEditor = thisDrawing.Editor;
            Database thisDatabase = thisDrawing.Database;

            List<Polyline3d> polyline3Ds = new List<Polyline3d>();
            List<Network> pipeNetworks = new List<Network>();
            int deletednetwork_count = 0;


            try
            {
                using(Transaction transaction = thisDrawing.TransactionManager.StartTransaction())
                {
                    BlockTable blockTable = transaction.GetObject(thisDatabase.BlockTableId, OpenMode.ForWrite) as BlockTable;
                    BlockTableRecord modelSpace = transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    foreach (ObjectId objectId in modelSpace)
                    {
                        //Find any polyline 3Ds and add them to the list of 3D Polylines
                        Autodesk.AutoCAD.DatabaseServices.DBObject dBObject = transaction.GetObject(objectId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.DBObject;
                        if (dBObject is Polyline3d)
                        {
                            Polyline3d polyline3D = transaction.GetObject(objectId, OpenMode.ForRead) as Polyline3d;
                            polyline3Ds.Add(polyline3D);
                        }

                        if (dBObject is Network)
                        {
                            //Find any pipenetworks and add them to the list of pipenetworks
                            Network network = transaction.GetObject(objectId, OpenMode.ForWrite) as Network;
                            pipeNetworks.Add(network);
                        }
                    }

                    //sort through the networks to delete the ones that overlap with 3D Polylines
                    var deleteme = from network in pipeNetworks
                                               from poly_3d in polyline3Ds
                                               where network.StartPoint == poly_3d.StartPoint && network.EndPoint == poly_3d.EndPoint ||
                                               network.StartPoint == poly_3d.EndPoint && network.EndPoint == poly_3d.StartPoint
                                               select network;

                    if(deleteme.Count()>0)
                    {
                        //Create a list of the results of the deleteme query
                        List<Network> networksToDelete = new List<Network>(deleteme);

                        //and try to delete them
                        try
                        {
                            if (networksToDelete != null && networksToDelete.Count() >= 0)
                            {
                                foreach (var network in networksToDelete)
                                {
                                    network.Erase(true);
                                    deletednetwork_count++;
                                }
                            }
                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                        transaction.Commit();
                        thisEditor.WriteMessage($"\r\n{deletednetwork_count} Pipe Networks that overlapped with 3DPolylines have been deleted.");
                    }
                    else
                    {
                        transaction.Abort();
                        thisEditor.WriteMessage($"\r\nNo overlaps have been found.");
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
