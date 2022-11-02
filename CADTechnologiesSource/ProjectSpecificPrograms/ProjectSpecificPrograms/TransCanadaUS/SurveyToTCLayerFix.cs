using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjectSpecificPrograms.TransCanadaUS
{
    public class SurveyToTCLayerFix
    {
        [CommandMethod("Survey_Layers_To_TC_Standard")]
        public void ConvertSurveyLayerstoTCUSLayers()
        {
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database thisDatabase = thisDrawing.Database;
            Editor thisEditor = thisDrawing.Editor;
            using (DocumentLock docLock = thisDrawing.LockDocument())
            {
                using (Transaction thisTransaction = thisDatabase.TransactionManager.StartTransaction())
                {
                    LayerTable layerTable = thisTransaction.GetObject(thisDatabase.LayerTableId, OpenMode.ForWrite) as LayerTable;
                    BlockTable blockTable = thisTransaction.GetObject(thisDatabase.BlockTableId, OpenMode.ForWrite) as BlockTable;
                    BlockTableRecord blockTableRecord = thisTransaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    foreach (ObjectId layerID in layerTable)
                    {
                        LayerTableRecord layer = thisTransaction.GetObject(layerID, OpenMode.ForWrite) as LayerTableRecord;

                        //Look at all the survey layers, see if they existing in the drawing, if they do, rename them to TC standard layers.
                        switch (layer.Name)
                        {
                            case "access_roads":
                                //see if the TC standard layer already exists.
                                if (layerTable.Has(TC_US_ExistingConditions_Layers.AccessRoad) == false)
                                {
                                    //if not, rename this layer to it, and now it exists.
                                    layer.Name = TC_US_ExistingConditions_Layers.AccessRoad;

                                }
                                else
                                {
                                    //if the layer does exist in the drawing.....find all objects on the survey layer, and put them on the TC Standard layer.
                                    foreach (ObjectId objectId in blockTableRecord)
                                    {
                                        Entity entity = thisTransaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
                                        if (entity.Layer != null)
                                        {
                                            if (entity.Layer == layer.Name)
                                            {
                                                {
                                                    entity.Layer = TC_US_ExistingConditions_Layers.AccessRoad;

                                                }
                                            }
                                        }
                                    }
                                }
                                break;

                            case "E_UTL_ClientPL":
                                //see if the TC standard layer already exists.
                                if (layerTable.Has(TC_US_ExistingConditions_Layers.ExistingTCPipeline) == false)
                                {
                                    //if not, rename this layer to it, and now it exists.
                                    layer.Name = TC_US_ExistingConditions_Layers.ExistingTCPipeline;

                                }
                                else
                                {
                                    //if the layer does exist in the drawing.....find all objects on the survey layer, and put them on the TC Standard layer.
                                    foreach (ObjectId objectId in blockTableRecord)
                                    {
                                        Entity entity = thisTransaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
                                        if (entity.Layer != null)
                                        {
                                            if (entity.Layer == layer.Name)
                                            {
                                                {
                                                    entity.Layer = TC_US_ExistingConditions_Layers.ExistingTCPipeline;

                                                }
                                            }
                                        }
                                    }
                                }
                                break;

                            case "E_UTL_ForeignPL":
                                //see if the TC standard layer already exists.
                                if (layerTable.Has(TC_US_ExistingConditions_Layers.ForeignPipeline) == false)
                                {
                                    //if not, rename this layer to it, and now it exists.
                                    layer.Name = TC_US_ExistingConditions_Layers.ForeignPipeline;

                                }
                                else
                                {
                                    //if the layer does exist in the drawing.....find all objects on the survey layer, and put them on the TC Standard layer.
                                    foreach (ObjectId objectId in blockTableRecord)
                                    {
                                        Entity entity = thisTransaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
                                        if (entity.Layer != null)
                                        {
                                            if (entity.Layer == layer.Name)
                                            {
                                                {
                                                    entity.Layer = TC_US_ExistingConditions_Layers.ForeignPipeline;

                                                }
                                            }
                                        }
                                    }
                                }
                                break;

                            case "E_TER_TreeLine":
                                //see if the TC standard layer already exists.
                                if (layerTable.Has(TC_US_ExistingConditions_Layers.TreeLine) == false)
                                {
                                    //if not, rename this layer to it, and now it exists.
                                    layer.Name = TC_US_ExistingConditions_Layers.TreeLine;

                                }
                                else
                                {
                                    //if the layer does exist in the drawing.....find all objects on the survey layer, and put them on the TC Standard layer.
                                    foreach (ObjectId objectId in blockTableRecord)
                                    {
                                        Entity entity = thisTransaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
                                        if (entity.Layer != null)
                                        {
                                            if (entity.Layer == layer.Name)
                                            {
                                                {
                                                    entity.Layer = TC_US_ExistingConditions_Layers.TreeLine;

                                                }
                                            }
                                        }
                                    }
                                }
                                break;

                            case "E_TER_BrushLine":
                                //see if the TC standard layer already exists.
                                if (layerTable.Has(TC_US_ExistingConditions_Layers.TreeLine) == false)
                                {
                                    //if not, rename this layer to it, and now it exists.
                                    layer.Name = TC_US_ExistingConditions_Layers.TreeLine;

                                }
                                else
                                {
                                    //if the layer does exist in the drawing.....find all objects on the survey layer, and put them on the TC Standard layer.
                                    foreach (ObjectId objectId in blockTableRecord)
                                    {
                                        Entity entity = thisTransaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
                                        if (entity.Layer != null)
                                        {
                                            if (entity.Layer == layer.Name)
                                            {
                                                {
                                                    entity.Layer = TC_US_ExistingConditions_Layers.TreeLine;
                                                }
                                            }
                                        }
                                    }

                                }
                                break;

                            case "E_TRA_CL_Road":
                                //see if the TC standard layer already exists.
                                if (layerTable.Has(TC_US_ExistingConditions_Layers.CenterofRoad) == false)
                                {
                                    //if not, rename this layer to it, and now it exists.
                                    layer.Name = TC_US_ExistingConditions_Layers.CenterofRoad;

                                }
                                else
                                {
                                    //if the layer does exist in the drawing.....find all objects on the survey layer, and put them on the TC Standard layer.
                                    foreach (ObjectId objectId in blockTableRecord)
                                    {
                                        Entity entity = thisTransaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
                                        if (entity.Layer != null)
                                        {
                                            if (entity.Layer == layer.Name)
                                            {
                                                {
                                                    entity.Layer = TC_US_ExistingConditions_Layers.CenterofRoad;

                                                }
                                            }
                                        }
                                    }
                                }
                                break;

                            case "E_FEA_Culvert":
                                //see if the TC standard layer already exists.
                                if (layerTable.Has(TC_US_ExistingConditions_Layers.Culvert) == false)
                                {
                                    //if not, rename this layer to it, and now it exists.
                                    layer.Name = TC_US_ExistingConditions_Layers.Culvert;

                                }
                                else
                                {
                                    //if the layer does exist in the drawing.....find all objects on the survey layer, and put them on the TC Standard layer.
                                    foreach (ObjectId objectId in blockTableRecord)
                                    {
                                        Entity entity = thisTransaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
                                        if (entity.Layer != null)
                                        {
                                            if (entity.Layer == layer.Name)
                                            {
                                                {
                                                    entity.Layer = TC_US_ExistingConditions_Layers.Culvert;

                                                }
                                            }
                                        }
                                    }
                                }
                                break;

                            case "E_PNT_Culvert":
                                //see if the TC standard layer already exists.
                                if (layerTable.Has(TC_US_ExistingConditions_Layers.Culvert) == false)
                                {
                                    //if not, rename this layer to it, and now it exists.
                                    layer.Name = TC_US_ExistingConditions_Layers.Culvert;

                                }
                                else
                                {
                                    //if the layer does exist in the drawing.....find all objects on the survey layer, and put them on the TC Standard layer.
                                    foreach (ObjectId objectId in blockTableRecord)
                                    {
                                        Entity entity = thisTransaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
                                        if (entity.Layer != null)
                                        {
                                            if (entity.Layer == layer.Name)
                                            {
                                                {
                                                    entity.Layer = TC_US_ExistingConditions_Layers.Culvert;

                                                }
                                            }
                                        }
                                    }
                                }
                                break;

                            case "Culvert":
                                //see if the TC standard layer already exists.
                                if (layerTable.Has(TC_US_ExistingConditions_Layers.Culvert) == false)
                                {
                                    //if not, rename this layer to it, and now it exists.
                                    layer.Name = TC_US_ExistingConditions_Layers.Culvert;

                                }
                                else
                                {
                                    //if the layer does exist in the drawing.....find all objects on the survey layer, and put them on the TC Standard layer.
                                    foreach (ObjectId objectId in blockTableRecord)
                                    {
                                        Entity entity = thisTransaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
                                        if (entity.Layer != null)
                                        {
                                            if (entity.Layer == layer.Name)
                                            {
                                                {
                                                    entity.Layer = TC_US_ExistingConditions_Layers.Culvert;

                                                }
                                            }
                                        }
                                    }
                                }
                                break;

                            case "E_TER_WaterCL":
                                //see if the TC standard layer already exists.
                                if (layerTable.Has(TC_US_Environmental_Layers.CenterlineOfWaterway) == false)
                                {
                                    //if not, rename this layer to it, and now it exists.
                                    layer.Name = TC_US_Environmental_Layers.CenterlineOfWaterway;

                                }
                                else
                                {
                                    //if the layer does exist in the drawing.....find all objects on the survey layer, and put them on the TC Standard layer.
                                    foreach (ObjectId objectId in blockTableRecord)
                                    {
                                        Entity entity = thisTransaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
                                        if (entity.Layer != null)
                                        {
                                            if (entity.Layer == layer.Name)
                                            {
                                                {
                                                    entity.Layer = TC_US_Environmental_Layers.CenterlineOfWaterway;

                                                }
                                            }
                                        }
                                    }
                                }
                                break;

                            case "E_TER_TopBank":
                                //see if the TC standard layer already exists.
                                if (layerTable.Has(TC_US_Environmental_Layers.TopofBank) == false)
                                {
                                    //if not, rename this layer to it, and now it exists.
                                    layer.Name = TC_US_Environmental_Layers.TopofBank;

                                }
                                else
                                {
                                    //if the layer does exist in the drawing.....find all objects on the survey layer, and put them on the TC Standard layer.
                                    foreach (ObjectId objectId in blockTableRecord)
                                    {
                                        Entity entity = thisTransaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
                                        if (entity.Layer != null)
                                        {
                                            if (entity.Layer == layer.Name)
                                            {
                                                {
                                                    entity.Layer = TC_US_Environmental_Layers.TopofBank;

                                                }
                                            }
                                        }
                                    }
                                }
                                break;

                            case "E_TER_ToeBank":
                                //see if the TC standard layer already exists.
                                if (layerTable.Has(TC_US_Environmental_Layers.ToeofBank) == false)
                                {
                                    //if not, rename this layer to it, and now it exists.
                                    layer.Name = TC_US_Environmental_Layers.ToeofBank;

                                }
                                else
                                {
                                    //if the layer does exist in the drawing.....find all objects on the survey layer, and put them on the TC Standard layer.
                                    foreach (ObjectId objectId in blockTableRecord)
                                    {
                                        Entity entity = thisTransaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
                                        if (entity.Layer != null)
                                        {
                                            if (entity.Layer == layer.Name)
                                            {
                                                {
                                                    entity.Layer = TC_US_Environmental_Layers.ToeofBank;

                                                }
                                            }
                                        }
                                    }
                                }
                                break;

                            case "E_UTL_PowerOH":
                                //see if the TC standard layer already exists.
                                if (layerTable.Has(TC_US_ExistingConditions_Layers.POWERLINE) == false)
                                {
                                    //if not, rename this layer to it, and now it exists.
                                    layer.Name = TC_US_ExistingConditions_Layers.POWERLINE;

                                }
                                else
                                {
                                    //if the layer does exist in the drawing.....find all objects on the survey layer, and put them on the TC Standard layer.
                                    foreach (ObjectId objectId in blockTableRecord)
                                    {
                                        Entity entity = thisTransaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
                                        if (entity.Layer != null)
                                        {
                                            if (entity.Layer == layer.Name)
                                            {
                                                {
                                                    entity.Layer = TC_US_ExistingConditions_Layers.POWERLINE;

                                                }
                                            }
                                        }
                                    }
                                }
                                break;

                            case "E_TRA_RD_EdgeAsph":
                                //see if the TC standard layer already exists.
                                if (layerTable.Has(TC_US_ExistingConditions_Layers.EdgeofRoad) == false)
                                {
                                    //if not, rename this layer to it, and now it exists.
                                    layer.Name = TC_US_ExistingConditions_Layers.EdgeofRoad;

                                }
                                else
                                {
                                    //if the layer does exist in the drawing.....find all objects on the survey layer, and put them on the TC Standard layer.
                                    foreach (ObjectId objectId in blockTableRecord)
                                    {
                                        Entity entity = thisTransaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
                                        if (entity.Layer != null)
                                        {
                                            if (entity.Layer == layer.Name)
                                            {
                                                {
                                                    entity.Layer = TC_US_ExistingConditions_Layers.EdgeofRoad;

                                                }
                                            }
                                        }
                                    }
                                }
                                break;

                            case "E_PNT_Centerline_Waterway":
                                //see if the TC standard layer already exists.
                                if (layerTable.Has(TC_US_Environmental_Layers.CenterlineOfWaterway) == false)
                                {
                                    //if not, rename this layer to it, and now it exists.
                                    layer.Name = TC_US_Environmental_Layers.CenterlineOfWaterway;

                                }
                                else
                                {
                                    //if the layer does exist in the drawing.....find all objects on the survey layer, and put them on the TC Standard layer.
                                    foreach (ObjectId objectId in blockTableRecord)
                                    {
                                        Entity entity = thisTransaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
                                        if (entity.Layer != null)
                                        {
                                            if (entity.Layer == layer.Name)
                                            {
                                                {
                                                    entity.Layer = TC_US_Environmental_Layers.CenterlineOfWaterway;

                                                }
                                            }
                                        }
                                    }
                                }
                                break;

                            case "E_TER_WaterEdge":
                                //see if the TC standard layer already exists.
                                if (layerTable.Has(TC_US_Environmental_Layers.WATERBODY) == false)
                                {
                                    //if not, rename this layer to it, and now it exists.
                                    layer.Name = TC_US_Environmental_Layers.WATERBODY;

                                }
                                else
                                {
                                    //if the layer does exist in the drawing.....find all objects on the survey layer, and put them on the TC Standard layer.
                                    foreach (ObjectId objectId in blockTableRecord)
                                    {
                                        Entity entity = thisTransaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
                                        if (entity.Layer != null)
                                        {
                                            if (entity.Layer == layer.Name)
                                            {
                                                {
                                                    entity.Layer = TC_US_Environmental_Layers.WATERBODY;

                                                }
                                            }
                                        }
                                    }
                                }
                                break;

                            case "Wetland_Polygon":
                                //see if the TC standard layer already exists.
                                if (layerTable.Has(TC_US_Environmental_Layers.Wetland) == false)
                                {
                                    //if not, rename this layer to it, and now it exists.
                                    layer.Name = TC_US_Environmental_Layers.Wetland;

                                }
                                else
                                {
                                    //if the layer does exist in the drawing.....find all objects on the survey layer, and put them on the TC Standard layer.
                                    foreach (ObjectId objectId in blockTableRecord)
                                    {
                                        Entity entity = thisTransaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
                                        if (entity.Layer != null)
                                        {
                                            if (entity.Layer == layer.Name)
                                            {
                                                {
                                                    entity.Layer = TC_US_Environmental_Layers.Wetland;

                                                }
                                            }
                                        }
                                    }
                                }
                                break;

                            case "E_TER_WETLAND AREA":
                                //see if the TC standard layer already exists.
                                if (layerTable.Has(TC_US_Environmental_Layers.Wetland) == false)
                                {
                                    //if not, rename this layer to it, and now it exists.
                                    layer.Name = TC_US_Environmental_Layers.Wetland;

                                }
                                else
                                {
                                    //if the layer does exist in the drawing.....find all objects on the survey layer, and put them on the TC Standard layer.
                                    foreach (ObjectId objectId in blockTableRecord)
                                    {
                                        Entity entity = thisTransaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
                                        if (entity.Layer != null)
                                        {
                                            if (entity.Layer == layer.Name)
                                            {
                                                {
                                                    entity.Layer = TC_US_Environmental_Layers.Wetland;

                                                }
                                            }
                                        }
                                    }
                                }
                                break;

                            case "Stream_Polygon":
                                //see if the TC standard layer already exists.
                                if (layerTable.Has(TC_US_Environmental_Layers.Creek) == false)
                                {
                                    //if not, rename this layer to it, and now it exists.
                                    layer.Name = TC_US_Environmental_Layers.Creek;

                                }
                                else
                                {
                                    //if the layer does exist in the drawing.....find all objects on the survey layer, and put them on the TC Standard layer.
                                    foreach (ObjectId objectId in blockTableRecord)
                                    {
                                        Entity entity = thisTransaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
                                        if (entity.Layer != null)
                                        {
                                            if (entity.Layer == layer.Name)
                                            {
                                                {
                                                    entity.Layer = TC_US_Environmental_Layers.Creek;

                                                }
                                            }
                                        }
                                    }
                                }
                                break;

                            case "E_TRA_Guardrail":
                                //see if the TC standard layer already exists.
                                if (layerTable.Has(TC_US_ExistingConditions_Layers.GuardRail) == false)
                                {
                                    //if not, rename this layer to it, and now it exists.
                                    layer.Name = TC_US_ExistingConditions_Layers.GuardRail;

                                }
                                else
                                {
                                    //if the layer does exist in the drawing.....find all objects on the survey layer, and put them on the TC Standard layer.
                                    foreach (ObjectId objectId in blockTableRecord)
                                    {
                                        Entity entity = thisTransaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
                                        if (entity.Layer != null)
                                        {
                                            if (entity.Layer == layer.Name)
                                            {
                                                {
                                                    entity.Layer = TC_US_ExistingConditions_Layers.GuardRail;

                                                }
                                            }
                                        }
                                    }
                                }
                                break;

                            case "E_FEA_Fence":
                                //see if the TC standard layer already exists.
                                if (layerTable.Has(TC_US_ExistingConditions_Layers.FENCE) == false)
                                {
                                    //if not, rename this layer to it, and now it exists.
                                    layer.Name = TC_US_ExistingConditions_Layers.FENCE;

                                }
                                else
                                {
                                    //if the layer does exist in the drawing.....find all objects on the survey layer, and put them on the TC Standard layer.
                                    foreach (ObjectId objectId in blockTableRecord)
                                    {
                                        Entity entity = thisTransaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
                                        if (entity.Layer != null)
                                        {
                                            if (entity.Layer == layer.Name)
                                            {
                                                {
                                                    entity.Layer = TC_US_ExistingConditions_Layers.FENCE;

                                                }
                                            }
                                        }
                                    }
                                }
                                break;

                            case "E_UTL_GuyWire":
                                //see if the TC standard layer already exists.
                                if (layerTable.Has(TC_US_ExistingConditions_Layers.GuyWireAnchorSymbol) == false)
                                {
                                    //if not, rename this layer to it, and now it exists.
                                    layer.Name = TC_US_ExistingConditions_Layers.GuyWireAnchorSymbol;

                                }
                                else
                                {
                                    //if the layer does exist in the drawing.....find all objects on the survey layer, and put them on the TC Standard layer.
                                    foreach (ObjectId objectId in blockTableRecord)
                                    {
                                        Entity entity = thisTransaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
                                        if (entity.Layer != null)
                                        {
                                            if (entity.Layer == layer.Name)
                                            {
                                                {
                                                    entity.Layer = TC_US_ExistingConditions_Layers.GuyWireAnchorSymbol;

                                                }
                                            }
                                        }
                                    }
                                }
                                break;

                            case "E_TER_TopBerm":
                                //see if the TC standard layer already exists.
                                if (layerTable.Has(TC_US_ExistingConditions_Layers.TopofBerm) == false)
                                {
                                    //if not, rename this layer to it, and now it exists.
                                    layer.Name = TC_US_ExistingConditions_Layers.TopofBerm;

                                }
                                else
                                {
                                    //if the layer does exist in the drawing.....find all objects on the survey layer, and put them on the TC Standard layer.
                                    foreach (ObjectId objectId in blockTableRecord)
                                    {
                                        Entity entity = thisTransaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
                                        if (entity.Layer != null)
                                        {
                                            if (entity.Layer == layer.Name)
                                            {
                                                {
                                                    entity.Layer = TC_US_ExistingConditions_Layers.TopofBerm;

                                                }
                                            }
                                        }
                                    }
                                }
                                break;

                            case "E_TER_ToeBerm":
                                //see if the TC standard layer already exists.
                                if (layerTable.Has(TC_US_ExistingConditions_Layers.ToeofBerm) == false)
                                {
                                    //if not, rename this layer to it, and now it exists.
                                    layer.Name = TC_US_ExistingConditions_Layers.ToeofBerm;

                                }
                                else
                                {
                                    //if the layer does exist in the drawing.....find all objects on the survey layer, and put them on the TC Standard layer.
                                    foreach (ObjectId objectId in blockTableRecord)
                                    {
                                        Entity entity = thisTransaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
                                        if (entity.Layer != null)
                                        {
                                            if (entity.Layer == layer.Name)
                                            {
                                                {
                                                    entity.Layer = TC_US_ExistingConditions_Layers.ToeofBerm;

                                                }
                                            }
                                        }
                                    }
                                }
                                break;


                            case "E_TER_TopSlope":
                                //see if the TC standard layer already exists.
                                if (layerTable.Has(TC_US_ExistingConditions_Layers.TopofSlope) == false)
                                {
                                    //if not, rename this layer to it, and now it exists.
                                    layer.Name = TC_US_ExistingConditions_Layers.TopofSlope;

                                }
                                else
                                {
                                    //if the layer does exist in the drawing.....find all objects on the survey layer, and put them on the TC Standard layer.
                                    foreach (ObjectId objectId in blockTableRecord)
                                    {
                                        Entity entity = thisTransaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
                                        if (entity.Layer != null)
                                        {
                                            if (entity.Layer == layer.Name)
                                            {
                                                {
                                                    entity.Layer = TC_US_ExistingConditions_Layers.TopofSlope;

                                                }
                                            }
                                        }
                                    }
                                }
                                break;

                            case "E_TER_ToeSlope":
                                //see if the TC standard layer already exists.
                                if (layerTable.Has(TC_US_ExistingConditions_Layers.ToeofSlope) == false)
                                {
                                    //if not, rename this layer to it, and now it exists.
                                    layer.Name = TC_US_ExistingConditions_Layers.ToeofSlope;

                                }
                                else
                                {
                                    //if the layer does exist in the drawing.....find all objects on the survey layer, and put them on the TC Standard layer.
                                    foreach (ObjectId objectId in blockTableRecord)
                                    {
                                        Entity entity = thisTransaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
                                        if (entity.Layer != null)
                                        {
                                            if (entity.Layer == layer.Name)
                                            {
                                                {
                                                    entity.Layer = TC_US_ExistingConditions_Layers.ToeofSlope;

                                                }
                                            }
                                        }
                                    }
                                }
                                break;

                            //case "E_TER_RockArea":
                            //    break;

                            //case "E_TER_AgField":
                            //    break;

                            //case "E_FEA_Tank":
                            //    break;

                            //case "E_FEA_WallRock":
                            //    break;

                            //case "E_FEA_Gate":
                            //    break;

                            case "E_FEA_DrainTile":
                                //see if the TC standard layer already exists.
                                if (layerTable.Has(TC_US_ExistingConditions_Layers.DrainBarDitch) == false)
                                {
                                    //if not, rename this layer to it, and now it exists.
                                    layer.Name = TC_US_ExistingConditions_Layers.DrainBarDitch;

                                }
                                else
                                {
                                    //if the layer does exist in the drawing.....find all objects on the survey layer, and put them on the TC Standard layer.
                                    foreach (ObjectId objectId in blockTableRecord)
                                    {
                                        Entity entity = thisTransaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
                                        if (entity.Layer != null)
                                        {
                                            if (entity.Layer == layer.Name)
                                            {
                                                {
                                                    entity.Layer = TC_US_ExistingConditions_Layers.DrainBarDitch;

                                                }
                                            }
                                        }
                                    }
                                }
                                break;

                            //case "E_FEA_Bollard_PNT":
                            //    break;

                            //case "E_TRA_Bridge":
                            //    break;

                            //case "E_FEA_WallRetain":
                            //    break;

                            case "E_FEA_Structure":
                                //see if the TC standard layer already exists.
                                if (layerTable.Has(TC_US_ExistingConditions_Layers.Struct) == false)
                                {
                                    //if not, rename this layer to it, and now it exists.
                                    layer.Name = TC_US_ExistingConditions_Layers.Struct;

                                }
                                else
                                {
                                    //if the layer does exist in the drawing.....find all objects on the survey layer, and put them on the TC Standard layer.
                                    foreach (ObjectId objectId in blockTableRecord)
                                    {
                                        Entity entity = thisTransaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
                                        if (entity.Layer != null)
                                        {
                                            if (entity.Layer == layer.Name)
                                            {
                                                {
                                                    entity.Layer = TC_US_ExistingConditions_Layers.Struct;

                                                }
                                            }
                                        }
                                    }
                                }
                                break;

                            case "E_FEA_StmDrain":
                                //see if the TC standard layer already exists.
                                if (layerTable.Has(TC_US_ExistingConditions_Layers.StormSewer) == false)
                                {
                                    //if not, rename this layer to it, and now it exists.
                                    layer.Name = TC_US_ExistingConditions_Layers.StormSewer;

                                }
                                else
                                {
                                    //if the layer does exist in the drawing.....find all objects on the survey layer, and put them on the TC Standard layer.
                                    foreach (ObjectId objectId in blockTableRecord)
                                    {
                                        Entity entity = thisTransaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
                                        if (entity.Layer != null)
                                        {
                                            if (entity.Layer == layer.Name)
                                            {
                                                {
                                                    entity.Layer = TC_US_ExistingConditions_Layers.StormSewer;

                                                }
                                            }
                                        }
                                    }
                                }
                                break;

                            case "E_UTL_StmSewer":
                                //see if the TC standard layer already exists.
                                if (layerTable.Has(TC_US_ExistingConditions_Layers.StormSewer) == false)
                                {
                                    //if not, rename this layer to it, and now it exists.
                                    layer.Name = TC_US_ExistingConditions_Layers.StormSewer;

                                }
                                else
                                {
                                    //if the layer does exist in the drawing.....find all objects on the survey layer, and put them on the TC Standard layer.
                                    foreach (ObjectId objectId in blockTableRecord)
                                    {
                                        Entity entity = thisTransaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
                                        if (entity.Layer != null)
                                        {
                                            if (entity.Layer == layer.Name)
                                            {
                                                {
                                                    entity.Layer = TC_US_ExistingConditions_Layers.StormSewer;

                                                }
                                            }
                                        }
                                    }
                                }
                                break;

                            case "E_UTL_Water":
                                //see if the TC standard layer already exists.
                                if (layerTable.Has(TC_US_ExistingConditions_Layers.WaterLine) == false)
                                {
                                    //if not, rename this layer to it, and now it exists.
                                    layer.Name = TC_US_ExistingConditions_Layers.WaterLine;

                                }
                                else
                                {
                                    //if the layer does exist in the drawing.....find all objects on the survey layer, and put them on the TC Standard layer.
                                    foreach (ObjectId objectId in blockTableRecord)
                                    {
                                        Entity entity = thisTransaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
                                        if (entity.Layer != null)
                                        {
                                            if (entity.Layer == layer.Name)
                                            {
                                                {
                                                    entity.Layer = TC_US_ExistingConditions_Layers.StormSewer;

                                                }
                                            }
                                        }
                                    }
                                }
                                break;

                            //case "E_TRA_Edge_Grav":
                            //    break;

                            //case "E_TRA_RD_EdgeDirt":
                            //    break;

                            case "E_STR_Bldg":
                                //see if the TC standard layer already exists.
                                if (layerTable.Has(TC_US_ExistingConditions_Layers.Struct) == false)
                                {
                                    //if not, rename this layer to it, and now it exists.
                                    layer.Name = TC_US_ExistingConditions_Layers.Struct;

                                }
                                else
                                {
                                    //if the layer does exist in the drawing.....find all objects on the survey layer, and put them on the TC Standard layer.
                                    foreach (ObjectId objectId in blockTableRecord)
                                    {
                                        Entity entity = thisTransaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
                                        if (entity.Layer != null)
                                        {
                                            if (entity.Layer == layer.Name)
                                            {
                                                {
                                                    entity.Layer = TC_US_ExistingConditions_Layers.Struct;

                                                }
                                            }
                                        }
                                    }
                                }
                                break;
                        }
                    }
                    thisTransaction.Commit();
                    //thisEditor.WriteMessage($"\r\n{numberOfEdits} edits have been made in the {thisDrawing.Name}.");
                }
            }
        }
    }
}
