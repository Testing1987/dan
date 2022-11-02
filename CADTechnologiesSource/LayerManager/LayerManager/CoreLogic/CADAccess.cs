using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Microsoft.Win32;
using System.Collections.Generic;
using System.IO;
using System.Windows;

namespace LayerManager.CoreLogic
{
    public class CADAccess
    {
        /// <summary>
        /// Creates a list of drawing files by asking the user to select them with an <see cref="OpenFileDialog"/> and returns that list.
        /// </summary>
        /// <returns></returns>
        public IList<string> AddtoDrawingList()
        {
            var items = new List<string>();

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Select Drawings";
            openFileDialog.Filter = "Drawing (*.dwg) | *.dwg";
            openFileDialog.Multiselect = true;
            if (openFileDialog.ShowDialog() == true)
            {
                foreach (string file in openFileDialog.FileNames)
                {
                    items.Add(file);
                }
            }
            return items;
        }

        internal void UpdateLayerProperties(string drawing, string layerName, bool isXrefLayer, bool isFrozen, bool isOff, bool isLocked, bool isPlottable, bool adjustTransparency, int transparency, bool adjustLinetype, string linetype, bool adjustLineweight, LineWeight newLineWeight, bool adjustColor, Color newColor)
        {
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock lock1 = thisDrawing.LockDocument())
            {
                using (Transaction trans1 = thisDrawing.TransactionManager.StartTransaction())
                {
                    if (File.Exists(drawing) && drawing != null)
                    {
                        try
                        {
                            using (Database database2 = new Database(false, true))
                            {
                                database2.ReadDwgFile(drawing, FileOpenMode.OpenForReadAndWriteNoShare, true, "");
                                database2.CloseInput(true);
                                if (isXrefLayer == true)
                                {
                                    database2.ResolveXrefs(useThreadEngine: true, doNewOnly: false);
                                }
                                HostApplicationServices.WorkingDatabase = database2;

                                using (Transaction transaction2 = database2.TransactionManager.StartTransaction())
                                {
                                    //Access the Layer Table
                                    LayerTable layerTable = transaction2.GetObject(database2.LayerTableId, OpenMode.ForWrite) as LayerTable;
                                    if (layerTable.Has(layerName) == true)
                                    {
                                        //Iterate through the layer table and find the layer we want...
                                        foreach (ObjectId objectId in layerTable)
                                        {
                                            LayerTableRecord layerTableRecord = transaction2.GetObject(objectId, OpenMode.ForRead) as LayerTableRecord;
                                            LinetypeTable linetypeTable = transaction2.GetObject(database2.LinetypeTableId, OpenMode.ForRead) as LinetypeTable;
                                            //If the layer name matches the passed in value...
                                            if (layerTableRecord.Name.ToLower() == layerName)
                                            {
                                                //Make sure it's not null
                                                if (layerTableRecord != null)
                                                {
                                                    //then upgrade the LayerTableRecord for writing...
                                                    layerTableRecord.UpgradeOpen();

                                                    //Update the properties
                                                    layerTableRecord.IsFrozen = isFrozen;
                                                    layerTableRecord.IsOff = isOff;
                                                    layerTableRecord.IsLocked = isLocked;
                                                    layerTableRecord.IsPlottable = isPlottable;

                                                    #region Transparency
                                                    {
                                                        try
                                                        {
                                                            if (adjustTransparency == true)
                                                            {
                                                                byte alpha = (byte)(255 * (100 - transparency) / 100);
                                                                Transparency newTransparency = new Transparency(alpha);
                                                                layerTableRecord.Transparency = newTransparency;
                                                                layerTableRecord.Transparency = newTransparency;
                                                            }
                                                        }
                                                        catch (System.Exception ex)
                                                        {
                                                            MessageBox.Show(ex.Message);
                                                        }
                                                    }
                                                    #endregion

                                                    #region Linetype
                                                    try
                                                    {
                                                        if (adjustLinetype == true)
                                                        {
                                                            if (linetypeTable.Has(linetype.ToLower()))
                                                            {
                                                                layerTableRecord.LinetypeObjectId = linetypeTable[linetype];
                                                                //DatabaseHelpers.GetLinetypeFromSideDatabase(transaction2, database2, linetype);
                                                            }
                                                            else
                                                            {
                                                                MessageBox.Show($"{Path.GetFileName(database2.OriginalFileName)} does not contain a linetype named '{linetype}'. Please make sure the linetype has been loaded into the drawing.",
                                                                    "Linetype not found",
                                                                    MessageBoxButton.OK,
                                                                    MessageBoxImage.Information);
                                                            }
                                                        }
                                                    }
                                                    catch (System.Exception ex)
                                                    {
                                                        MessageBox.Show(ex.Message);
                                                    }
                                                    #endregion

                                                    #region Lineweight
                                                    try
                                                    {
                                                        if (adjustLineweight)
                                                        {
                                                            layerTableRecord.LineWeight = newLineWeight;
                                                        }
                                                    }
                                                    catch (System.Exception ex)
                                                    {
                                                        MessageBox.Show(ex.Message);
                                                    }
                                                    #endregion

                                                    #region Color
                                                    try
                                                    {
                                                        layerTableRecord.Color = newColor;
                                                    }
                                                    catch (System.Exception ex)
                                                    {
                                                        MessageBox.Show(ex.Message + "The value you entered for True Color is not correct. Please use an RGB style value. It must be three numbers between 1-255, separated by commas. For example: '125, 17, 255'",
                                                            "Invalid True Color",
                                                            MessageBoxButton.OK,
                                                            MessageBoxImage.Warning);
                                                        return;
                                                    }
                                                    #endregion
                                                    layerTableRecord.DowngradeOpen();
                                                }
                                            }
                                        }
                                        //Commit the transaction and save the drawing.
                                        transaction2.Commit();
                                    }
                                    HostApplicationServices.WorkingDatabase = thisDrawing.Database;
                                    database2.SaveAs(drawing, true, DwgVersion.Current, thisDrawing.Database.SecurityParameters);
                                }
                            }
                        }
                        catch (System.Exception)
                        {
                            MessageBox.Show($"Problem accessing database of {drawing}",
                                "The drawing could not be accessed.",
                                MessageBoxButton.OK,
                                MessageBoxImage.Error);
                        }
                    }
                }
            }
        }
    }
}














