using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using CADTechnologiesSource.All.AutoCADHelpers;
using CADTechnologiesSource.All.Models;
using CADTechnologiesSource.All.PropertyHelpers;
using LayerComparison.Core.Constants;
using LayerComparison.Core.Models;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows;

namespace LayerComparison.Core.CoreLogic
{
    public class DataAccess
    {
        #region Save/Load
        public void SaveComparison(List<string> savedComparison)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Layer Comparison (*.lcomp)|*.lcomp";
            saveFileDialog.Title = "Save the current layer comparison.";
            saveFileDialog.RestoreDirectory = true;
            Nullable<bool> result = saveFileDialog.ShowDialog();

            if (result == true)
            {
                System.IO.File.WriteAllLines(saveFileDialog.FileName, savedComparison);
            }
            SaveRecentComparisonList(saveFileDialog.FileName);
        }

        public void SaveRecentComparisonList(string recentFileName)
        {
            string path = FilePaths.RecentComparisonFile;
            if (!File.Exists(path))
            {
                //Create the folder
                System.IO.Directory.CreateDirectory(FilePaths.RecentComparisonPath);
                // Create a file to write to.
                using (StreamWriter sw = File.AppendText(path))
                {
                    sw.WriteLine(recentFileName);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(path))
                {
                    sw.WriteLine(recentFileName);
                }
            }
        }
        #endregion

        #region Source and Target Population
        /// <summary>
        /// Creates a list of drawing files by asking the user to select them with an <see cref="OpenFileDialog"/> and returns that list.
        /// </summary>
        /// <returns></returns>
        public IList<TargetDrawingModel> AddtoTargetDrawingList()
        {
            var items = new List<TargetDrawingModel>();

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Select The Target Drawings To Add To Add To The Comparison";
            openFileDialog.Filter = "Drawing (*.dwg) | *.dwg";
            openFileDialog.Multiselect = true;
            if (openFileDialog.ShowDialog() == true)
            {
                foreach (string file in openFileDialog.FileNames)
                {
                    TargetDrawingModel targetDrawing = new TargetDrawingModel();
                    targetDrawing.DrawingPath = file;
                    items.Add(targetDrawing);
                }
            }
            return items;
        }

        /// <summary>
        /// Opens an OpenFileDialog so the user can select a single .dwg file, and returns the filepath of that file as a string.
        /// </summary>
        /// <returns></returns>
        public string GetSource()
        {
            string filePath = "";
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Select The Source Drawing To Be Used For This Comparison";
            openFileDialog.Filter = "Drawing (*.dwg) | *.dwg";
            openFileDialog.Multiselect = false;
            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    if (openFileDialog.FileName != null && openFileDialog.FileName.Contains(".dwg"))
                    {
                        filePath = openFileDialog.FileName;
                        return filePath;
                    }
                    else
                    {
                        return filePath = MessagesAndNotifications.SourceDrawingNotSet;
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("A valid source drawing could not be determined.", "No source drawing found.", MessageBoxButton.OK, MessageBoxImage.Error);
                    return filePath = MessagesAndNotifications.SourceDrawingNotSet;
                }
            }
            else
            {
                return filePath = MessagesAndNotifications.SourceDrawingNotSet;
            }
        }
        #endregion

        #region AutoCAD Database Access

        /// <summary>
        /// Creates an observable collection of type <see cref="LayerComparisonLayerModel"/> and returns it to the caller.
        /// </summary>
        /// <param name="s">The file path of the drawing to build the collection from.</param>
        /// <returns></returns>
        public ObservableCollection<LayerComparisonLayerModel> BuildDrawingLayersCollection(string s)
        {
            ObservableCollection<LayerComparisonLayerModel> output = new ObservableCollection<LayerComparisonLayerModel>();

            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor editor = thisDrawing.Editor;
            using (DocumentLock lock1 = thisDrawing.LockDocument())
            {
                using (Transaction trans1 = thisDrawing.TransactionManager.StartTransaction())
                {
                    if (File.Exists(s) && s != null)
                        try
                        {
                            using (Database database2 = new Database(false, true))
                            {
                                database2.ReadDwgFile(s, FileOpenMode.OpenForReadAndWriteNoShare, true, "");
                                database2.CloseInput(true);
                                //database2.ResolveXrefs(useThreadEngine: true, doNewOnly: false);

                                HostApplicationServices.WorkingDatabase = database2;

                                using (Transaction trans2 = database2.TransactionManager.StartTransaction())
                                {
                                    LayerTable layerTable2 = trans2.GetObject(database2.LayerTableId, OpenMode.ForRead) as LayerTable;

                                    foreach (ObjectId id in layerTable2)
                                    {
                                        LayerTableRecord layerTableRecord = trans2.GetObject(id, OpenMode.ForRead) as LayerTableRecord;
                                        LinetypeTableRecord lineTypeRecord = trans2.GetObject(layerTableRecord.LinetypeObjectId, OpenMode.ForRead) as LinetypeTableRecord;

                                        if (layerTableRecord != null && lineTypeRecord != null)
                                        {
                                            LayerComparisonLayerModel layer = new LayerComparisonLayerModel();

                                            layer.DrawingPath = s;
                                            layer.Name = layerTableRecord.Name;
                                            layer.OnOff = layerTableRecord.IsOff ? "Off" : "On";
                                            layer.Freeze = layerTableRecord.IsFrozen ? "Frozen" : "Thawed";
                                            #region Color
                                            //Get the layer color by creating an autocad color, setting it to the value of the color found on the layer, and converting it as necessary.
                                            Autodesk.AutoCAD.Colors.Color color1 = layerTableRecord.Color;
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
                                                layer.Color = color_string;
                                            }
                                            else
                                            {
                                                //Return the index color.
                                                layer.Color = color_string;
                                            }
                                            #endregion
                                            layer.Linetype = lineTypeRecord.Name;
                                            layer.Lineweight = layerTableRecord.LineWeight.ToString();
                                            layer.Plot = layerTableRecord.IsPlottable ? "Yes" : "No";
                                            #region Transparency
                                            if (layerTableRecord.Transparency.IsByAlpha)
                                            {
                                                int percentage = (int)(((255 - layerTableRecord.Transparency.Alpha) * 100) / 255);
                                                layer.Transparency = percentage.ToString();
                                            }
                                            else
                                            {
                                                layer.Transparency = "0";
                                            }
                                            #endregion

                                            output.Add(layer);
                                        }
                                    }
                                }
                                HostApplicationServices.WorkingDatabase = thisDrawing.Database;
                            }
                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                }
            }

            //Return the list of AutoCAD models to the ObversableCollection
            return output;
        }

        /// <summary>
        /// Creates an observable collection of type <see cref="ViewportLayerModel"/> and returns it to the caller.
        /// </summary>
        /// <param name="s">The file path of the drawing to build the collection from.</param>
        /// <returns></returns>
        public ObservableCollection<LayerComparisonViewportLayerModel> BuildDrawingViewportLayersCollection(string s)
        {
            ObservableCollection<LayerComparisonViewportLayerModel> output = new ObservableCollection<LayerComparisonViewportLayerModel>();
            ObjectIdCollection viewportCollection = new ObjectIdCollection();

            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor editor = thisDrawing.Editor;
            using (DocumentLock lock1 = thisDrawing.LockDocument())
            {
                using (Transaction trans1 = thisDrawing.TransactionManager.StartTransaction())
                {
                    if (File.Exists(s) && s != null)
                    {
                        try
                        {
                            using (Database database2 = new Database(false, true))
                            {
                                database2.ReadDwgFile(s, FileOpenMode.OpenForReadAndWriteNoShare, true, "");
                                database2.CloseInput(true);
                                //database2.ResolveXrefs(useThreadEngine: true, doNewOnly: false);

                                HostApplicationServices.WorkingDatabase = database2;

                                using (Transaction trans2 = database2.TransactionManager.StartTransaction())
                                {
                                    //access the layer table of transaction2
                                    LayerTable layerTable2 = trans2.GetObject(database2.LayerTableId, OpenMode.ForRead) as LayerTable;

                                    //access the Layout DBDictionary of database 2 to build a collection of viewports
                                    DBDictionary dBLayoutDictionary = trans2.GetObject(database2.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                                    foreach (DBDictionaryEntry entry in dBLayoutDictionary)
                                    {
                                        Layout layout = trans2.GetObject(entry.Value, OpenMode.ForRead) as Layout;
                                        if (entry.Key != "Model")
                                        {
                                            viewportCollection = layout.GetViewports();
                                            viewportCollection.RemoveAt((0));
                                        }
                                    }

                                    foreach (ObjectId viewportID in viewportCollection)
                                    {
                                        Viewport viewport = trans2.GetObject(viewportID, OpenMode.ForRead, false, true) as Viewport;

                                        foreach (ObjectId id2 in layerTable2)
                                        {
                                            //Set up a LayerTableRecord and a LinetypeTableRecord for each id2
                                            LayerTableRecord vpLayer = trans2.GetObject(id2, OpenMode.ForRead) as LayerTableRecord;
                                            LinetypeTableRecord vpLineTypeRecord = trans2.GetObject(vpLayer.LinetypeObjectId, OpenMode.ForRead) as LinetypeTableRecord;

                                            if (vpLayer != null && vpLineTypeRecord != null)
                                            {
                                                //Get the viewport layer overrides
                                                LayerViewportProperties layerViewportProperties = vpLayer.GetViewportOverrides(viewport.ObjectId);
                                                //Get the name of the linetype, based on ObjectId. This will be used later to set the value on the layer model.
                                                LinetypeTableRecord linetypeTableRecord = trans2.GetObject(layerViewportProperties.LinetypeObjectId, OpenMode.ForRead) as LinetypeTableRecord;
                                                string viewportLinetypeName = linetypeTableRecord.Name;

                                                //Create a model to contain the information
                                                LayerComparisonViewportLayerModel viewportLayerModel = new LayerComparisonViewportLayerModel();

                                                //Set the properties in the model
                                                viewportLayerModel.ViewportLayer = viewport.Layer;
                                                viewportLayerModel.DrawingPath = s;
                                                viewportLayerModel.Name = vpLayer.Name;
                                                viewportLayerModel.ViewportPosition = viewport.CenterPoint;
                                                viewportLayerModel.ViewportFreeze = viewport.IsLayerFrozenInViewport(vpLayer.Id) ? "Frozen" : "Thawed";
                                                viewportLayerModel.ViewportLinetype = viewportLinetypeName;
                                                viewportLayerModel.ViewportLineweight = layerViewportProperties.LineWeight.ToString();
                                                #region Color
                                                //Get the layer color by creating an autocad color, setting it to the value of the color found on the layer, and converting it as necessary.
                                                Autodesk.AutoCAD.Colors.Color color1 = layerViewportProperties.Color;
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
                                                    viewportLayerModel.ViewportColor = color_string;
                                                }
                                                else
                                                {
                                                    //Return the index color.
                                                    viewportLayerModel.ViewportColor = color_string;
                                                }
                                                #endregion
                                                #region Transparency
                                                if (layerViewportProperties.Transparency.IsByAlpha)
                                                {
                                                    int percentage = (int)(((255 - layerViewportProperties.Transparency.Alpha) * 100) / 255);
                                                    viewportLayerModel.ViewportTransparency = percentage.ToString();
                                                }
                                                else
                                                {
                                                    viewportLayerModel.ViewportTransparency = "0";
                                                }
                                                #endregion
                                                output.Add(viewportLayerModel);
                                            }
                                        }
                                    }

                                    HostApplicationServices.WorkingDatabase = thisDrawing.Database;
                                    return output;
                                }
                            }
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("The drawing could not be accessed.", $"Problem accessing database of {s}", MessageBoxButton.OK, MessageBoxImage.Error);
                            throw;
                        }
                    }
                    return output;
                }
            }
        }

        /// <summary>
        /// Searches the given drawing for the given layer, and returns true if the layer is found.
        /// </summary>
        /// <param name="path">the drawing path</param>
        /// <param name="layerName">the layer to look for</param>
        /// <returns></returns>
        public bool FindDrawingLayer(string path, string layerName)
        {
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor editor = thisDrawing.Editor;
            using (DocumentLock lock1 = thisDrawing.LockDocument())
            {
                using (Transaction trans1 = thisDrawing.TransactionManager.StartTransaction())
                {
                    if (File.Exists(path) && path != null)
                        try
                        {
                            using (Database database2 = new Database(false, true))
                            {
                                database2.ReadDwgFile(path, FileOpenMode.OpenForReadAndWriteNoShare, true, "");
                                database2.CloseInput(true);
                                //database2.ResolveXrefs(useThreadEngine: true, doNewOnly: false);

                                HostApplicationServices.WorkingDatabase = database2;

                                using (Transaction trans2 = database2.TransactionManager.StartTransaction())
                                {
                                    LayerTable layerTable2 = trans2.GetObject(database2.LayerTableId, OpenMode.ForRead) as LayerTable;

                                    if (layerTable2.Has(layerName))
                                    {
                                        HostApplicationServices.WorkingDatabase = thisDrawing.Database;
                                        return true;
                                    }
                                }
                            }
                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                }
            }
            HostApplicationServices.WorkingDatabase = thisDrawing.Database;
            return false;
        }

        /// <summary>
        /// searches a <see cref="DrawingModel"/> for the given layer. If found, returns true/>.
        /// </summary>
        /// <param name="drawingModel"></param>
        /// <param name="layer"></param>
        /// <returns></returns>
        public bool FindDrawingModelLayer(DrawingModel drawingModel, string layer)
        {
            if (drawingModel.Layers.Contains(layer))
            {
                return true;
            }
            return false;
        }

        /// <summary>
        /// Creates a drawing model and populates a <see cref="List{string}"/> contained within the model with all of the layers found inside the given drawing.
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public DrawingModel CreateDrawingModel(string path)
        {
            DrawingModel drawingModel = new DrawingModel();
            drawingModel.DrawingPath = path;

            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor editor = thisDrawing.Editor;
            using (DocumentLock lock1 = thisDrawing.LockDocument())
            {
                using (Transaction trans1 = thisDrawing.TransactionManager.StartTransaction())
                {
                    if (File.Exists(path) && path != null)
                        try
                        {
                            using (Database database2 = new Database(false, true))
                            {
                                database2.ReadDwgFile(path, FileOpenMode.OpenForReadAndWriteNoShare, true, "");
                                database2.CloseInput(true);
                                //database2.ResolveXrefs(useThreadEngine: true, doNewOnly: false);

                                HostApplicationServices.WorkingDatabase = database2;

                                using (Transaction trans2 = database2.TransactionManager.StartTransaction())
                                {
                                    LayerTable layerTable2 = trans2.GetObject(database2.LayerTableId, OpenMode.ForRead) as LayerTable;

                                    foreach (ObjectId id in layerTable2)
                                    {
                                        LayerTableRecord layerTableRecord = trans2.GetObject(id, OpenMode.ForRead) as LayerTableRecord;
                                        LinetypeTableRecord lineTypeRecord = trans2.GetObject(layerTableRecord.LinetypeObjectId, OpenMode.ForRead) as LinetypeTableRecord;

                                        if (layerTableRecord != null && lineTypeRecord != null)
                                        {
                                            drawingModel.Layers.Add(layerTableRecord.Name);
                                        }
                                    }
                                }
                            }
                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                }
            }
            HostApplicationServices.WorkingDatabase = thisDrawing.Database;

            return drawingModel;
        }
        #endregion

        #region Fix Single Conflicts
        /// <summary>
        /// Corrects layer property conflicts.
        /// </summary>
        /// <param name="path">The drawing path of the .dwg file.</param>
        /// <param name="layerName">The name of the layer to be updated</param>
        /// <param name="property">The property to update.</param>
        /// <param name="desiredSetting">The layer property you want to apply to the layer in question.</param>
        /// <summary>
        /// Corrects layer property conflicts.
        /// </summary>
        /// <param name="path">The drawing path of the .dwg file.</param>
        /// <param name="layerName">The name of the layer to be updated</param>
        /// <param name="property">The property to update.</param>
        /// <param name="desiredSetting">The layer property you want to apply to the layer in question.</param>
        public void FixLayerConflict(string path, string layerName, string property, string desiredSetting)
        {
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor editor = thisDrawing.Editor;
            string shortLayerName = "";
            using (DocumentLock lock1 = thisDrawing.LockDocument())
            {
                using (Transaction trans1 = thisDrawing.TransactionManager.StartTransaction())
                {
                    if (File.Exists(path) && !path.Contains(MessagesAndNotifications.SourceDrawingNotSet) && path != null)
                    {
                        try
                        {
                            using (Database database2 = new Database(false, true))
                            {
                                database2.ReadDwgFile(path, FileOpenMode.OpenForReadAndWriteNoShare, true, "");
                                database2.CloseInput(true);
                                //database2.ResolveXrefs(useThreadEngine: true, doNewOnly: false);

                                HostApplicationServices.WorkingDatabase = database2;

                                using (Transaction trans2 = database2.TransactionManager.StartTransaction())
                                {
                                    DrawingSettingHelpers drawingSettingHelpers = new DrawingSettingHelpers();
                                    if (drawingSettingHelpers.IsVisretainOn() == false)
                                    {
                                        MessageBox.Show($"VISRETAIN is not enabled on {path}. Changes to it's xref layers will not be saved.");
                                    }
                                    LayerTable layerTable2 = trans2.GetObject(database2.LayerTableId, OpenMode.ForWrite) as LayerTable;

                                    foreach (ObjectId id in layerTable2)
                                    {
                                        LayerTableRecord layerTableRecord = trans2.GetObject(id, OpenMode.ForWrite) as LayerTableRecord;
                                        LinetypeTable linetypeTable = trans2.GetObject(database2.LinetypeTableId, OpenMode.ForRead) as LinetypeTable;
                                        LinetypeTableRecord lineTypeRecord = trans2.GetObject(layerTableRecord.LinetypeObjectId, OpenMode.ForWrite) as LinetypeTableRecord;

                                        if (layerTableRecord != null && lineTypeRecord != null && layerTableRecord.Name == layerName)
                                        {
                                            switch (property)
                                            {
                                                case "OnOff":
                                                    if (desiredSetting == "Off")
                                                    {
                                                        layerTableRecord.IsOff = true;
                                                    }
                                                    else
                                                    {
                                                        layerTableRecord.IsOff = false;
                                                    }
                                                    break;

                                                case "FreezeThaw":
                                                    if (desiredSetting == "Frozen")
                                                    {
                                                        layerTableRecord.IsFrozen = true;
                                                    }
                                                    else
                                                    {
                                                        layerTableRecord.IsFrozen = false;
                                                    }
                                                    break;

                                                case "Color":
                                                    #region Color Crazyness
                                                    Autodesk.AutoCAD.Colors.Color color1 = new Autodesk.AutoCAD.Colors.Color();
                                                    string color_string = desiredSetting.ToLower();

                                                    if (color_string.ToLower().Contains(",") == true)
                                                    {
                                                        int idx1 = color_string.IndexOf(",", 0);
                                                        byte R = Convert.ToByte(color_string.Substring(0, idx1));
                                                        int idx2 = color_string.IndexOf(",", idx1 + 1);
                                                        byte G = Convert.ToByte(color_string.Substring(idx1 + 1, idx2 - idx1 - 1));
                                                        byte B = Convert.ToByte(color_string.Substring(idx2 + 1, color_string.Length - idx2 - 1));
                                                        color1 = Autodesk.AutoCAD.Colors.Color.FromRgb(R, G, B);

                                                        //Return the truecolor value.
                                                        layerTableRecord.Color = color1;
                                                    }
                                                    else
                                                    {
                                                        switch (color_string)
                                                        {
                                                            case "byblock":
                                                                layerTableRecord.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_ByBlock);
                                                                break;
                                                            case "red":
                                                                layerTableRecord.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_Red);
                                                                break;
                                                            case "yellow":
                                                                layerTableRecord.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_Yellow);
                                                                break;
                                                            case "green":
                                                                layerTableRecord.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_Green);
                                                                break;
                                                            case "cyan":
                                                                layerTableRecord.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_Cyan);
                                                                break;
                                                            case "blue":
                                                                layerTableRecord.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_Blue);
                                                                break;
                                                            case "magenta":
                                                                layerTableRecord.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_Magenta);
                                                                break;
                                                            case "white":
                                                                layerTableRecord.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_White_Black);
                                                                break;
                                                            case "bylayer":
                                                                layerTableRecord.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_ByLayer);
                                                                break;
                                                            default:
                                                                short shortColor = Convert.ToInt16(color_string);
                                                                layerTableRecord.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, shortColor);
                                                                break;
                                                        }
                                                    }
                                                    #endregion
                                                    break;

                                                case "Linetype":
                                                    if (linetypeTable.Has(desiredSetting))
                                                    {
                                                        layerTableRecord.LinetypeObjectId = linetypeTable[desiredSetting];
                                                    }
                                                    else
                                                    {
                                                        MessageBox.Show($"{path} does not contain the Linetype '{desiredSetting}'. Please open the drawing and load in the linetype manually.");
                                                    }
                                                    break;

                                                case "Lineweight":
                                                    #region Lineweight Crazyness
                                                    switch (desiredSetting)
                                                    {
                                                        case "ByLineWeightDefault":
                                                            layerTableRecord.LineWeight = LineWeight.ByLineWeightDefault;
                                                            break;
                                                        case "ByBlock":
                                                            layerTableRecord.LineWeight = LineWeight.ByBlock;
                                                            break;
                                                        case "ByLayer":
                                                            layerTableRecord.LineWeight = LineWeight.ByLayer;
                                                            break;
                                                        case "LineWeight000":
                                                            layerTableRecord.LineWeight = LineWeight.LineWeight000;
                                                            break;
                                                        case "LineWeight005":
                                                            layerTableRecord.LineWeight = LineWeight.LineWeight005;
                                                            break;
                                                        case "LineWeight009":
                                                            layerTableRecord.LineWeight = LineWeight.LineWeight009;
                                                            break;
                                                        case "LineWeight013":
                                                            layerTableRecord.LineWeight = LineWeight.LineWeight013;
                                                            break;
                                                        case "LineWeight015":
                                                            layerTableRecord.LineWeight = LineWeight.LineWeight015;
                                                            break;
                                                        case "LineWeight018":
                                                            layerTableRecord.LineWeight = LineWeight.LineWeight018;
                                                            break;
                                                        case "LineWeight020":
                                                            layerTableRecord.LineWeight = LineWeight.LineWeight020;
                                                            break;
                                                        case "LineWeight025":
                                                            layerTableRecord.LineWeight = LineWeight.LineWeight025;
                                                            break;
                                                        case "LineWeight030":
                                                            layerTableRecord.LineWeight = LineWeight.LineWeight030;
                                                            break;
                                                        case "LineWeight035":
                                                            layerTableRecord.LineWeight = LineWeight.LineWeight035;
                                                            break;
                                                        case "LineWeight040":
                                                            layerTableRecord.LineWeight = LineWeight.LineWeight040;
                                                            break;
                                                        case "LineWeight050":
                                                            layerTableRecord.LineWeight = LineWeight.LineWeight050;
                                                            break;
                                                        case "LineWeight053":
                                                            layerTableRecord.LineWeight = LineWeight.LineWeight053;
                                                            break;
                                                        case "LineWeight060":
                                                            layerTableRecord.LineWeight = LineWeight.LineWeight060;
                                                            break;
                                                        case "LineWeight070":
                                                            layerTableRecord.LineWeight = LineWeight.LineWeight070;
                                                            break;
                                                        case "LineWeight080":
                                                            layerTableRecord.LineWeight = LineWeight.LineWeight080;
                                                            break;
                                                        case "LineWeight090":
                                                            layerTableRecord.LineWeight = LineWeight.LineWeight090;
                                                            break;
                                                        case "LineWeight100":
                                                            layerTableRecord.LineWeight = LineWeight.LineWeight100;
                                                            break;
                                                        case "LineWeight106":
                                                            layerTableRecord.LineWeight = LineWeight.LineWeight106;
                                                            break;
                                                        case "LineWeight120":
                                                            layerTableRecord.LineWeight = LineWeight.LineWeight120;
                                                            break;
                                                        case "LineWeight140":
                                                            layerTableRecord.LineWeight = LineWeight.LineWeight140;
                                                            break;
                                                        case "LineWeight158":
                                                            layerTableRecord.LineWeight = LineWeight.LineWeight158;
                                                            break;
                                                        case "LineWeight200":
                                                            layerTableRecord.LineWeight = LineWeight.LineWeight200;
                                                            break;
                                                        case "LineWeight211":
                                                            layerTableRecord.LineWeight = LineWeight.LineWeight211;
                                                            break;
                                                        default:
                                                            layerTableRecord.LineWeight = LineWeight.ByLayer;
                                                            MessageBox.Show($"The lineweight specified from {layerName} on the source drawing is not present in {path}. Lineweight set to ByLayer. Please reslove manually.");
                                                            break;
                                                    }
                                                    #endregion
                                                    break;

                                                case "Transparency":
                                                    int intSetting = Int32.Parse(desiredSetting);
                                                    byte alpha = (byte)(255 * (100 - intSetting) / 100);
                                                    Transparency transparency = new Transparency(alpha);
                                                    layerTableRecord.Transparency = transparency;
                                                    break;

                                                case "Plot":
                                                    if (desiredSetting == "Yes")
                                                    {
                                                        layerTableRecord.IsPlottable = true;
                                                    }
                                                    else
                                                    {
                                                        layerTableRecord.IsPlottable = false;
                                                    }
                                                    break;
                                                default:
                                                    MessageBox.Show("No value was found to be changed.");
                                                    break;
                                            }
                                            layerTableRecord.IsReconciled = true;
                                        }
                                    }
                                    trans2.Commit();
                                    database2.SaveAs(path, true, DwgVersion.Current, thisDrawing.Database.SecurityParameters);
                                }
                                HostApplicationServices.WorkingDatabase = thisDrawing.Database;
                            }

                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }
                }
            }
        }
        public void FixViewportLayerConflict(string path, string layerName, string property, string desiredSetting)
        {
            ObjectIdCollection viewportCollection = new ObjectIdCollection();
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor editor = thisDrawing.Editor;
            string shortLayerName = "";

            using (DocumentLock lock1 = thisDrawing.LockDocument())
            {
                using (Transaction trans1 = thisDrawing.TransactionManager.StartTransaction())
                {
                    if (File.Exists(path) && !path.Contains(MessagesAndNotifications.SourceDrawingNotSet) && path != null)
                    {
                        try
                        {
                            using (Database database2 = new Database(false, true))
                            {
                                database2.ReadDwgFile(path, FileOpenMode.OpenForReadAndWriteNoShare, true, "");
                                database2.CloseInput(true);
                                //database2.ResolveXrefs(useThreadEngine: true, doNewOnly: false);

                                HostApplicationServices.WorkingDatabase = database2;

                                using (Transaction trans2 = database2.TransactionManager.StartTransaction())
                                {
                                    DrawingSettingHelpers drawingSettingHelpers = new DrawingSettingHelpers();
                                    if (drawingSettingHelpers.IsVisretainOn() == false)
                                    {
                                        MessageBox.Show($"VISRETAIN is not enabled on {path}. Changes to it's xref layers will not be saved.");
                                    }
                                    LayerTable layerTable2 = trans2.GetObject(database2.LayerTableId, OpenMode.ForWrite) as LayerTable;

                                    //access the Layout DBDictionary of database 2 to build a collection of viewports
                                    DBDictionary dBLayoutDictionary = trans2.GetObject(database2.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                                    foreach (DBDictionaryEntry entry in dBLayoutDictionary)
                                    {
                                        Layout layout = trans2.GetObject(entry.Value, OpenMode.ForRead) as Layout;
                                        if (entry.Key != "Model")
                                        {
                                            viewportCollection = layout.GetViewports();
                                            viewportCollection.RemoveAt((0));
                                        }
                                    }

                                    foreach (ObjectId viewportID in viewportCollection)
                                    {
                                        Viewport viewport = trans2.GetObject(viewportID, OpenMode.ForWrite, false, true) as Viewport;

                                        foreach (ObjectId id2 in layerTable2)
                                        {
                                            //Set up a LayerTableRecord and a LinetypeTableRecord for each id2
                                            LayerTableRecord vpLayer = trans2.GetObject(id2, OpenMode.ForWrite) as LayerTableRecord;
                                            LinetypeTable linetypeTable = trans2.GetObject(database2.LinetypeTableId, OpenMode.ForRead) as LinetypeTable;
                                            LinetypeTableRecord vpLineTypeRecord = trans2.GetObject(vpLayer.LinetypeObjectId, OpenMode.ForWrite) as LinetypeTableRecord;

                                            if (vpLayer != null && vpLineTypeRecord != null)
                                            {
                                                //Get the viewport layer overrides
                                                LayerViewportProperties layerViewportProperties = vpLayer.GetViewportOverrides(viewport.ObjectId);
                                                //Get the name of the linetype, based on ObjectId. This will be used later to set the value on the layer model.
                                                LinetypeTableRecord linetypeTableRecord = trans2.GetObject(layerViewportProperties.LinetypeObjectId, OpenMode.ForRead) as LinetypeTableRecord;
                                                string viewportLinetypeName = linetypeTableRecord.Name;

                                                ObjectIdCollection objectIdCollection = new ObjectIdCollection();
                                                objectIdCollection.Add(vpLayer.ObjectId);

                                                if (vpLayer != null && vpLineTypeRecord != null && vpLayer.Name == layerName)
                                                {
                                                    switch (property)
                                                    {
                                                        case "FreezeThaw":
                                                            if (desiredSetting == "Frozen")
                                                            {
                                                                viewport.FreezeLayersInViewport(objectIdCollection.GetEnumerator());
                                                            }
                                                            else
                                                            {
                                                                viewport.ThawLayersInViewport(objectIdCollection.GetEnumerator());
                                                            }
                                                            break;

                                                        case "Color":
                                                            #region Color Crazyness
                                                            Autodesk.AutoCAD.Colors.Color color1 = new Autodesk.AutoCAD.Colors.Color();
                                                            string color_string = desiredSetting.ToLower();

                                                            if (color_string.ToLower().Contains(",") == true)
                                                            {
                                                                int idx1 = color_string.IndexOf(",", 0);
                                                                byte R = Convert.ToByte(color_string.Substring(0, idx1));
                                                                int idx2 = color_string.IndexOf(",", idx1 + 1);
                                                                byte G = Convert.ToByte(color_string.Substring(idx1 + 1, idx2 - idx1 - 1));
                                                                byte B = Convert.ToByte(color_string.Substring(idx2 + 1, color_string.Length - idx2 - 1));
                                                                color1 = Autodesk.AutoCAD.Colors.Color.FromRgb(R, G, B);

                                                                //Return the truecolor value.
                                                                layerViewportProperties.Color = color1;
                                                            }
                                                            else
                                                            {
                                                                switch (color_string)
                                                                {
                                                                    case "byblock":
                                                                        layerViewportProperties.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_ByBlock);
                                                                        break;
                                                                    case "red":
                                                                        layerViewportProperties.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_Red);
                                                                        break;
                                                                    case "yellow":
                                                                        layerViewportProperties.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_Yellow);
                                                                        break;
                                                                    case "green":
                                                                        layerViewportProperties.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_Green);
                                                                        break;
                                                                    case "cyan":
                                                                        layerViewportProperties.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_Cyan);
                                                                        break;
                                                                    case "blue":
                                                                        layerViewportProperties.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_Blue);
                                                                        break;
                                                                    case "magenta":
                                                                        layerViewportProperties.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_Magenta);
                                                                        break;
                                                                    case "white":
                                                                        layerViewportProperties.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_White_Black);
                                                                        break;
                                                                    case "bylayer":
                                                                        layerViewportProperties.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_ByLayer);
                                                                        break;
                                                                    default:
                                                                        short shortColor = Convert.ToInt16(color_string);
                                                                        layerViewportProperties.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, shortColor);
                                                                        break;
                                                                }
                                                            }
                                                            #endregion
                                                            break;

                                                        case "Linetype":
                                                            if (linetypeTable.Has(desiredSetting))
                                                            {
                                                                layerViewportProperties.LinetypeObjectId = linetypeTable[desiredSetting];
                                                            }
                                                            else
                                                            {
                                                                MessageBox.Show($"{path} does not contain the Linetype '{desiredSetting}'. Please open the drawing and load in the linetype manually.");
                                                            }
                                                            break;

                                                        case "Lineweight":
                                                            #region Lineweight Crazyness
                                                            switch (desiredSetting)
                                                            {
                                                                case "ByLineWeightDefault":
                                                                    layerViewportProperties.LineWeight = LineWeight.ByLineWeightDefault;
                                                                    break;
                                                                case "ByBlock":
                                                                    layerViewportProperties.LineWeight = LineWeight.ByBlock;
                                                                    break;
                                                                case "ByLayer":
                                                                    layerViewportProperties.LineWeight = LineWeight.ByLayer;
                                                                    break;
                                                                case "LineWeight000":
                                                                    layerViewportProperties.LineWeight = LineWeight.LineWeight000;
                                                                    break;
                                                                case "LineWeight005":
                                                                    layerViewportProperties.LineWeight = LineWeight.LineWeight005;
                                                                    break;
                                                                case "LineWeight009":
                                                                    layerViewportProperties.LineWeight = LineWeight.LineWeight009;
                                                                    break;
                                                                case "LineWeight013":
                                                                    layerViewportProperties.LineWeight = LineWeight.LineWeight013;
                                                                    break;
                                                                case "LineWeight015":
                                                                    layerViewportProperties.LineWeight = LineWeight.LineWeight015;
                                                                    break;
                                                                case "LineWeight018":
                                                                    layerViewportProperties.LineWeight = LineWeight.LineWeight018;
                                                                    break;
                                                                case "LineWeight020":
                                                                    layerViewportProperties.LineWeight = LineWeight.LineWeight020;
                                                                    break;
                                                                case "LineWeight025":
                                                                    layerViewportProperties.LineWeight = LineWeight.LineWeight025;
                                                                    break;
                                                                case "LineWeight030":
                                                                    layerViewportProperties.LineWeight = LineWeight.LineWeight030;
                                                                    break;
                                                                case "LineWeight035":
                                                                    layerViewportProperties.LineWeight = LineWeight.LineWeight035;
                                                                    break;
                                                                case "LineWeight040":
                                                                    layerViewportProperties.LineWeight = LineWeight.LineWeight040;
                                                                    break;
                                                                case "LineWeight050":
                                                                    layerViewportProperties.LineWeight = LineWeight.LineWeight050;
                                                                    break;
                                                                case "LineWeight053":
                                                                    layerViewportProperties.LineWeight = LineWeight.LineWeight053;
                                                                    break;
                                                                case "LineWeight060":
                                                                    layerViewportProperties.LineWeight = LineWeight.LineWeight060;
                                                                    break;
                                                                case "LineWeight070":
                                                                    layerViewportProperties.LineWeight = LineWeight.LineWeight070;
                                                                    break;
                                                                case "LineWeight080":
                                                                    layerViewportProperties.LineWeight = LineWeight.LineWeight080;
                                                                    break;
                                                                case "LineWeight090":
                                                                    layerViewportProperties.LineWeight = LineWeight.LineWeight090;
                                                                    break;
                                                                case "LineWeight100":
                                                                    layerViewportProperties.LineWeight = LineWeight.LineWeight100;
                                                                    break;
                                                                case "LineWeight106":
                                                                    layerViewportProperties.LineWeight = LineWeight.LineWeight106;
                                                                    break;
                                                                case "LineWeight120":
                                                                    layerViewportProperties.LineWeight = LineWeight.LineWeight120;
                                                                    break;
                                                                case "LineWeight140":
                                                                    layerViewportProperties.LineWeight = LineWeight.LineWeight140;
                                                                    break;
                                                                case "LineWeight158":
                                                                    layerViewportProperties.LineWeight = LineWeight.LineWeight158;
                                                                    break;
                                                                case "LineWeight200":
                                                                    layerViewportProperties.LineWeight = LineWeight.LineWeight200;
                                                                    break;
                                                                case "LineWeight211":
                                                                    layerViewportProperties.LineWeight = LineWeight.LineWeight211;
                                                                    break;
                                                                default:
                                                                    layerViewportProperties.LineWeight = LineWeight.ByLayer;
                                                                    MessageBox.Show($"The lineweight specified from {layerName} on the source drawing is not present in {path}. Lineweight set to ByLayer. Please reslove manually.");
                                                                    break;
                                                            }
                                                            #endregion
                                                            break;

                                                        case "Transparency":
                                                            int intSetting = Int32.Parse(desiredSetting);
                                                            byte alpha = (byte)(255 * (100 - intSetting) / 100);
                                                            Transparency transparency = new Transparency(alpha);
                                                            layerViewportProperties.Transparency = transparency;
                                                            break;

                                                        default:
                                                            MessageBox.Show("No value was found to be changed.");
                                                            break;
                                                    }
                                                    vpLayer.IsReconciled = true;
                                                }
                                            }
                                        }
                                    }
                                    trans2.Commit();
                                    database2.SaveAs(path, true, DwgVersion.Current, thisDrawing.Database.SecurityParameters);
                                }
                                HostApplicationServices.WorkingDatabase = thisDrawing.Database;
                            }

                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }
                }
            }
        }
        #endregion

        #region Fix Group Conflicts
        public void FixGroupOfLayerConflicts(string path, string property, IEnumerable<LayerComparisonLayerModel> conflicts)
        {
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor editor = thisDrawing.Editor;
            string shortLayerName = "";
            string shortConflictName = "";

            using (DocumentLock lock1 = thisDrawing.LockDocument())
            {
                using (Transaction trans1 = thisDrawing.TransactionManager.StartTransaction())
                {
                    if (File.Exists(path) && !path.Contains(MessagesAndNotifications.SourceDrawingNotSet) && path != null)
                    {
                        try
                        {
                            using (Database database2 = new Database(false, true))
                            {
                                database2.ReadDwgFile(path, FileOpenMode.OpenForReadAndWriteNoShare, true, "");
                                database2.CloseInput(true);
                                //database2.ResolveXrefs(useThreadEngine: true, doNewOnly: false);

                                HostApplicationServices.WorkingDatabase = database2;

                                using (Transaction trans2 = database2.TransactionManager.StartTransaction())
                                {
                                    DrawingSettingHelpers drawingSettingHelpers = new DrawingSettingHelpers();
                                    if (drawingSettingHelpers.IsVisretainOn() == false)
                                    {
                                        MessageBox.Show($"VISRETAIN is not enabled on {path}. Changes to it's xref layers will not be retained.");
                                    }
                                    LayerTable layerTable2 = trans2.GetObject(database2.LayerTableId, OpenMode.ForWrite) as LayerTable;
                                    foreach (var conflict in conflicts)
                                    {
                                        foreach (ObjectId id2 in layerTable2)
                                        {
                                            LayerTableRecord layerTableRecord = trans2.GetObject(id2, OpenMode.ForWrite) as LayerTableRecord;
                                            LinetypeTable linetypeTable = trans2.GetObject(database2.LinetypeTableId, OpenMode.ForRead) as LinetypeTable;
                                            LinetypeTableRecord lineTypeRecord = trans2.GetObject(layerTableRecord.LinetypeObjectId, OpenMode.ForWrite) as LinetypeTableRecord;
                                            layerTableRecord.IsReconciled = true;
                                            if (layerTableRecord != null && lineTypeRecord != null && layerTableRecord.Name == conflict.Name)
                                            {
                                                switch (property)
                                                {
                                                    case "OnOff":
                                                        if (conflict.OnOff == "Off")
                                                        {
                                                            layerTableRecord.IsOff = true;
                                                        }
                                                        else
                                                        {
                                                            layerTableRecord.IsOff = false;
                                                        }
                                                        break;

                                                    case "FreezeThaw":
                                                        if (conflict.Freeze == "Frozen")
                                                        {
                                                            layerTableRecord.IsFrozen = true;
                                                        }
                                                        else
                                                        {
                                                            layerTableRecord.IsFrozen = false;
                                                        }
                                                        break;

                                                    case "Color":
                                                        #region Color Crazyness
                                                        Autodesk.AutoCAD.Colors.Color color1 = new Autodesk.AutoCAD.Colors.Color();
                                                        string color_string = conflict.Color.ToLower();

                                                        if (color_string.ToLower().Contains(",") == true)
                                                        {
                                                            int idx1 = color_string.IndexOf(",", 0);
                                                            byte R = Convert.ToByte(color_string.Substring(0, idx1));
                                                            int idx2 = color_string.IndexOf(",", idx1 + 1);
                                                            byte G = Convert.ToByte(color_string.Substring(idx1 + 1, idx2 - idx1 - 1));
                                                            byte B = Convert.ToByte(color_string.Substring(idx2 + 1, color_string.Length - idx2 - 1));
                                                            color1 = Autodesk.AutoCAD.Colors.Color.FromRgb(R, G, B);

                                                            //Return the truecolor value.
                                                            layerTableRecord.Color = color1;
                                                        }
                                                        else
                                                        {
                                                            switch (color_string)
                                                            {
                                                                case "byblock":
                                                                    layerTableRecord.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_ByBlock);
                                                                    break;
                                                                case "red":
                                                                    layerTableRecord.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_Red);
                                                                    break;
                                                                case "yellow":
                                                                    layerTableRecord.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_Yellow);
                                                                    break;
                                                                case "green":
                                                                    layerTableRecord.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_Green);
                                                                    break;
                                                                case "cyan":
                                                                    layerTableRecord.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_Cyan);
                                                                    break;
                                                                case "blue":
                                                                    layerTableRecord.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_Blue);
                                                                    break;
                                                                case "magenta":
                                                                    layerTableRecord.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_Magenta);
                                                                    break;
                                                                case "white":
                                                                    layerTableRecord.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_White_Black);
                                                                    break;
                                                                case "bylayer":
                                                                    layerTableRecord.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_ByLayer);
                                                                    break;
                                                                default:
                                                                    short shortColor = Convert.ToInt16(color_string);
                                                                    layerTableRecord.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, shortColor);
                                                                    break;
                                                            }
                                                        }
                                                        #endregion
                                                        break;

                                                    case "Linetype":
                                                        if (linetypeTable.Has(conflict.Linetype))
                                                        {
                                                            layerTableRecord.LinetypeObjectId = linetypeTable[conflict.Linetype];
                                                        }
                                                        else
                                                        {
                                                            MessageBox.Show($"{path} does not contain the Linetype '{conflict.Linetype}'. Please open the drawing and load in the linetype manually.");
                                                        }
                                                        break;

                                                    case "Lineweight":
                                                        #region Lineweight Crazyness
                                                        switch (conflict.Lineweight)
                                                        {
                                                            case "ByLineWeightDefault":
                                                                layerTableRecord.LineWeight = LineWeight.ByLineWeightDefault;
                                                                break;
                                                            case "ByBlock":
                                                                layerTableRecord.LineWeight = LineWeight.ByBlock;
                                                                break;
                                                            case "ByLayer":
                                                                layerTableRecord.LineWeight = LineWeight.ByLayer;
                                                                break;
                                                            case "LineWeight000":
                                                                layerTableRecord.LineWeight = LineWeight.LineWeight000;
                                                                break;
                                                            case "LineWeight005":
                                                                layerTableRecord.LineWeight = LineWeight.LineWeight005;
                                                                break;
                                                            case "LineWeight009":
                                                                layerTableRecord.LineWeight = LineWeight.LineWeight009;
                                                                break;
                                                            case "LineWeight013":
                                                                layerTableRecord.LineWeight = LineWeight.LineWeight013;
                                                                break;
                                                            case "LineWeight015":
                                                                layerTableRecord.LineWeight = LineWeight.LineWeight015;
                                                                break;
                                                            case "LineWeight018":
                                                                layerTableRecord.LineWeight = LineWeight.LineWeight018;
                                                                break;
                                                            case "LineWeight020":
                                                                layerTableRecord.LineWeight = LineWeight.LineWeight020;
                                                                break;
                                                            case "LineWeight025":
                                                                layerTableRecord.LineWeight = LineWeight.LineWeight025;
                                                                break;
                                                            case "LineWeight030":
                                                                layerTableRecord.LineWeight = LineWeight.LineWeight030;
                                                                break;
                                                            case "LineWeight035":
                                                                layerTableRecord.LineWeight = LineWeight.LineWeight035;
                                                                break;
                                                            case "LineWeight040":
                                                                layerTableRecord.LineWeight = LineWeight.LineWeight040;
                                                                break;
                                                            case "LineWeight050":
                                                                layerTableRecord.LineWeight = LineWeight.LineWeight050;
                                                                break;
                                                            case "LineWeight053":
                                                                layerTableRecord.LineWeight = LineWeight.LineWeight053;
                                                                break;
                                                            case "LineWeight060":
                                                                layerTableRecord.LineWeight = LineWeight.LineWeight060;
                                                                break;
                                                            case "LineWeight070":
                                                                layerTableRecord.LineWeight = LineWeight.LineWeight070;
                                                                break;
                                                            case "LineWeight080":
                                                                layerTableRecord.LineWeight = LineWeight.LineWeight080;
                                                                break;
                                                            case "LineWeight090":
                                                                layerTableRecord.LineWeight = LineWeight.LineWeight090;
                                                                break;
                                                            case "LineWeight100":
                                                                layerTableRecord.LineWeight = LineWeight.LineWeight100;
                                                                break;
                                                            case "LineWeight106":
                                                                layerTableRecord.LineWeight = LineWeight.LineWeight106;
                                                                break;
                                                            case "LineWeight120":
                                                                layerTableRecord.LineWeight = LineWeight.LineWeight120;
                                                                break;
                                                            case "LineWeight140":
                                                                layerTableRecord.LineWeight = LineWeight.LineWeight140;
                                                                break;
                                                            case "LineWeight158":
                                                                layerTableRecord.LineWeight = LineWeight.LineWeight158;
                                                                break;
                                                            case "LineWeight200":
                                                                layerTableRecord.LineWeight = LineWeight.LineWeight200;
                                                                break;
                                                            case "LineWeight211":
                                                                layerTableRecord.LineWeight = LineWeight.LineWeight211;
                                                                break;
                                                            default:
                                                                layerTableRecord.LineWeight = LineWeight.ByLayer;
                                                                MessageBox.Show($"{conflict.Lineweight} from {conflict.Name} on the source drawing is not present in {path}. The lineweight for {layerTableRecord.Name} has been set to ByLayer by default. Please resolve this issue manually.");
                                                                break;
                                                        }
                                                        #endregion
                                                        break;

                                                    case "Transparency":
                                                        int intSetting = Int32.Parse(conflict.Transparency);
                                                        byte alpha = (byte)(255 * (100 - intSetting) / 100);
                                                        Transparency transparency = new Transparency(alpha);
                                                        layerTableRecord.Transparency = transparency;
                                                        break;

                                                    case "Plot":
                                                        if (conflict.Plot == "Yes")
                                                        {
                                                            layerTableRecord.IsPlottable = true;
                                                        }
                                                        else
                                                        {
                                                            layerTableRecord.IsPlottable = false;
                                                        }
                                                        break;
                                                    default:
                                                        MessageBox.Show("No value was found to be changed.");
                                                        break;
                                                }
                                                layerTableRecord.IsReconciled = true;
                                            }
                                        }
                                    }
                                    trans2.Commit();
                                    database2.SaveAs(path, true, DwgVersion.Current, thisDrawing.Database.SecurityParameters);
                                }
                                HostApplicationServices.WorkingDatabase = thisDrawing.Database;
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }
                }
            }
        }
        public void FixGroupOfViewportLayerConflicts(string path, string property, IEnumerable<LayerComparisonViewportLayerModel> conflicts)
        {
            ObjectIdCollection viewportCollection = new ObjectIdCollection();
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor editor = thisDrawing.Editor;
            string shortLayerName = "";

            using (DocumentLock lock1 = thisDrawing.LockDocument())
            {
                using (Transaction trans1 = thisDrawing.TransactionManager.StartTransaction())
                {
                    if (File.Exists(path) && !path.Contains(MessagesAndNotifications.SourceDrawingNotSet) && path != null)
                    {
                        try
                        {
                            using (Database database2 = new Database(false, true))
                            {
                                database2.ReadDwgFile(path, FileOpenMode.OpenForReadAndWriteNoShare, true, "");
                                database2.CloseInput(true);
                                //database2.ResolveXrefs(useThreadEngine: true, doNewOnly: false);

                                HostApplicationServices.WorkingDatabase = database2;

                                using (Transaction trans2 = database2.TransactionManager.StartTransaction())
                                {
                                    DrawingSettingHelpers drawingSettingHelpers = new DrawingSettingHelpers();
                                    if (drawingSettingHelpers.IsVisretainOn() == false)
                                    {
                                        MessageBox.Show($"VISRETAIN is not enabled on {path}. Changes to it's xref layers will not be saved.");
                                    }
                                    LayerTable layerTable2 = trans2.GetObject(database2.LayerTableId, OpenMode.ForWrite) as LayerTable;

                                    //access the Layout DBDictionary of database 2 to build a collection of viewports
                                    DBDictionary dBLayoutDictionary = trans2.GetObject(database2.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                                    foreach (DBDictionaryEntry entry in dBLayoutDictionary)
                                    {
                                        Layout layout = trans2.GetObject(entry.Value, OpenMode.ForRead) as Layout;
                                        if (entry.Key != "Model")
                                        {
                                            viewportCollection = layout.GetViewports();
                                            viewportCollection.RemoveAt((0));
                                        }
                                    }

                                    foreach (var conflict in conflicts)
                                    {
                                        foreach (ObjectId viewportID in viewportCollection)
                                        {
                                            Viewport viewport = trans2.GetObject(viewportID, OpenMode.ForWrite, false, true) as Viewport;

                                            foreach (ObjectId id2 in layerTable2)
                                            {
                                                //Set up a LayerTableRecord and a LinetypeTableRecord for each id2
                                                LayerTableRecord vpLayer = trans2.GetObject(id2, OpenMode.ForWrite) as LayerTableRecord;
                                                LinetypeTable linetypeTable = trans2.GetObject(database2.LinetypeTableId, OpenMode.ForRead) as LinetypeTable;
                                                LinetypeTableRecord vpLineTypeRecord = trans2.GetObject(vpLayer.LinetypeObjectId, OpenMode.ForWrite) as LinetypeTableRecord;

                                                if (vpLayer != null && vpLineTypeRecord != null)
                                                {
                                                    //Get the viewport layer overrides
                                                    LayerViewportProperties layerViewportProperties = vpLayer.GetViewportOverrides(viewport.ObjectId);
                                                    //Get the name of the linetype, based on ObjectId. This will be used later to set the value on the layer model.
                                                    LinetypeTableRecord linetypeTableRecord = trans2.GetObject(layerViewportProperties.LinetypeObjectId, OpenMode.ForRead) as LinetypeTableRecord;
                                                    string viewportLinetypeName = linetypeTableRecord.Name;

                                                    ObjectIdCollection objectIdCollection = new ObjectIdCollection();
                                                    objectIdCollection.Add(vpLayer.ObjectId);
                                                    if (vpLayer != null && vpLineTypeRecord != null && vpLayer.Name == conflict.Name)
                                                    {
                                                        switch (property)
                                                        {
                                                            case "FreezeThaw":
                                                                if (conflict.ViewportFreeze == "Frozen")
                                                                {
                                                                    viewport.FreezeLayersInViewport(objectIdCollection.GetEnumerator());
                                                                }
                                                                else
                                                                {
                                                                    viewport.ThawLayersInViewport(objectIdCollection.GetEnumerator());
                                                                }
                                                                break;

                                                            case "Color":
                                                                #region Color Crazyness
                                                                Autodesk.AutoCAD.Colors.Color color1 = new Autodesk.AutoCAD.Colors.Color();
                                                                string color_string = conflict.ViewportColor.ToLower();

                                                                if (color_string.ToLower().Contains(",") == true)
                                                                {
                                                                    int idx1 = color_string.IndexOf(",", 0);
                                                                    byte R = Convert.ToByte(color_string.Substring(0, idx1));
                                                                    int idx2 = color_string.IndexOf(",", idx1 + 1);
                                                                    byte G = Convert.ToByte(color_string.Substring(idx1 + 1, idx2 - idx1 - 1));
                                                                    byte B = Convert.ToByte(color_string.Substring(idx2 + 1, color_string.Length - idx2 - 1));
                                                                    color1 = Autodesk.AutoCAD.Colors.Color.FromRgb(R, G, B);

                                                                    //Return the truecolor value.
                                                                    layerViewportProperties.Color = color1;
                                                                }
                                                                else
                                                                {
                                                                    switch (color_string)
                                                                    {
                                                                        case "byblock":
                                                                            layerViewportProperties.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_ByBlock);
                                                                            break;
                                                                        case "red":
                                                                            layerViewportProperties.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_Red);
                                                                            break;
                                                                        case "yellow":
                                                                            layerViewportProperties.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_Yellow);
                                                                            break;
                                                                        case "green":
                                                                            layerViewportProperties.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_Green);
                                                                            break;
                                                                        case "cyan":
                                                                            layerViewportProperties.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_Cyan);
                                                                            break;
                                                                        case "blue":
                                                                            layerViewportProperties.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_Blue);
                                                                            break;
                                                                        case "magenta":
                                                                            layerViewportProperties.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_Magenta);
                                                                            break;
                                                                        case "white":
                                                                            layerViewportProperties.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_White_Black);
                                                                            break;
                                                                        case "bylayer":
                                                                            layerViewportProperties.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_ByLayer);
                                                                            break;
                                                                        default:
                                                                            short shortColor = Convert.ToInt16(color_string);
                                                                            layerViewportProperties.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, shortColor);
                                                                            break;
                                                                    }
                                                                }
                                                                #endregion
                                                                break;

                                                            case "Linetype":
                                                                if (linetypeTable.Has(conflict.ViewportLinetype))
                                                                {
                                                                    layerViewportProperties.LinetypeObjectId = linetypeTable[conflict.ViewportLinetype];
                                                                }
                                                                else
                                                                {
                                                                    MessageBox.Show($"{path} does not contain the Linetype '{conflict.ViewportLinetype}'. Please open the drawing and load in the linetype manually.");
                                                                }
                                                                break;

                                                            case "Lineweight":
                                                                #region Lineweight Crazyness
                                                                switch (conflict.ViewportLineweight)
                                                                {
                                                                    case "ByLineWeightDefault":
                                                                        layerViewportProperties.LineWeight = LineWeight.ByLineWeightDefault;
                                                                        break;
                                                                    case "ByBlock":
                                                                        layerViewportProperties.LineWeight = LineWeight.ByBlock;
                                                                        break;
                                                                    case "ByLayer":
                                                                        layerViewportProperties.LineWeight = LineWeight.ByLayer;
                                                                        break;
                                                                    case "LineWeight000":
                                                                        layerViewportProperties.LineWeight = LineWeight.LineWeight000;
                                                                        break;
                                                                    case "LineWeight005":
                                                                        layerViewportProperties.LineWeight = LineWeight.LineWeight005;
                                                                        break;
                                                                    case "LineWeight009":
                                                                        layerViewportProperties.LineWeight = LineWeight.LineWeight009;
                                                                        break;
                                                                    case "LineWeight013":
                                                                        layerViewportProperties.LineWeight = LineWeight.LineWeight013;
                                                                        break;
                                                                    case "LineWeight015":
                                                                        layerViewportProperties.LineWeight = LineWeight.LineWeight015;
                                                                        break;
                                                                    case "LineWeight018":
                                                                        layerViewportProperties.LineWeight = LineWeight.LineWeight018;
                                                                        break;
                                                                    case "LineWeight020":
                                                                        layerViewportProperties.LineWeight = LineWeight.LineWeight020;
                                                                        break;
                                                                    case "LineWeight025":
                                                                        layerViewportProperties.LineWeight = LineWeight.LineWeight025;
                                                                        break;
                                                                    case "LineWeight030":
                                                                        layerViewportProperties.LineWeight = LineWeight.LineWeight030;
                                                                        break;
                                                                    case "LineWeight035":
                                                                        layerViewportProperties.LineWeight = LineWeight.LineWeight035;
                                                                        break;
                                                                    case "LineWeight040":
                                                                        layerViewportProperties.LineWeight = LineWeight.LineWeight040;
                                                                        break;
                                                                    case "LineWeight050":
                                                                        layerViewportProperties.LineWeight = LineWeight.LineWeight050;
                                                                        break;
                                                                    case "LineWeight053":
                                                                        layerViewportProperties.LineWeight = LineWeight.LineWeight053;
                                                                        break;
                                                                    case "LineWeight060":
                                                                        layerViewportProperties.LineWeight = LineWeight.LineWeight060;
                                                                        break;
                                                                    case "LineWeight070":
                                                                        layerViewportProperties.LineWeight = LineWeight.LineWeight070;
                                                                        break;
                                                                    case "LineWeight080":
                                                                        layerViewportProperties.LineWeight = LineWeight.LineWeight080;
                                                                        break;
                                                                    case "LineWeight090":
                                                                        layerViewportProperties.LineWeight = LineWeight.LineWeight090;
                                                                        break;
                                                                    case "LineWeight100":
                                                                        layerViewportProperties.LineWeight = LineWeight.LineWeight100;
                                                                        break;
                                                                    case "LineWeight106":
                                                                        layerViewportProperties.LineWeight = LineWeight.LineWeight106;
                                                                        break;
                                                                    case "LineWeight120":
                                                                        layerViewportProperties.LineWeight = LineWeight.LineWeight120;
                                                                        break;
                                                                    case "LineWeight140":
                                                                        layerViewportProperties.LineWeight = LineWeight.LineWeight140;
                                                                        break;
                                                                    case "LineWeight158":
                                                                        layerViewportProperties.LineWeight = LineWeight.LineWeight158;
                                                                        break;
                                                                    case "LineWeight200":
                                                                        layerViewportProperties.LineWeight = LineWeight.LineWeight200;
                                                                        break;
                                                                    case "LineWeight211":
                                                                        layerViewportProperties.LineWeight = LineWeight.LineWeight211;
                                                                        break;
                                                                    default:
                                                                        layerViewportProperties.LineWeight = LineWeight.ByLayer;
                                                                        MessageBox.Show($"{conflict.ViewportLineweight} from {conflict.Name} on the source drawing is not present in {path}. The lineweight for {vpLayer.Name} has been set to ByLayer by default. Please resolve this issue manually.");
                                                                        break;
                                                                }
                                                                #endregion
                                                                break;

                                                            case "Transparency":
                                                                int intSetting = Int32.Parse(conflict.ViewportTransparency);
                                                                byte alpha = (byte)(255 * (100 - intSetting) / 100);
                                                                Transparency transparency = new Transparency(alpha);
                                                                layerViewportProperties.Transparency = transparency;
                                                                break;

                                                            default:
                                                                MessageBox.Show("No value was found to be changed.");
                                                                break;
                                                        }
                                                        vpLayer.IsReconciled = true;

                                                    }
                                                }
                                            }
                                        }
                                    }
                                    trans2.Commit();
                                    database2.SaveAs(path, true, DwgVersion.Current, thisDrawing.Database.SecurityParameters);
                                }
                                HostApplicationServices.WorkingDatabase = thisDrawing.Database;
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }
                }
            }
        }
        #endregion

        #region Helpers
        public void EnableVisretain(string drawing)
        {
            try
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
                                    database2.ResolveXrefs(useThreadEngine: true, doNewOnly: false);

                                    HostApplicationServices.WorkingDatabase = database2;

                                    using (Transaction trans2 = database2.TransactionManager.StartTransaction())
                                    {
                                        if (System.Convert.ToInt32(Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("VISRETAIN")) != 1)
                                        {
                                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("VISRETAIN", 1);
                                        }
                                        trans2.Commit();
                                        database2.SaveAs(drawing, true, DwgVersion.Current, thisDrawing.Database.SecurityParameters);
                                    }
                                    HostApplicationServices.WorkingDatabase = thisDrawing.Database;
                                }
                            }
                            catch (System.Exception)
                            {
                                MessageBox.Show("The drawing could not be accessed.", $"Problem accessing database of {drawing}", MessageBoxButton.OK, MessageBoxImage.Error);
                                throw;
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public void ReconcileAllLayers(string drawing)
        {
            try
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
                                    database2.ResolveXrefs(useThreadEngine: true, doNewOnly: false);

                                    HostApplicationServices.WorkingDatabase = database2;

                                    using (Transaction trans2 = database2.TransactionManager.StartTransaction())
                                    {
                                        LayerTable layerTable2 = trans2.GetObject(database2.LayerTableId, OpenMode.ForWrite) as LayerTable;

                                        foreach (ObjectId id in layerTable2)
                                        {
                                            LayerTableRecord layerTableRecord = trans2.GetObject(id, OpenMode.ForWrite) as LayerTableRecord;
                                            layerTableRecord.IsReconciled = true;
                                        }
                                        trans2.Commit();
                                        database2.SaveAs(drawing, true, DwgVersion.Current, thisDrawing.Database.SecurityParameters);
                                    }
                                    HostApplicationServices.WorkingDatabase = thisDrawing.Database;
                                }
                            }
                            catch (System.Exception)
                            {
                                MessageBox.Show("The drawing could not be accessed.", $"Problem accessing database of {drawing}", MessageBoxButton.OK, MessageBoxImage.Error);
                                throw;
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region Unused
        /// <summary>
        /// Builds an <see cref="ObservableCollection{CombinedLayerModel}"/> of <see cref="CombinedLayerModel"/> that contains all regular layer and viewport layer data.
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public ObservableCollection<CombinedLayerModel> BuildCombinedLayersCollection(string s)
        {
            ObservableCollection<CombinedLayerModel> output = new ObservableCollection<CombinedLayerModel>();
            ObjectIdCollection viewportCollection = new ObjectIdCollection();

            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor editor = thisDrawing.Editor;
            using (DocumentLock lock1 = thisDrawing.LockDocument())
            {
                using (Transaction trans1 = thisDrawing.TransactionManager.StartTransaction())
                {
                    if (File.Exists(s) && s != null)
                    {
                        try
                        {
                            using (Database database2 = new Database(false, true))
                            {
                                database2.ReadDwgFile(s, FileOpenMode.OpenForReadAndWriteNoShare, true, "");
                                database2.CloseInput(true);
                                database2.ResolveXrefs(useThreadEngine: true, doNewOnly: false);

                                HostApplicationServices.WorkingDatabase = database2;

                                using (Transaction trans2 = database2.TransactionManager.StartTransaction())
                                {
                                    LayerTable layerTable2 = trans2.GetObject(database2.LayerTableId, OpenMode.ForRead) as LayerTable;

                                    #region Regular Layers
                                    foreach (ObjectId id in layerTable2)
                                    {
                                        LayerTableRecord layerTableRecord = trans2.GetObject(id, OpenMode.ForRead) as LayerTableRecord;
                                        LinetypeTableRecord lineTypeRecord = trans2.GetObject(layerTableRecord.LinetypeObjectId, OpenMode.ForRead) as LinetypeTableRecord;

                                        if (layerTableRecord != null && lineTypeRecord != null)
                                        {
                                            CombinedLayerModel layer = new CombinedLayerModel();

                                            layer.DrawingPath = s;
                                            layer.Name = layerTableRecord.Name;
                                            layer.OnOff = layerTableRecord.IsOff ? "Off" : "On";
                                            layer.Freeze = layerTableRecord.IsFrozen ? "Frozen" : "Thawed";
                                            #region Color
                                            //Get the layer color by creating an autocad color, setting it to the value of the color found on the layer, and converting it as necessary.
                                            Autodesk.AutoCAD.Colors.Color color1 = layerTableRecord.Color;
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
                                                layer.Color = color_string;
                                            }
                                            else
                                            {
                                                //Return the index color.
                                                layer.Color = color_string;
                                            }
                                            #endregion
                                            layer.Linetype = lineTypeRecord.Name;
                                            layer.Lineweight = layerTableRecord.LineWeight.ToString();
                                            layer.Plot = layerTableRecord.IsPlottable ? "Yes" : "No";
                                            #region Transparency
                                            if (layerTableRecord.Transparency.IsByAlpha)
                                            {
                                                int percentage = (int)(((255 - layerTableRecord.Transparency.Alpha) * 100) / 255);
                                                layer.Transparency = percentage.ToString();
                                            }
                                            else
                                            {
                                                layer.Transparency = "0";
                                            }
                                            output.Add(layer);
                                            #endregion

                                            output.Add(layer);
                                        }
                                    }
                                    #endregion

                                    #region Viewport Layers
                                    //access the Layout DBDictionary of database 2 to build a collection of viewports
                                    DBDictionary dBLayoutDictionary = trans2.GetObject(database2.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                                    foreach (DBDictionaryEntry entry in dBLayoutDictionary)
                                    {
                                        Layout layout = trans2.GetObject(entry.Value, OpenMode.ForRead) as Layout;
                                        if (entry.Key != "Model")
                                        {
                                            viewportCollection = layout.GetViewports();
                                            viewportCollection.RemoveAt((0));
                                        }
                                    }
                                    foreach (ObjectId viewportID in viewportCollection)
                                    {
                                        Viewport viewport = trans2.GetObject(viewportID, OpenMode.ForRead, false, true) as Viewport;

                                        foreach (ObjectId id2 in layerTable2)
                                        {
                                            //Set up a LayerTableRecord and a LinetypeTableRecord for each id2
                                            LayerTableRecord vpLayer = trans2.GetObject(id2, OpenMode.ForRead) as LayerTableRecord;
                                            LinetypeTableRecord vpLineTypeRecord = trans2.GetObject(vpLayer.LinetypeObjectId, OpenMode.ForRead) as LinetypeTableRecord;

                                            if (vpLayer != null && vpLineTypeRecord != null)
                                            {
                                                //Get the viewport layer overrides
                                                LayerViewportProperties layerViewportProperties = vpLayer.GetViewportOverrides(viewport.ObjectId);
                                                //Get the name of the linetype, based on ObjectId. This will be used later to set the value on the layer model.
                                                LinetypeTableRecord linetypeTableRecord = trans2.GetObject(layerViewportProperties.LinetypeObjectId, OpenMode.ForRead) as LinetypeTableRecord;
                                                string viewportLinetypeName = linetypeTableRecord.Name;

                                                //Create a model to contain the information
                                                CombinedLayerModel viewportLayerModel = new CombinedLayerModel();

                                                //Set the properties in the model
                                                viewportLayerModel.ViewportLayer = viewport.Layer;
                                                viewportLayerModel.DrawingPath = s;
                                                viewportLayerModel.Name = vpLayer.Name;
                                                viewportLayerModel.ViewportPosition = viewport.CenterPoint;
                                                viewportLayerModel.ViewportFreeze = viewport.IsLayerFrozenInViewport(vpLayer.ObjectId) ? "Frozen" : "Thawed";
                                                viewportLayerModel.ViewportLinetype = viewportLinetypeName;
                                                viewportLayerModel.ViewportLineweight = layerViewportProperties.LineWeight.ToString();
                                                #region Color
                                                //Get the layer color by creating an autocad color, setting it to the value of the color found on the layer, and converting it as necessary.
                                                Autodesk.AutoCAD.Colors.Color color1 = layerViewportProperties.Color;
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
                                                    viewportLayerModel.ViewportColor = color_string;
                                                }
                                                else
                                                {
                                                    //Return the index color.
                                                    viewportLayerModel.ViewportColor = color_string;
                                                }
                                                #endregion
                                                #region Transparency
                                                if (layerViewportProperties.Transparency.IsByAlpha)
                                                {
                                                    int percentage = (int)(((255 - layerViewportProperties.Transparency.Alpha) * 100) / 255);
                                                    viewportLayerModel.ViewportTransparency = percentage.ToString();
                                                }
                                                else
                                                {
                                                    viewportLayerModel.ViewportTransparency = "0";
                                                }
                                                #endregion

                                                output.Add(viewportLayerModel);
                                            }
                                        }
                                    }
                                    #endregion
                                }
                                HostApplicationServices.WorkingDatabase = thisDrawing.Database;
                                return output;
                            }
                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }
                    return output;
                }
            }
        }
        #endregion
    }
}
