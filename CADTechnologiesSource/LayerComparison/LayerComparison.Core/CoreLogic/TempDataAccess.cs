using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using LayerComparison.Core.Models;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows;

namespace LayerComparison.Core.CoreLogic
{
    public class TempDataAccess
    {
        public ObservableCollection<LayerItemViewModel> BuildLayerItem(string s)
        {
            ObservableCollection<LayerItemViewModel> layerItemViewModels = new ObservableCollection<LayerItemViewModel>();

            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database thisDatabase = thisDrawing.Database;
            using (DocumentLock documentLock = thisDrawing.LockDocument())
            {
                using (Transaction transaction = thisDrawing.TransactionManager.StartTransaction())
                {
                    if (File.Exists(s) && s != null)
                    {
                        try
                        {
                            using (Database sideDatabase = new Database(false, true))
                            {
                                sideDatabase.ReadDwgFile(s, FileOpenMode.OpenForReadAndWriteNoShare, true, "");
                                sideDatabase.CloseInput(true);
                                sideDatabase.ResolveXrefs(useThreadEngine: true, doNewOnly: false);

                                HostApplicationServices.WorkingDatabase = sideDatabase;

                                using (Transaction sideTransaction = sideDatabase.TransactionManager.StartTransaction())
                                {
                                    LayerTable layerTable = sideTransaction.GetObject(sideDatabase.LayerTableId, OpenMode.ForRead) as LayerTable;
                                    foreach (ObjectId objectId in layerTable)
                                    {
                                        LayerTableRecord layerTableRecord = sideTransaction.GetObject(objectId, OpenMode.ForRead) as LayerTableRecord;
                                        LinetypeTableRecord lineTypeRecord = sideTransaction.GetObject(layerTableRecord.LinetypeObjectId, OpenMode.ForRead) as LinetypeTableRecord;

                                        if (layerTableRecord != null && lineTypeRecord != null)
                                        {
                                            LayerItemViewModel layerItemViewModel = new LayerItemViewModel();

                                            layerItemViewModel.Drawing = s;
                                            layerItemViewModel.Name = layerTableRecord.Name;
                                            layerItemViewModel.On = layerTableRecord.IsOff ? false : true;
                                            layerItemViewModel.Freeze = layerTableRecord.IsFrozen ? true : false;
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
                                                layerItemViewModel.Color = color_string;
                                            }
                                            else
                                            {
                                                //Return the index color.
                                                layerItemViewModel.Color = color_string;
                                            }
                                            #endregion
                                            layerItemViewModel.Linetype = lineTypeRecord.Name;
                                            layerItemViewModel.Lineweight = layerTableRecord.LineWeight.ToString();
                                            layerItemViewModel.Plot = layerTableRecord.IsPlottable ? true : false;
                                            #region Transparency
                                            if (layerTableRecord.Transparency.IsByAlpha)
                                            {
                                                int percentage = (int)(((255 - layerTableRecord.Transparency.Alpha) * 100) / 255);
                                                layerItemViewModel.Transparency = percentage.ToString();
                                            }
                                            else
                                            {
                                                layerItemViewModel.Transparency = "0";
                                            }
                                            #endregion

                                            layerItemViewModels.Add(layerItemViewModel);
                                        }
                                    }
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

            return layerItemViewModels;
        }
    }
}
