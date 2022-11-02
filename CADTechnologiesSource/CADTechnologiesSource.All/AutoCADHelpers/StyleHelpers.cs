using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using CADTechnologiesSource.All.Models;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace CADTechnologiesSource.All.AutoCADHelpers
{
    public class AutoCADStyleHelpers
    {
        #region Text Styles
        /// <summary>
        /// Returns a <see cref="List{string}"/> of the text styles in use in your current drawing.
        /// </summary>
        public List<string> ReturnTextStylesFromCurrentDrawing()
        {
            List<string> textStyles = new List<string>();

            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = thisDrawing.LockDocument())
            {
                using (Transaction trans1 = thisDrawing.TransactionManager.StartTransaction())
                {
                    TextStyleTable textStyleTable = trans1.GetObject(thisDrawing.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                    foreach (ObjectId iD in textStyleTable)
                    {
                        TextStyleTableRecord textStyle = trans1.GetObject(iD, OpenMode.ForRead) as TextStyleTableRecord;
                        if (textStyle != null)
                        {
                            string stylename = textStyle.Name;
                            if (string.IsNullOrEmpty(stylename) == false)
                            {
                                //Make sure the text style isn't coming from an XREF.
                                if (stylename.Contains("|") == false)
                                {
                                    textStyles.Add(stylename);
                                }
                            }
                        }
                    }
                    return textStyles;
                }
            }
        }

        public ObservableCollection<TextStyleModel> ReturnTextStyleModelsFromCurrentDrawing()
        {
            ObservableCollection<TextStyleModel> textStyleModels = new ObservableCollection<TextStyleModel>();

            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = thisDrawing.LockDocument())
            {
                using (Transaction trans1 = thisDrawing.TransactionManager.StartTransaction())
                {
                    TextStyleTable textStyleTable = trans1.GetObject(thisDrawing.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                    foreach (ObjectId iD in textStyleTable)
                    {
                        TextStyleTableRecord textStyle = trans1.GetObject(iD, OpenMode.ForRead) as TextStyleTableRecord;
                        if (textStyle != null)
                        {
                            TextStyleModel textStyleModel = new TextStyleModel();
                            if (string.IsNullOrEmpty(textStyle.Name) == false)
                            {
                                //Make sure the text style isn't coming from an XREF.
                                if (textStyle.Name.Contains("|") == false)
                                {
                                    textStyleModel.Name = textStyle.Name;
                                    textStyleModel.Font = textStyle.Font;
                                    textStyleModel.TextSize = textStyle.TextSize;
                                    textStyleModel.ObliquingAngle = textStyle.ObliquingAngle;
                                    textStyleModel.PaperOrientation = textStyle.PaperOrientation;
                                    textStyleModel.IsVertical = textStyleModel.IsVertical;

                                    textStyleModels.Add(textStyleModel);
                                }
                            }
                        }
                    }
                    return textStyleModels;
                }
            }
        }
        #endregion


        #region MLeader Styles


        #endregion


        #region Dimension Styles


        #endregion
    }
}
