using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System.Windows;

namespace CADTechnologiesSource.All.Commands
{
    public class AddCommonAnnoScales
    {
        [CommandMethod("ADDCOMMONANNOSCALES")]
        public void AddCommonAnnotativeScalesToCurrentDWG()
        {
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database thisDatabase = thisDrawing.Database;
            Editor thisEditor = thisDrawing.Editor;
            try
            {
                //object context manager databse object contains object context related information
                ObjectContextManager objectContextManager = thisDatabase.ObjectContextManager;
                if (objectContextManager != null)
                {
                    //this gets the context collection called ACDB_ANNOTATIONSCALES which contains the scales
                    ObjectContextCollection objectContexts = objectContextManager.GetContextCollection("ACDB_ANNOTATIONSCALES");

                    //Remove all of the BS scales that come with drawings
                    foreach (AnnotationScale scale in objectContexts)
                    {
                        if (scale != null)
                        {
                            if (scale.Name != "1:1")
                            {
                                objectContexts.RemoveContext(scale.Name);
                            }
                        }
                    }

                    //Create scales in multiples of 20, up to 200.
                    if (objectContexts != null)
                        for (int i = 20; i <= 200; i += 20)
                        {
                            AnnotationScale annotationScale = new AnnotationScale();
                            annotationScale.Name = $"1\" = {i}\'";
                            annotationScale.PaperUnits = 1;
                            annotationScale.DrawingUnits = i;

                            //add every scale up to (and including) 100
                            if (i <= 100)
                            {
                                objectContexts.AddContext(annotationScale);
                            }

                            //Don't add 120, 140, 160, or 180
                            if (i > 100 && i / 200 == 1)
                            {
                                objectContexts.AddContext(annotationScale);
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
}
