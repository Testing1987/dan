using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows;

namespace SystemVariableSetter.AutoCADHelpers
{
    public class DatabaseHelpers
    {
        #region File Access
        /// <summary>
        /// Attempts to open a .dwg file in AutoCAD.
        /// </summary>
        /// <param name="s">The path of the .dwg file you wish to open.</param>
        public void OpenDrawing(string s)
        {
            DocumentCollection documentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
            if (File.Exists(s) && s != null)
            {
                documentCollection.Open(s, false);
            }
            else
            {
                MessageBox.Show($"{s} could not be found. Please verify that the file exists and try again.");
                return;
            }
        }

        /// <summary>
        /// Attempts to open a .dwg file in AutoCAD.
        /// </summary>
        /// <param name="s">The path of the .dwg file you wish to open.</param>
        public void OpenDrawingAndZoomTo(string s, Handle handle, string objectLayout)
        {
            bool isOpen = false;
            //Get the document collection to check to see if the drawing is open
            DocumentCollection documentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
            foreach (Document document in documentCollection)
            {
                //Document is found inside the collection so activate it and zoom to the object
                if (document.Name == s)
                {
                    isOpen = true;
                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument = document;

                    //Zoom To
                    Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    Database thisDatabase = thisDrawing.Database;
                    Editor editor = thisDrawing.Editor;
                    using (Transaction transaction1 = thisDatabase.TransactionManager.StartTransaction())
                    {
                        try
                        {
                            //Check to see if our object's layout is in model space....
                            if (objectLayout == "Model")
                            {
                                //if our current layout is not Model
                                if (objectLayout != LayoutManager.Current.CurrentLayout)
                                {
                                    //Switch to Model
                                    LayoutManager.Current.CurrentLayout = objectLayout;
                                }
                            }

                            //Check to see if our object's layout is in a paper space layout
                            if (objectLayout != "Model")
                            {
                                //Check to make sure it's not our current layout and...
                                if (objectLayout != LayoutManager.Current.CurrentLayout)
                                {
                                    //switch to it if it isn't
                                    LayoutManager.Current.CurrentLayout = objectLayout;
                                }
                            }
                            using (ViewTableRecord view = editor.GetCurrentView())
                            {
                                Matrix3d WCS2DCS =
                                    (Matrix3d.Rotation(-view.ViewTwist, view.ViewDirection, view.Target) *
                                    Matrix3d.Displacement(view.Target - Point3d.Origin) *
                                    Matrix3d.PlaneToWorld(view.ViewDirection))
                                    .Inverse();
                                ObjectId objectId = thisDatabase.GetObjectId(false, handle, 0);
                                Entity entity = (Entity)transaction1.GetObject(objectId, OpenMode.ForRead);
                                Extents3d extents3D = entity.GeometricExtents;
                                extents3D.AddExtents(entity.GeometricExtents);
                                extents3D.TransformBy(WCS2DCS);
                                view.Width = extents3D.MaxPoint.X - extents3D.MinPoint.X;
                                view.Height = extents3D.MaxPoint.Y - extents3D.MinPoint.Y;
                                view.CenterPoint =
                                    new Point2d((extents3D.MaxPoint.X + extents3D.MinPoint.X) / 2.0, (extents3D.MaxPoint.Y + extents3D.MinPoint.Y) / 2.0);
                                editor.SetCurrentView(view);
                                transaction1.Commit();
                            }
                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }
                }
            }

            //Drawing is not open, so open it, then zoom.
            if (File.Exists(s) && s != null && isOpen == false)
            {
                documentCollection.Open(s, false);
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument();

                //Zoom To
                Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Database thisDatabase = thisDrawing.Database;
                Editor editor = thisDrawing.Editor;
                using (Transaction transaction1 = thisDatabase.TransactionManager.StartTransaction())
                {
                    try
                    {
                        //Check to see if our object's layout is in model space....
                        if (objectLayout == "Model")
                        {
                            //if our current layout is not Model
                            if (objectLayout != LayoutManager.Current.CurrentLayout)
                            {
                                //Switch to Model
                                LayoutManager.Current.CurrentLayout = objectLayout;
                            }
                        }

                        //Check to see if our object's layout is in a paper space layout
                        if (objectLayout != "Model")
                        {
                            //Check to make sure it's not our current layout and...
                            if (objectLayout != LayoutManager.Current.CurrentLayout)
                            {
                                //switch to it if it isn't
                                LayoutManager.Current.CurrentLayout = objectLayout;
                            }
                        }

                        using (ViewTableRecord view = editor.GetCurrentView())
                        {
                            Matrix3d WCS2DCS =
                                (Matrix3d.Rotation(-view.ViewTwist, view.ViewDirection, view.Target) *
                                Matrix3d.Displacement(view.Target - Point3d.Origin) *
                                Matrix3d.PlaneToWorld(view.ViewDirection))
                                .Inverse();
                            ObjectId objectId = thisDatabase.GetObjectId(false, handle, 0);
                            Entity entity = (Entity)transaction1.GetObject(objectId, OpenMode.ForRead);
                            Extents3d extents3D = entity.GeometricExtents;
                            extents3D.AddExtents(entity.GeometricExtents);
                            extents3D.TransformBy(WCS2DCS);
                            view.Width = extents3D.MaxPoint.X - extents3D.MinPoint.X;
                            view.Height = extents3D.MaxPoint.Y - extents3D.MinPoint.Y;
                            view.CenterPoint =
                                new Point2d((extents3D.MaxPoint.X + extents3D.MinPoint.X) / 2.0, (extents3D.MaxPoint.Y + extents3D.MinPoint.Y) / 2.0);
                            editor.SetCurrentView(view);
                            transaction1.Commit();
                        }
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }

            }
            else
            {
                return;
            }
        }
        public string GetCurrentDrawingPath()
        {
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            string Drawingpath = Path.GetFileName(thisDrawing.Name);

            return Drawingpath;
        }
        #endregion

    }
}
