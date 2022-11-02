using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;

namespace LayerManager.Helpers.AutoCADHelpers
{
    public class DocumentHelpers
    {
        public static void OpenAutoCADDrawing(string path)
        {
            try
            {
                DocumentCollection documentManager = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
                DocumentCollectionExtension.Open(documentManager, path);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public static void MakeFirstLayoutActive(Transaction transaction, Database database)
        {
            HostApplicationServices.WorkingDatabase = database;
            DBDictionary dBDictionaryEntries = (DBDictionary)transaction.GetObject(database.LayoutDictionaryId, OpenMode.ForRead);
            LayoutManager layoutManager = (LayoutManager)Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current;
            Layout layout = null;
            foreach (DBDictionaryEntry dBDictionaryEntry in dBDictionaryEntries)
            {
                layout = (Layout)transaction.GetObject(layoutManager.GetLayoutId(dBDictionaryEntry.Key), OpenMode.ForRead);
                if (layout.TabOrder == 1)
                {
                    layoutManager.CurrentLayout = layout.LayoutName;
                    return;
                }
            }
        }

        public static BlockTableRecord GetPaperspace(Transaction transaction, Database database)
        {

            HostApplicationServices.WorkingDatabase = database;
            DBDictionary dBDictionaryEntries = (DBDictionary)transaction.GetObject(database.LayoutDictionaryId, OpenMode.ForRead);
            LayoutManager layoutManager = (LayoutManager)Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current;
            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord paperspace = null;
            foreach (DBDictionaryEntry dBDictionaryEntry in dBDictionaryEntries)
            {
                Layout layout = (Layout)transaction.GetObject(layoutManager.GetLayoutId(dBDictionaryEntry.Key), OpenMode.ForRead);
                if (layout.TabOrder == 1)
                {
                    return (BlockTableRecord)transaction.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                }
            }
            return paperspace;
        }

        public static Layout GetFirstLayout(Transaction thisTransaction, Database thisDatabase)
        {
            HostApplicationServices.WorkingDatabase = thisDatabase;
            DBDictionary dBDictionary = (DBDictionary)thisTransaction.GetObject(thisDatabase.LayoutDictionaryId, OpenMode.ForRead);
            LayoutManager layoutManager = (LayoutManager)Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current;
            Layout layout = null;
            foreach (DBDictionaryEntry dBDictionaryEntry in dBDictionary)
            {
                layout = (Layout)thisTransaction.GetObject(layoutManager.GetLayoutId(dBDictionaryEntry.Key), OpenMode.ForRead);
                if (layout.TabOrder == 1)
                {
                    return layout;
                }
            }
            return layout;
        }

        public static Viewport CreateViewport(Viewport viewport, Point2d viewCenter)
        {
            Viewport newViewport = new Viewport();

            newViewport.SetDatabaseDefaults();
            newViewport.CenterPoint = viewport.CenterPoint;
            newViewport.Height = viewport.Height;
            newViewport.Width = viewport.Width;
            newViewport.ViewDirection = viewport.ViewDirection;
            newViewport.ViewTarget = viewport.ViewTarget;
            newViewport.ViewCenter = viewCenter;
            newViewport.TwistAngle = viewport.TwistAngle;
            newViewport.CustomScale = viewport.CustomScale;
            newViewport.Locked = viewport.Locked;


            return newViewport;
        }
    }
}
