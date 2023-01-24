using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Microsoft.Win32;
using System.Collections.Generic;
using System.IO;
using System.Windows;

namespace GFAR.Core.CoreLogic
{
    public class DataAccess
    {
        public void FindAndReplaceText(string drawing, string findText, string replaceText, string searchParameter)
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
                                    //database2.ResolveXrefs(useThreadEngine: true, doNewOnly: true);

                                    HostApplicationServices.WorkingDatabase = database2;

                                    using (Transaction trans2 = database2.TransactionManager.StartTransaction())
                                    {
                                        if (searchParameter == "Model Space")
                                        {
                                            BlockTable blockTable = trans2.GetObject(database2.BlockTableId, OpenMode.ForWrite) as BlockTable;
                                            BlockTableRecord blockTableRecord = trans2.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                                            foreach (ObjectId objectId in blockTableRecord)
                                            {
                                                BlockReference blockReference = trans2.GetObject(objectId, OpenMode.ForWrite) as BlockReference;
                                                if (blockReference != null)
                                                {
                                                    AttributeCollection attributeCollection = blockReference.AttributeCollection;
                                                    foreach (ObjectId attributeReferenceID in attributeCollection)
                                                    {
                                                        AttributeReference attributeReference = trans2.GetObject(attributeReferenceID, OpenMode.ForWrite) as AttributeReference;
                                                        if (attributeReference.IsMTextAttribute == true)
                                                        {
                                                            if (attributeReference.MTextAttribute.Contents.Contains(findText))
                                                            {
                                                                MText newMtext = attributeReference.MTextAttribute;
                                                                newMtext.Contents = attributeReference.MTextAttribute.Contents.Replace(findText, replaceText);
                                                                attributeReference.MTextAttribute = newMtext;
                                                                
                                                            }
                                                        }
                                                        else
                                                        {
                                                            if (attributeReference.TextString.Contains(findText))
                                                            {
                                                                string oldValue = attributeReference.TextString;
                                                                string newValue = oldValue.Replace(findText, replaceText);
                                                                attributeReference.TextString = newValue;
                                                                
                                                            }
                                                        }
                                                    }
                                                }

                                                MLeader mLeader = trans2.GetObject(objectId, OpenMode.ForWrite) as MLeader;
                                                if (mLeader != null)
                                                {
                                                    if (mLeader.MText.Contents.Contains(findText))
                                                    {
                                                        MText newMLeaderMText = new MText();
                                                        newMLeaderMText.Contents = mLeader.MText.Contents.Replace(findText, replaceText);
                                                        mLeader.MText = newMLeaderMText;
                                                        
                                                    }
                                                }

                                                MText mText = trans2.GetObject(objectId, OpenMode.ForWrite) as MText;
                                                if (mText != null)
                                                {
                                                    if (mText.Contents.Contains(findText))
                                                    {
                                                        string oldValue = mText.Contents;
                                                        string newValue = oldValue.Replace(findText, replaceText);
                                                        mText.Contents = newValue;
                                                        
                                                    }
                                                }

                                                DBText dBText = trans2.GetObject(objectId, OpenMode.ForWrite) as DBText;
                                                if (dBText != null)
                                                {
                                                    if (dBText.TextString.Contains(findText))
                                                    {
                                                        string oldValue = dBText.TextString;
                                                        string newValue = dBText.TextString.Replace(findText, replaceText);
                                                        dBText.TextString = newValue;
                                                        
                                                    }
                                                }
                                            }
                                        }

                                        if (searchParameter == "Paper Space (All Layouts)")
                                        {
                                            DBDictionary dBDictionary = trans2.GetObject(database2.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                                            foreach (DBDictionaryEntry dBDictionaryEntry in dBDictionary)
                                            {
                                                Layout layout = dBDictionaryEntry.Value.GetObject(OpenMode.ForRead) as Layout;
                                                if (layout.LayoutName != "Model")
                                                {
                                                    BlockTableRecord blockTableRecord = trans2.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite) as BlockTableRecord;
                                                    foreach (ObjectId objectId in blockTableRecord)
                                                    {
                                                        BlockReference blockReference = trans2.GetObject(objectId, OpenMode.ForWrite) as BlockReference;
                                                        if (blockReference != null)
                                                        {
                                                            AttributeCollection attributeCollection = blockReference.AttributeCollection;
                                                            foreach (ObjectId attributeReferenceID in attributeCollection)
                                                            {
                                                                AttributeReference attributeReference = trans2.GetObject(attributeReferenceID, OpenMode.ForWrite) as AttributeReference;
                                                                if (attributeReference.IsMTextAttribute == true)
                                                                {
                                                                    if (attributeReference.MTextAttribute.Contents.Contains(findText))
                                                                    {
                                                                        MText newMtext = attributeReference.MTextAttribute;
                                                                        newMtext.Contents = attributeReference.MTextAttribute.Contents.Replace(findText, replaceText);
                                                                        attributeReference.MTextAttribute = newMtext;
                                                                        
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    if (attributeReference.TextString.Contains(findText))
                                                                    {
                                                                        string oldValue = attributeReference.TextString;
                                                                        string newValue = oldValue.Replace(findText, replaceText);
                                                                        attributeReference.TextString = newValue;
                                                                        
                                                                    }
                                                                }
                                                            }
                                                        }

                                                        MLeader mLeader = trans2.GetObject(objectId, OpenMode.ForWrite) as MLeader;
                                                        if (mLeader != null)
                                                        {
                                                            if (mLeader.MText.Contents.Contains(findText))
                                                            {
                                                                MText newMLeaderMText = new MText();
                                                                newMLeaderMText.Contents = mLeader.MText.Contents.Replace(findText, replaceText);
                                                                mLeader.MText = newMLeaderMText;
                                                                
                                                            }
                                                        }

                                                        MText mText = trans2.GetObject(objectId, OpenMode.ForWrite) as MText;
                                                        if (mText != null)
                                                        {
                                                            if (mText.Contents.Contains(findText))
                                                            {
                                                                string oldValue = mText.Contents;
                                                                string newValue = oldValue.Replace(findText, replaceText);
                                                                mText.Contents = newValue;
                                                                
                                                            }
                                                        }

                                                        DBText dBText = trans2.GetObject(objectId, OpenMode.ForWrite) as DBText;
                                                        if (dBText != null)
                                                        {
                                                            if (dBText.TextString.Contains(findText))
                                                            {
                                                                string oldValue = dBText.TextString;
                                                                string newValue = dBText.TextString.Replace(findText, replaceText);
                                                                dBText.TextString = newValue;
                                                                
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                        if (searchParameter == "Both")
                                        {
                                            BlockTable blockTable = trans2.GetObject(database2.BlockTableId, OpenMode.ForWrite) as BlockTable;
                                            BlockTableRecord blockTableRecord = trans2.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                                            foreach (ObjectId objectId in blockTableRecord)
                                            {
                                                BlockReference blockReference = trans2.GetObject(objectId, OpenMode.ForWrite) as BlockReference;
                                                if (blockReference != null)
                                                {
                                                    AttributeCollection attributeCollection = blockReference.AttributeCollection;
                                                    foreach (ObjectId attributeReferenceID in attributeCollection)
                                                    {
                                                        AttributeReference attributeReference = trans2.GetObject(attributeReferenceID, OpenMode.ForWrite) as AttributeReference;
                                                        if (attributeReference.IsMTextAttribute == true)
                                                        {
                                                            if (attributeReference.MTextAttribute.Contents.Contains(findText))
                                                            {
                                                                MText newMtext = attributeReference.MTextAttribute;
                                                                newMtext.Contents = attributeReference.MTextAttribute.Contents.Replace(findText, replaceText);
                                                                attributeReference.MTextAttribute = newMtext;
                                                                
                                                            }
                                                        }
                                                        else
                                                        {
                                                            if (attributeReference.TextString.Contains(findText))
                                                            {
                                                                string oldValue = attributeReference.TextString;
                                                                string newValue = oldValue.Replace(findText, replaceText);
                                                                attributeReference.TextString = newValue;
                                                                
                                                            }
                                                        }
                                                    }
                                                }

                                                MLeader mLeader = trans2.GetObject(objectId, OpenMode.ForWrite) as MLeader;
                                                if (mLeader != null)
                                                {
                                                    if (mLeader.MText.Contents.Contains(findText))
                                                    {
                                                        MText newMLeaderMText = new MText();
                                                        newMLeaderMText.Contents = mLeader.MText.Contents.Replace(findText, replaceText);
                                                        mLeader.MText = newMLeaderMText;
                                                        
                                                    }
                                                }

                                                MText mText = trans2.GetObject(objectId, OpenMode.ForWrite) as MText;
                                                if (mText != null)
                                                {
                                                    if (mText.Contents.Contains(findText))
                                                    {
                                                        string oldValue = mText.Contents;
                                                        string newValue = oldValue.Replace(findText, replaceText);
                                                        mText.Contents = newValue;
                                                        
                                                    }
                                                }

                                                DBText dBText = trans2.GetObject(objectId, OpenMode.ForWrite) as DBText;
                                                if (dBText != null)
                                                {
                                                    if (dBText.TextString.Contains(findText))
                                                    {
                                                        string oldValue = dBText.TextString;
                                                        string newValue = dBText.TextString.Replace(findText, replaceText);
                                                        dBText.TextString = newValue;
                                                        
                                                    }
                                                }
                                            }

                                            DBDictionary dBDictionary = trans2.GetObject(database2.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                                            foreach (DBDictionaryEntry dBDictionaryEntry in dBDictionary)
                                            {
                                                Layout layout = dBDictionaryEntry.Value.GetObject(OpenMode.ForRead) as Layout;
                                                if (layout.LayoutName != "Model")
                                                {
                                                    BlockTableRecord layoutBlockTableRecord = trans2.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite) as BlockTableRecord;
                                                    foreach (ObjectId objectId in layoutBlockTableRecord)
                                                    {
                                                        BlockReference blockReference = trans2.GetObject(objectId, OpenMode.ForWrite) as BlockReference;
                                                        if (blockReference != null)
                                                        {
                                                            AttributeCollection attributeCollection = blockReference.AttributeCollection;
                                                            foreach (ObjectId attributeReferenceID in attributeCollection)
                                                            {
                                                                AttributeReference attributeReference = trans2.GetObject(attributeReferenceID, OpenMode.ForWrite) as AttributeReference;
                                                                if (attributeReference.IsMTextAttribute == true)
                                                                {
                                                                    if (attributeReference.MTextAttribute.Contents.Contains(findText))
                                                                    {
                                                                        MText newMtext = attributeReference.MTextAttribute;
                                                                        newMtext.Contents = attributeReference.MTextAttribute.Contents.Replace(findText, replaceText);
                                                                        attributeReference.MTextAttribute = newMtext;
                                                                        
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    if (attributeReference.TextString.Contains(findText))
                                                                    {
                                                                        string oldValue = attributeReference.TextString;
                                                                        string newValue = oldValue.Replace(findText, replaceText);
                                                                        attributeReference.TextString = newValue;
                                                                        
                                                                    }
                                                                }
                                                            }
                                                        }

                                                        MLeader mLeader = trans2.GetObject(objectId, OpenMode.ForWrite) as MLeader;
                                                        if (mLeader != null)
                                                        {
                                                            if (mLeader.MText.Contents.Contains(findText))
                                                            {
                                                                MText newMLeaderMText = new MText();
                                                                newMLeaderMText.Contents = mLeader.MText.Contents.Replace(findText, replaceText);
                                                                mLeader.MText = newMLeaderMText;
                                                                
                                                            }
                                                        }

                                                        MText mText = trans2.GetObject(objectId, OpenMode.ForWrite) as MText;
                                                        if (mText != null)
                                                        {
                                                            if (mText.Contents.Contains(findText))
                                                            {
                                                                string oldValue = mText.Contents;
                                                                string newValue = oldValue.Replace(findText, replaceText);
                                                                mText.Contents = newValue;
                                                                
                                                            }
                                                        }

                                                        DBText dBText = trans2.GetObject(objectId, OpenMode.ForWrite) as DBText;
                                                        if (dBText != null)
                                                        {
                                                            if (dBText.TextString.Contains(findText))
                                                            {
                                                                string oldValue = dBText.TextString;
                                                                string newValue = dBText.TextString.Replace(findText, replaceText);
                                                                dBText.TextString = newValue;
                                                                
                                                            }
                                                        }
                                                    }
                                                }
                                            }
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


        public void FindAndReplaceTextInPaperSpace(string drawing, string findText, string replaceText)
        {
            System.Windows.Forms.DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Added drawings will be modified and saved. This process cannot be undone. Continue?", "Save drawings?", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);
            if (dialogResult == System.Windows.Forms.DialogResult.Yes)
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
                                        DBDictionary dBDictionary = trans2.GetObject(database2.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                                        foreach (DBDictionaryEntry dBDictionaryEntry in dBDictionary)
                                        {
                                            Layout layout = dBDictionaryEntry.Value.GetObject(OpenMode.ForRead) as Layout;
                                            if (layout.LayoutName != "Model")
                                            {
                                                BlockTableRecord blockTableRecord = trans2.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite) as BlockTableRecord;
                                                foreach (ObjectId objectId in blockTableRecord)
                                                {
                                                    BlockReference blockReference = trans2.GetObject(objectId, OpenMode.ForWrite) as BlockReference;
                                                    if (blockReference != null)
                                                    {
                                                        AttributeCollection attributeCollection = blockReference.AttributeCollection;
                                                        foreach (ObjectId attributeReferenceID in attributeCollection)
                                                        {
                                                            AttributeReference attributeReference = trans2.GetObject(attributeReferenceID, OpenMode.ForWrite) as AttributeReference;
                                                            if (attributeReference.IsMTextAttribute == true)
                                                            {
                                                                if (attributeReference.MTextAttribute.Contents.Contains(findText))
                                                                {
                                                                    MText newMtext = attributeReference.MTextAttribute;
                                                                    newMtext.Contents = attributeReference.MTextAttribute.Contents.Replace(findText, replaceText);
                                                                    attributeReference.MTextAttribute = newMtext;
                                                                }
                                                            }
                                                            else
                                                            {
                                                                if (attributeReference.TextString.Contains(findText))
                                                                {
                                                                    string oldValue = attributeReference.TextString;
                                                                    string newValue = oldValue.Replace(findText, replaceText);
                                                                    attributeReference.TextString = newValue;
                                                                }
                                                            }
                                                        }
                                                    }

                                                    MLeader mLeader = trans2.GetObject(objectId, OpenMode.ForWrite) as MLeader;
                                                    if (mLeader != null)
                                                    {
                                                        if (mLeader.MText.Contents.Contains(findText))
                                                        {
                                                            MText newMLeaderMText = new MText();
                                                            newMLeaderMText.Contents = mLeader.MText.Contents.Replace(findText, replaceText);
                                                            mLeader.MText = newMLeaderMText;
                                                        }
                                                    }

                                                    MText mText = trans2.GetObject(objectId, OpenMode.ForWrite) as MText;
                                                    if (mText != null)
                                                    {
                                                        if (mText.Contents.Contains(findText))
                                                        {
                                                            string oldValue = mText.Contents;
                                                            string newValue = oldValue.Replace(findText, replaceText);
                                                            mText.Contents = newValue;
                                                        }
                                                    }

                                                    DBText dBText = trans2.GetObject(objectId, OpenMode.ForWrite) as DBText;
                                                    if (dBText != null)
                                                    {
                                                        if (dBText.TextString.Contains(findText))
                                                        {
                                                            string oldValue = dBText.TextString;
                                                            string newValue = dBText.TextString.Replace(findText, replaceText);
                                                            dBText.TextString = newValue;
                                                        }
                                                    }
                                                }
                                            }
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
                MessageBox.Show("Done.");
            }
        }


        /// <summary>
        /// Creates a list of drawing files by asking the user to select them with an <see cref="OpenFileDialog"/> and returns that list.
        /// </summary>
        /// <returns></returns>
        public IList<string> AddtoDrawingList()
        {
            var items = new List<string>();

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Select The Target Drawings To Add To Add To The List";
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

    }
}
