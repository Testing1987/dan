using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Internal;
using BlockManager.Core.CoreLogic;
using BlockManager.Core.Models;
using CADTechnologiesSource.All.AutoCADHelpers;
using CADTechnologiesSource.All.Base;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace BlockManager.UI.ViewModels
{
    public class MainWindow_BM_ViewModel : BaseViewModel
    {
        #region Singleton

        public static MainWindow_BM_ViewModel Instance = new MainWindow_BM_ViewModel();

        #endregion

        #region Public Properties

        #region Bam
        public ObservableCollection<TargetDrawingModelBAM> TargetDrawings { get; set; }
        public TargetDrawingModelBAM SelectedTargetDrawing { get; set; }
        public ObservableCollection<BlockItemViewModel> BlockItems { get; set; }
        public ObservableCollection<string> Filters { get; set; }
        public ObservableCollection<string> BlockFilters { get; set; }
        public string SelectedFilter { get; set; }
        public string BlockSelectedFilter { get; set; }
        #endregion


        #region Insert

        BlockTableRecord InsertBlock = null;
        public string InsertBlockName { get; set; }
        public string InsertBlockHostDrawingPath { get; set; }
        public bool InsertInModelRadio { get; set; }
        public bool InsertInLayoutRadio { get; set; }
        public bool InsertInAllLayoutsCheckbox { get; set; }
        public int InsertLayoutIndex { get; set; }
        public Point3d InsertionPoint { get; set; }
        public double InsertX { get; set; }
        public double InsertY { get; set; }
        public double InsertZ { get; set; }
        public bool InsertNewReferenceRadio { get; set; }
        public bool ReplaceBlockRadio { get; set; }
        public bool DoNothingRadio { get; set; }
        public bool SpecifyInsertionPoint { get; set; }
        public bool CloneFromCurrentDrawing { get; set; }
        public bool CloneFromExternalDrawing { get; set; }

        #endregion

        #endregion

        #region PublicCommands

        #region Drawing List Controls
        public ICommand AddCommand { get; set; }
        public ICommand RemoveCommand { get; set; }
        public ICommand ClearCommand { get; set; }
        public ICommand OpenCommand { get; set; }
        public ICommand SaveCommand { get; set; }
        #endregion


        #region Block Item Buttons
        public ICommand RefreshBlockItemCommand { get; set; }
        public ICommand UpdateBlockCommand { get; set; }
        public ICommand FiltersChangedCommand { get; set; }
        public ICommand ZoomCommand { get; set; }
        #endregion

        #region Insert Blocks Commands
        public ICommand GetFileNameFromSelectionCommand { get; set; }
        public ICommand SelectBlockFromCurrentDrawingCommand { get; set; }
        public ICommand SelectInsertionPointCommand { get; set; }
        public ICommand InsertBlocksCommand { get; set; }
        #endregion


        #endregion

        #region Constructor
        public MainWindow_BM_ViewModel()
        {
            //Instantiate the collections
            if (TargetDrawings == null)
            {
                TargetDrawings = new ObservableCollection<TargetDrawingModelBAM>();
            }
            if (BlockItems == null)
            {
                BlockItems = new ObservableCollection<BlockItemViewModel>();
            }
            if (Filters == null)
            {
                Filters = new ObservableCollection<string>();
                Filters.Add("All");
            }
            if (BlockFilters == null)
            {
                BlockFilters = new ObservableCollection<string>();
                BlockFilters.Add("All");
            }

            InsertInModelRadio = true;
            InsertInLayoutRadio = false;

            InsertNewReferenceRadio = false;
            ReplaceBlockRadio = true;
            DoNothingRadio = false;

            CloneFromCurrentDrawing = true;
            CloneFromExternalDrawing = false;

            //Define the commands that exist on the mainwindow_bm
            AddCommand = new RelayCommand(() => AddCommandClick());
            RemoveCommand = new RelayCommand(() => RemoveCommandClick());
            ClearCommand = new RelayCommand(() => ClearCommandClick());
            SaveCommand = new RelayCommand(() => SaveCommandClick());
            OpenCommand = new RelayCommand(() => OpenCommandClick());
            FiltersChangedCommand = new RelayCommand(() => FiltersChangedCommandClick());
            GetFileNameFromSelectionCommand = new RelayCommand(() => GetFileNameFromSelectionCommandClick());
            SelectBlockFromCurrentDrawingCommand = new RelayCommand(() => SelectBlockFromCurrentDrawingCommandClick());
            InsertBlocksCommand = new RelayCommand(() => InsertBlocksCommandClick());
            SelectInsertionPointCommand = new RelayCommand(() => SelectInsertionPointCommandClick());
        }

        private void OpenCommandClick()
        {
            DocumentHelpers.OpenAutoCADDrawing(SelectedTargetDrawing.DrawingPath);
        }

        #endregion

        #region Command Methods

        #region BAM Commands
        private void SaveCommandClick()
        {
            MessageBox.Show("You Pressed Save");
        }

        /// <summary>
        /// Clears the target drawing list which in-turn removes all blockitems.
        /// </summary>
        private void ClearCommandClick()
        {
            try
            {
                TargetDrawings.Clear();
                BlockItems.Clear();
                var newdrawingFilter = Filters.Where(x => x == "All").ToList().AsParallel();
                Filters = new ObservableCollection<string>(newdrawingFilter);

                var newblockFilter = Filters.Where(x => x == "All").ToList().AsParallel();
                BlockFilters = new ObservableCollection<string>(newblockFilter);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Removes the selected target drawing from the list
        /// </summary>
        private void RemoveCommandClick()
        {
            try
            {
                if (TargetDrawings.Count > 0 && SelectedTargetDrawing != null)
                {
                    var newBlockItems = BlockItems.Where(x => x.Drawing != SelectedTargetDrawing.DrawingPath).ToList().AsParallel();
                    BlockItems = new ObservableCollection<BlockItemViewModel>(newBlockItems);
                    Filters.Remove(SelectedTargetDrawing.DrawingPath);
                    TargetDrawings.Remove(SelectedTargetDrawing);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Adds target drawings to the list
        /// </summary>
        private void AddCommandClick()
        {
            DataAccess dataAcces = new DataAccess();
            foreach (var result in dataAcces.AddtoTargetDrawingList())
            {
                if (result != null)
                {
                    if (TargetDrawings.Any(p => p.DrawingPath == result.DrawingPath) == false)
                        TargetDrawings.Add(result);

                    //Create the blockItemViewModels by taking the collection of target drawings and reading the attribute blocks in the drawings
                    foreach (var blockModel in dataAcces.GetAttributedBlocksFromDrawings(result.DrawingPath))
                    {
                        if (blockModel != null && BlockItems != null)
                        {
                            //Create and populate the blockitemviewmodels
                            BlockItemViewModel blockItemViewModel = new BlockItemViewModel();
                            blockItemViewModel.BlockName = blockModel.BlockName;
                            blockItemViewModel.Drawing = blockModel.HostDrawing;
                            blockItemViewModel.BlockHandle = blockModel.BlockHandle;
                            blockItemViewModel.BlockObjectId = blockModel.BlockObjectId;
                            blockItemViewModel.SavedDateTime = File.GetLastWriteTime(blockItemViewModel.Drawing).ToString();
                            blockItemViewModel.BlockLocation = blockModel.BlockLocation;
                            blockItemViewModel.LayoutName = blockModel.LayoutName;
                            blockItemViewModel.StringBlockLocation = blockModel.StringBlockLocation;
                            blockItemViewModel.CurrentSpace = blockModel.CurrentSpace;

                            //Create the attribute item viewmodels for each attribute inside the block
                            foreach (var attribute in blockModel.Attributes)
                            {
                                AttributeItemViewModel attributeItemViewModel = new AttributeItemViewModel();
                                attributeItemViewModel.AttributeTag = attribute.AttributeTag;
                                attributeItemViewModel.AttributeValue = attribute.AttributeValue;
                                if (blockItemViewModel.AttributesList == null)
                                {
                                    blockItemViewModel.AttributesList = new ObservableCollection<AttributeItemViewModel>();
                                }
                                blockItemViewModel.AttributesList.Add(attributeItemViewModel);
                            }

                            //Instantiate the comands ont he blockitemmodels and pass the correct information into the associated command methods
                            UpdateBlockCommand = new RelayCommand(() => UpdateBlockCommandClick(blockItemViewModel));
                            OpenCommand = new RelayCommand(() => OpenCommandClick(blockItemViewModel.Drawing));
                            ZoomCommand = new RelayCommand(() => ZoomCommandClick(blockItemViewModel.Drawing, blockItemViewModel.BlockHandle, blockItemViewModel.LayoutName));
                            blockItemViewModel.RefreshCommand = RefreshBlockItemCommand;
                            blockItemViewModel.UpdateCommand = UpdateBlockCommand;
                            blockItemViewModel.OpenCommand = OpenCommand;
                            blockItemViewModel.ZoomCommand = ZoomCommand;
                            if (Filters.Any(p => p == blockItemViewModel.Drawing) == false)
                            {
                                Filters.Add(blockItemViewModel.Drawing);
                            }
                            if (BlockFilters.Any(p => p == blockItemViewModel.BlockName) == false)
                            {
                                BlockFilters.Add(blockItemViewModel.BlockName);
                            }
                            blockItemViewModel.MyVisibility = System.Windows.Visibility.Visible;
                            BlockItems.Add(blockItemViewModel);
                            SelectedFilter = "All";
                            BlockSelectedFilter = "All";
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Opens the drawing represented by the item you clicked on.
        /// </summary>
        /// <param name="drawing"></param>
        private void OpenCommandClick(string drawing)
        {
            DatabaseHelpers databaseHelpers = new DatabaseHelpers();
            databaseHelpers.OpenDrawing(drawing);
        }

        /// <summary>
        /// Opens the drawing that hosts the block and zooms to the selected block.
        /// </summary>
        /// <param name="drawing"></param>
        /// <param name="blockHandle"></param>
        /// <param name="objectLayout"></param>
        private void ZoomCommandClick(string drawing, Handle blockHandle, string objectLayout)
        {
            DatabaseHelpers databaseHelpers = new DatabaseHelpers();
            databaseHelpers.OpenDrawingAndZoomTo(drawing, blockHandle, objectLayout);
        }

        /// <summary>
        /// Accessess AutoCAD and updates the block reference attributes with any changes made by the user.
        /// </summary>
        /// <param name="blockItemViewModel"></param>
        private void UpdateBlockCommandClick(BlockItemViewModel blockItemViewModel)
        {
            //MessageBox.Show($"You Pressed Update on {blockItemViewModel.Drawing}");
            DataAccess dataAccess = new DataAccess();
            string drawing = blockItemViewModel.Drawing;
            Handle blockHandle = blockItemViewModel.BlockHandle;
            ObjectId blockObjectId = blockItemViewModel.BlockObjectId;
            string blockName = blockItemViewModel.BlockName;
            string currentSpace = blockItemViewModel.CurrentSpace;

            //DON'T REMEMBER WHY I ADDED THIS. I SHOULD HAVE COMMENTED WHEN I DID THIS. THE BLOCKITEMS ATTRIBUTE LIST SHOULD ALREADY BE CURRENT WHEN THEY ARE ADDED.
            //foreach (var blockItem in BlockItems)
            //{
            //    if (blockItem.BlockObjectId == blockItemViewModel.BlockObjectId)
            //    {
            //        blockItemViewModel.AttributesList = blockItem.AttributesList;
            //    }
            //}

            //Populate the blockmodels attributes list
            //Create AttributeModels for each attribute on the attributes list so it can be passed in.
            ObservableCollection<AttributeModel> attributeModels = new ObservableCollection<AttributeModel>();
            foreach (var attributeItem in blockItemViewModel.AttributesList)
            {
                AttributeModel attributeModel = new AttributeModel();
                attributeModel.AttributeTag = attributeItem.AttributeTag;
                attributeModel.AttributeValue = attributeItem.AttributeValue;
                attributeModels.Add(attributeModel);
            }
            dataAccess.UpdateBlock(drawing, currentSpace, blockHandle, blockName, attributeModels);
            blockItemViewModel.SavedDateTime = File.GetLastWriteTime(blockItemViewModel.Drawing).ToString();
            foreach (var item in BlockItems)
            {
                if (item.Drawing == blockItemViewModel.Drawing)
                {
                    item.SavedDateTime = blockItemViewModel.SavedDateTime = File.GetLastWriteTime(blockItemViewModel.Drawing).ToString();
                }
            }
        }

        /// <summary>
        /// Filters the visibility of blockitems based on the selected drawing from the combobox.
        /// </summary>
        private void FiltersChangedCommandClick()
        {
            if (BlockItems != null)
            {
                foreach (var blockItem in BlockItems)
                {
                    if (SelectedFilter == "All")
                    {
                        blockItem.MyVisibility = System.Windows.Visibility.Visible;
                    }
                    else
                    {
                        if (SelectedFilter == blockItem.Drawing)
                        {
                            blockItem.MyVisibility = System.Windows.Visibility.Visible;
                        }
                        else
                        {
                            if (SelectedFilter != blockItem.Drawing)
                            {
                                blockItem.MyVisibility = System.Windows.Visibility.Collapsed;
                            }
                        }
                    }
                }
            }
        }
        #endregion


        #region Insert Commands
        private void GetFileNameFromSelectionCommandClick()
        {
            DataAccess da = new DataAccess();
            string hostDrawingPath = da.GetDrawingPath();
            if (string.IsNullOrEmpty(hostDrawingPath) == false)
            {
                InsertBlockHostDrawingPath = hostDrawingPath;
            }
        }

        private void SelectInsertionPointCommandClick()
        {
            DataAccess da = new DataAccess();
            InsertionPoint = da.GetInsertionPointForBlock();
            InsertX = InsertionPoint.X;
            InsertY = InsertionPoint.Y;
            InsertZ = InsertionPoint.Z;
        }

        private void SelectBlockFromCurrentDrawingCommandClick()
        {
            DataAccess da = new DataAccess();

            BlockTableRecord btr = null;
            //btr = da.GetBlockBySelection();
            if (btr != null)
            {
                InsertBlock = btr;
                InsertBlockName = btr.Name;
                string currentDrawingPath = da.GetCurrentDWGPath();
                if (string.IsNullOrEmpty(currentDrawingPath) == false)
                {
                    InsertBlockHostDrawingPath = currentDrawingPath;
                }
            }
        }

        private void InsertBlocksCommandClick()
        {
            System.Windows.Forms.DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Target drawings will be modified and saved. " +
                    "This process cannot be undone. Continue?", "Save drawings?", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);
            if (dialogResult == System.Windows.Forms.DialogResult.Yes)
            {
                try
                {
                    DataAccess da = new DataAccess();
                    if (TargetDrawings != null && TargetDrawings.Count > 0)
                    {


                        //make sure the user didn't add the source drawing to the target drawings
                        bool isSourceOnTargetList = TargetDrawings.Any(target => target.DrawingPath == InsertBlockHostDrawingPath);
                        if (isSourceOnTargetList == true)
                        {
                            MessageBox.Show(
                                    "You selected a target drawing as the host. You must selected a host that is not on the list of target drawings.",
                                    "Host is Target Drawing",
                                    MessageBoxButton.OK,
                                    MessageBoxImage.Error);
                            return;
                        }

                        InsertionPoint = new Point3d(InsertX, InsertY, InsertZ);


                        if (CloneFromCurrentDrawing == true)
                        {
                            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                            Database thisDatabase = thisDrawing.Database;
                            Editor thisEditor = thisDrawing.Editor;
                            Utils.SetFocusToDwgView();
                            PromptEntityOptions peo = new PromptEntityOptions("Select a block");
                            peo.AllowNone = false;

                            PromptEntityResult per = thisEditor.GetEntity(peo);
                            if (per.Status != PromptStatus.OK)
                            {
                                return;
                            }
                            else
                            {
                                ObjectId perId = per.ObjectId;
                                using (DocumentLock docLock = thisDrawing.LockDocument())
                                {
                                    using (Transaction thisTransaction = thisDatabase.TransactionManager.StartTransaction())
                                    {
                                        //BlockTable blockTable = thisTransaction.GetObject(thisDatabase.BlockTableId, OpenMode.ForRead) as BlockTable;
                                        try
                                        {
                                            BlockReference selectedBlockReference = thisTransaction.GetObject(perId, OpenMode.ForRead) as BlockReference;
                                            InsertBlockName = selectedBlockReference.Name.ToLower();
                                            InsertionPoint = selectedBlockReference.Position;
                                        }
                                        catch (Exception)
                                        {
                                            MessageBox.Show("The selected object is not a block.", "Not a block.", MessageBoxButton.OK, MessageBoxImage.Information);
                                        }
                                        thisTransaction.Commit();
                                    }
                                }
                            }
                            foreach (var targetDrawing in TargetDrawings)
                            {
                                da.CloneBlockFromCurrentDrawing(InsertBlockName, targetDrawing.DrawingPath, InsertInModelRadio, InsertLayoutIndex, InsertInAllLayoutsCheckbox, InsertionPoint, InsertNewReferenceRadio, ReplaceBlockRadio, DoNothingRadio);
                                da.AuditDrawing(targetDrawing.DrawingPath);
                            }
                        }

                        if (CloneFromExternalDrawing == true)
                        {
                            //Make sure the user has the block name entered
                            if (string.IsNullOrEmpty(InsertBlockName))
                            {
                                MessageBox.Show(
                                    "You must enter the name of the block you want to clone.",
                                    "No name given.",
                                      MessageBoxButton.OK,
                                      MessageBoxImage.Warning);
                                return;
                            }

                            foreach (var targetDrawing in TargetDrawings)
                            {
                                da.CloneBlockFromExternalDrawing(InsertBlockName.ToLower(), InsertBlockHostDrawingPath, targetDrawing.DrawingPath, InsertInModelRadio, InsertLayoutIndex, InsertInAllLayoutsCheckbox, InsertionPoint, InsertNewReferenceRadio, ReplaceBlockRadio, DoNothingRadio);
                                da.AuditDrawing(targetDrawing.DrawingPath);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

            //reload the block models
            if (TargetDrawings != null && TargetDrawings.Count > 0)
            {
                BlockItems.Clear();
                DataAccess dataAccess = new DataAccess();
                foreach (var targetDrawing in TargetDrawings)
                {
                    //Create the blockItemViewModels by taking the collection of target drawings and reading the attribute blocks in the drawings
                    foreach (var blockModel in dataAccess.GetAttributedBlocksFromDrawings(targetDrawing.DrawingPath))
                    {
                        if (blockModel != null && BlockItems != null)
                        {
                            //Create and populate the blockitemviewmodels
                            BlockItemViewModel blockItemViewModel = new BlockItemViewModel();
                            blockItemViewModel.BlockName = blockModel.BlockName;
                            blockItemViewModel.Drawing = blockModel.HostDrawing;
                            blockItemViewModel.BlockHandle = blockModel.BlockHandle;
                            blockItemViewModel.BlockObjectId = blockModel.BlockObjectId;
                            blockItemViewModel.SavedDateTime = File.GetLastWriteTime(blockItemViewModel.Drawing).ToString();
                            blockItemViewModel.BlockLocation = blockModel.BlockLocation;
                            blockItemViewModel.LayoutName = blockModel.LayoutName;
                            blockItemViewModel.StringBlockLocation = blockModel.StringBlockLocation;
                            blockItemViewModel.CurrentSpace = blockModel.CurrentSpace;

                            //Create the attribute item viewmodels for each attribute inside the block
                            foreach (var attribute in blockModel.Attributes)
                            {
                                AttributeItemViewModel attributeItemViewModel = new AttributeItemViewModel();
                                attributeItemViewModel.AttributeTag = attribute.AttributeTag;
                                attributeItemViewModel.AttributeValue = attribute.AttributeValue;
                                if (blockItemViewModel.AttributesList == null)
                                {
                                    blockItemViewModel.AttributesList = new ObservableCollection<AttributeItemViewModel>();
                                }
                                blockItemViewModel.AttributesList.Add(attributeItemViewModel);
                            }

                            //Instantiate the comands ont he blockitemmodels and pass the correct information into the associated command methods
                            UpdateBlockCommand = new RelayCommand(() => UpdateBlockCommandClick(blockItemViewModel));
                            OpenCommand = new RelayCommand(() => OpenCommandClick(blockItemViewModel.Drawing));
                            ZoomCommand = new RelayCommand(() => ZoomCommandClick(blockItemViewModel.Drawing, blockItemViewModel.BlockHandle, blockItemViewModel.LayoutName));
                            blockItemViewModel.RefreshCommand = RefreshBlockItemCommand;
                            blockItemViewModel.UpdateCommand = UpdateBlockCommand;
                            blockItemViewModel.OpenCommand = OpenCommand;
                            blockItemViewModel.ZoomCommand = ZoomCommand;
                            if (Filters.Any(p => p == blockItemViewModel.Drawing) == false)
                            {
                                Filters.Add(blockItemViewModel.Drawing);
                            }
                            if (BlockFilters.Any(p => p == blockItemViewModel.BlockName) == false)
                            {
                                BlockFilters.Add(blockItemViewModel.BlockName);
                            }
                            blockItemViewModel.MyVisibility = System.Windows.Visibility.Visible;
                            BlockItems.Add(blockItemViewModel);
                            SelectedFilter = "All";
                            BlockSelectedFilter = "All";
                        }
                    }
                }
            }


            else
            {
                return;
            }
        }
        #endregion


        #endregion
    }
}
