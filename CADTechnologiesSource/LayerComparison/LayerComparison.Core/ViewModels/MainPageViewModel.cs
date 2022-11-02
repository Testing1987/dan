using CADTechnologiesSource.All.AutoCADHelpers;
using CADTechnologiesSource.All.Base;
using LayerComparison.Core.Constants;
using LayerComparison.Core.CoreLogic;
using LayerComparison.Core.Enums;
using LayerComparison.Core.Models;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using LayerComparison.Core.IoCContainer;
using System.Diagnostics;
using CADTechnologiesSource.All.Models;

namespace LayerComparison.Core.ViewModels
{
    public class MainPageViewModel : BaseViewModel
    {
        #region Singleton
        /// <summary>
        /// Creates a single instance of the <see cref="MainPageViewModel"/> to bind to.
        /// </summary>
        public static MainPageViewModel Instance = new MainPageViewModel();
        #endregion

        #region Public Properties
        public string ElapsedTime { get; set; }
        public bool ComparisonRan { get; set; } = false;
        public ObservableCollection<TargetDrawingModel> TargetDrawings { get; set; }
        public ObservableCollection<LayerItemViewModel> LayerItemList { get; set; }
        public ObservableCollection<LayerItemViewModel> LayerControls { get; set; }
        public ObservableCollection<ConflictModel> ConflictReport { get; set; }
        public ObservableCollection<LayerComparisonLayerModel> SourceDrawingLayers { get; set; }
        public ObservableCollection<LayerComparisonLayerModel> TargetDrawingLayers { get; set; }
        public ObservableCollection<LayerComparisonLayerModel> AllConflictLayers { get; set; }
        public ObservableCollection<LayerComparisonLayerModel> OnOffConflictLayers { get; set; }
        public ObservableCollection<LayerComparisonLayerModel> FreezeConflictLayers { get; set; }
        public ObservableCollection<LayerComparisonLayerModel> ColorConflictLayers { get; set; }
        public ObservableCollection<LayerComparisonLayerModel> LinetypeConflictLayers { get; set; }
        public ObservableCollection<LayerComparisonLayerModel> LineweightConflictLayers { get; set; }
        public ObservableCollection<LayerComparisonLayerModel> TransparencyConflictLayers { get; set; }
        public ObservableCollection<LayerComparisonLayerModel> PlotConflictLayers { get; set; }
        public ObservableCollection<LayerComparisonLayerModel> ExtraLayers { get; set; }
        public ObservableCollection<MissingLayerModel> MissingLayers { get; set; }
        public ObservableCollection<LayerComparisonViewportLayerModel> SourceDrawingViewportLayers { get; set; }
        public ObservableCollection<LayerComparisonViewportLayerModel> TargetDrawingViewportLayers { get; set; }
        public ObservableCollection<LayerComparisonViewportLayerModel> ViewportFreezeConflictLayers { get; set; }
        public ObservableCollection<LayerComparisonViewportLayerModel> ViewportColorConflictLayers { get; set; }
        public ObservableCollection<LayerComparisonViewportLayerModel> ViewportLinetypeConflictLayers { get; set; }
        public ObservableCollection<LayerComparisonViewportLayerModel> ViewportLineweightConflictLayers { get; set; }
        public ObservableCollection<LayerComparisonViewportLayerModel> ViewportTransparencyConflictLayers { get; set; }
        public ObservableCollection<DrawingModel> DrawingModels { get; set; }
        public LayerComparisonLayerModel SelectedConflict { get; set; }
        public LayerComparisonViewportLayerModel SelectedViewportConflict { get; set; }
        public TargetDrawingModel SelectedTargetDrawing { get; set; }
        public List<string> SavedComparison { get; set; }
        public List<RecentItemViewModel> RecentItems { get; set; }
        public int MainPageTabIndex { get; set; }
        public int SelectedConflictTabIndex { get; set; }
        public int SelectedAdvancedTabIndex { get; set; }
        public int SelectedViewportConflictTabIndex { get; set; }
        public string TitleText { get; set; } = "MottMacDonald Layer Comparison - BETA";
        public string SourceDrawing { get; set; } = MessagesAndNotifications.PleaseLoadSourceDrawing;
        public string TotalConflictsCount { get; set; }
        public string OnOffConflictsCount { get; set; }
        public string FreezeConflictsCount { get; set; }
        public string ColorConflictsCount { get; set; }
        public string LinetypeConflictsCount { get; set; }
        public string LineweightConflictsCount { get; set; }
        public string TransparencyConflictsCount { get; set; }
        public string PlotConflictsCount { get; set; }
        public string ExtraLayersCount { get; set; }
        public string MissingLayersCount { get; set; }
        public string ViewportFreezeConflictsCount { get; set; }
        public string ViewportColorConflictsCount { get; set; }
        public string ViewportLinetypeConflictsCount { get; set; }
        public string ViewportLineweightConflictsCount { get; set; }
        public string ViewportTransparencyConflictsCount { get; set; }

        #endregion

        #region Public Commands
        public ICommand SaveCommand { get; set; }
        public ICommand NewComparisonCommand { get; set; }
        public ICommand LoadCommand { get; set; }
        public ICommand SaveRecentCommand { get; set; }
        public ICommand BackCommand { get; set; }
        public ICommand SetSourceDrawingCommand { get; set; }
        public ICommand AddTargetDrawingCommand { get; set; }
        public ICommand RemoveTargetDrawingCommand { get; set; }
        public ICommand RemoveViewportConflictCommand { get; set; }
        public ICommand OpenSelectedTargetDrawingCommand { get; set; }
        public ICommand OpenSourceDrawingCommand { get; set; }
        public ICommand ReviewLayerStyleCommand { get; set; }
        public ICommand CompareCommand { get; set; }
        public ICommand ResolveOnOffConflictsCommand { get; set; }
        public ICommand ResolveFreezeThawConflictsCommand { get; set; }
        public ICommand ResolveColorConflictsCommand { get; set; }
        public ICommand ResolveLinetypeConflictsCommand { get; set; }
        public ICommand ResolveLineweightConflictsCommand { get; set; }
        public ICommand ResolveTransparencyConflictsCommand { get; set; }
        public ICommand ResolvePlotConflictsCommand { get; set; }
        public ICommand ResolveSelectedConflictCommand { get; set; }
        public ICommand ResolveViewportFreezeConflictsCommand { get; set; }
        public ICommand ResolveViewportColorConflictsCommand { get; set; }
        public ICommand ResolveViewportLinetypeConflictsCommand { get; set; }
        public ICommand ResolveViewportLineweightConflictsCommand { get; set; }
        public ICommand ResolveViewportTransparencyConflictsCommand { get; set; }
        public ICommand ResolveSelectedViewportConflictCommand { get; set; }
        public ICommand NavigateToReviewAndCompareCommand { get; set; }
        public ICommand NavigateToStyleCommand { get; set; }
        public ICommand NavigateToConflictCommand { get; set; }
        public ICommand FindMissingLayersCommand { get; set; }
        public ICommand CompareViewportCommand { get; set; }
        public ICommand NavigateToAdvancedCommand { get; set; }
        public ICommand NavigateToMainCommand { get; set; }
        public ICommand LoadRecentItemCommand { get; set; }
        public ICommand VisretainCommand { get; set; }
        public ICommand LayerReconcileCommand { get; set; }



        #endregion

        #region Constructor
        public MainPageViewModel()
        {
            //When the user navigates to the Main Page, this constructor will automatically run the methods to populate the datagrids.
            //PopulateTargetDrawingList();
            if (SourceDrawingLayers == null)
            {
                SourceDrawingLayers = new ObservableCollection<LayerComparisonLayerModel>();
            }
            if (SourceDrawingViewportLayers == null)
            {
                SourceDrawingViewportLayers = new ObservableCollection<LayerComparisonViewportLayerModel>();
            }
            if (TargetDrawings == null)
            {
                TargetDrawings = new ObservableCollection<TargetDrawingModel>();
            }
            if (DrawingModels == null)
            {
                DrawingModels = new ObservableCollection<DrawingModel>();
            }
            if (TargetDrawingLayers == null)
            {
                TargetDrawingLayers = new ObservableCollection<LayerComparisonLayerModel>();
            }
            if (TargetDrawingViewportLayers == null)
            {
                TargetDrawingViewportLayers = new ObservableCollection<LayerComparisonViewportLayerModel>();
            }
            if (AllConflictLayers == null)
            {
                AllConflictLayers = new ObservableCollection<LayerComparisonLayerModel>();
            }
            if (OnOffConflictLayers == null)
            {
                OnOffConflictLayers = new ObservableCollection<LayerComparisonLayerModel>();
            }
            if (MissingLayers == null)
            {
                MissingLayers = new ObservableCollection<MissingLayerModel>();
            }
            if (SavedComparison == null)
            {
                SavedComparison = new List<string>();
            }
            if (LayerItemList == null)
            {
                LayerItemList = new ObservableCollection<LayerItemViewModel>();
            }
            if (LayerControls == null)
            {
                LayerControls = new ObservableCollection<LayerItemViewModel>();
            }

            this.PropertyChanged += MainPageViewModel_SourceDrawingPropertyChanged;

            void MainPageViewModel_SourceDrawingPropertyChanged(object sender, PropertyChangedEventArgs e)
            {
                switch (e.PropertyName)
                {
                    case nameof(SourceDrawing):
                        ClearComparison();
                        break;
                }
            }

            SaveCommand = new RelayCommand(() => SaveCommandClick());
            NewComparisonCommand = new RelayCommand(() => NewComparisonCommandClick());
            LoadCommand = new RelayCommand(() => LoadCommandClick());
            SaveRecentCommand = new RelayCommand(() => SaveRecentCommandClick());
            BackCommand = new RelayCommand(() => NavigateToStartPage());
            NavigateToAdvancedCommand = new RelayCommand(() => NavigateToAdvancedCommandClick());
            NavigateToMainCommand = new RelayCommand(() => NavigateToMainCommandClick());
            SetSourceDrawingCommand = new RelayCommand(() => SetSourceDrawingClick());
            AddTargetDrawingCommand = new RelayCommand(() => AddTargetDrawingClick());
            RemoveTargetDrawingCommand = new RelayCommand(() => RemoveTargetDrawingClick());
            RemoveViewportConflictCommand = new RelayCommand(() => RemoveSelectedViewportConflictClick());
            OpenSelectedTargetDrawingCommand = new RelayCommand(() => OpenSelectedTargetDrawingClick());
            NavigateToReviewAndCompareCommand = new RelayCommand(() => NavigateToReviewAndCompareCommandClick());
            NavigateToStyleCommand = new RelayCommand(() => NavigateToStyleCommandClick());
            NavigateToConflictCommand = new RelayCommand(() => NavigateToConflictCommandClick());
            OpenSourceDrawingCommand = new RelayCommand(() => OpenSourceDrawingClick());
            ReviewLayerStyleCommand = new RelayCommand(() => GetSourceDrawingLayersClick());
            CompareCommand = new RelayCommand(() => CompareCommandClick());
            CompareViewportCommand = new RelayCommand(() => CompareViewportCommandClick());
            ResolveOnOffConflictsCommand = new RelayCommand(() => ResolveOnOffConflictsClick());
            ResolveFreezeThawConflictsCommand = new RelayCommand(() => ResolveFreezeThawConflictsClick());
            ResolveColorConflictsCommand = new RelayCommand(() => ResolveColorConflictsClick());
            ResolveLinetypeConflictsCommand = new RelayCommand(() => ResolveLinetypeConflictsClick());
            ResolveLineweightConflictsCommand = new RelayCommand(() => ResolveLineweightConflictsClick());
            ResolveTransparencyConflictsCommand = new RelayCommand(() => ResolveTransparencyConflictsClick());
            ResolvePlotConflictsCommand = new RelayCommand(() => ResolvePlotConflictsClick());
            ResolveSelectedConflictCommand = new RelayCommand(() => ResolveSelectedConflictClick());
            ResolveViewportFreezeConflictsCommand = new RelayCommand(() => ResolveViewportFreezeConflictsClick());
            ResolveViewportColorConflictsCommand = new RelayCommand(() => ResolveViewportColorConflictsClick());
            ResolveViewportLinetypeConflictsCommand = new RelayCommand(() => ResolveViewportLinetypeConflictsClick());
            ResolveViewportLineweightConflictsCommand = new RelayCommand(() => ResolveViewportLineweightConflictsClick());
            ResolveViewportTransparencyConflictsCommand = new RelayCommand(() => ResolveViewportTransparencyConflictsClick());
            ResolveSelectedViewportConflictCommand = new RelayCommand(() => ResolveSelectedViewportConflictClick());
            FindMissingLayersCommand = new RelayCommand(() => FindMissingLayersCommandClick());
            VisretainCommand = new RelayCommand(() => VisretainCommandClick());
            LayerReconcileCommand = new RelayCommand(() => ReconcileAllLayersCommandClick());

            BuildRecentComparisons();
        }

        #endregion

        #region Save/Load
        private void SaveCommandClick()
        {
            DataAccess da = new DataAccess();
            if (SavedComparison != null && SavedComparison.Count != 0)
            {
                SavedComparison.Clear();
            }
            SavedComparison.Add(SourceDrawing);
            foreach (var item in TargetDrawings)
            {
                SavedComparison.Add(item.DrawingPath);
            }
            da.SaveComparison(SavedComparison);
        }
        private void LoadCommandClick()
        {
            List<string> missingFiles = new List<string>();
            string source = MessagesAndNotifications.SourceDrawingNotSet;
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = false;
            openFileDialog.Filter = "Layer Comparison (*.lcomp)|*.lcomp";
            Nullable<bool> result = openFileDialog.ShowDialog();
            if (result == true)
            {
                try
                {
                    using (var fileStream = File.OpenRead(openFileDialog.FileName))
                    {
                        using (var streamReader = new StreamReader(fileStream.Name))
                        {
                            source = streamReader.ReadLine() ?? "";
                            if (source != null && source.Contains(".dwg"))
                            {
                                SourceDrawing = source;
                                GetSourceDrawingLayersClick();
                            }
                            else
                            {
                                source = MessagesAndNotifications.SourceDrawingNotSet;
                            }
                        }
                        if (TargetDrawings != null && TargetDrawings.Count != 0)
                        {
                            System.Windows.Forms.DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Do you wish to clear the current target drawings before loading a new comparison?", "Clear target drawings?", System.Windows.Forms.MessageBoxButtons.YesNoCancel, System.Windows.Forms.MessageBoxIcon.Question);
                            if (dialogResult == System.Windows.Forms.DialogResult.Yes)
                            {
                                TargetDrawings.Clear();
                                foreach (var line in File.ReadAllLines(openFileDialog.FileName, System.Text.Encoding.GetEncoding(1250)).Skip(1))
                                {
                                    TargetDrawingModel targetDrawingModel = new TargetDrawingModel();
                                    if (line != null && line.Contains(".dwg"))
                                    {
                                        targetDrawingModel.DrawingPath = line;
                                        TargetDrawings.Add(targetDrawingModel);
                                    }
                                    else
                                    {
                                        MessageBox.Show(MessagesAndNotifications.TargetDrawingsNotSet, "Target Drawings Not Found", MessageBoxButton.OK, MessageBoxImage.Error);
                                    }
                                    IoC.Get<LayerComparisonApplicationViewModel>().GoToPage(ApplicationPage.MainPage);
                                }
                            }
                            else if (dialogResult == System.Windows.Forms.DialogResult.No)
                            {
                                foreach (var line in File.ReadAllLines(openFileDialog.FileName, System.Text.Encoding.GetEncoding(1250)).Skip(1))
                                {
                                    TargetDrawingModel targetDrawingModel = new TargetDrawingModel();
                                    if (line != null && line.Contains(".dwg"))
                                    {
                                        targetDrawingModel.DrawingPath = line;
                                        TargetDrawings.Add(targetDrawingModel);
                                    }
                                    else
                                    {
                                        MessageBox.Show(MessagesAndNotifications.TargetDrawingsNotSet, "Target Drawings Not Found", MessageBoxButton.OK, MessageBoxImage.Error);
                                    }
                                }
                                IoC.Get<LayerComparisonApplicationViewModel>().GoToPage(ApplicationPage.MainPage);
                            }
                            else if (dialogResult == System.Windows.Forms.DialogResult.Cancel)
                            {
                                return;
                            }
                        }
                        else
                        {
                            foreach (var line in File.ReadAllLines(openFileDialog.FileName, System.Text.Encoding.GetEncoding(1250)).Skip(1))
                            {
                                if (File.Exists(line) == true)
                                {
                                    TargetDrawingModel targetDrawingModel = new TargetDrawingModel();
                                    if (line != null && line.Contains(".dwg"))
                                    {
                                        targetDrawingModel.DrawingPath = line;
                                        TargetDrawings.Add(targetDrawingModel);
                                    }
                                    else
                                    {
                                        MessageBox.Show(MessagesAndNotifications.TargetDrawingsNotSet, "Target Drawings Not Found", MessageBoxButton.OK, MessageBoxImage.Error);
                                    }
                                }
                                else
                                {
                                    missingFiles.Add(line);
                                }
                            }
                            if (missingFiles.Count > 0)
                            {
                                MessageBox.Show($"The following drawings could not be located and weren't added\r\n{string.Join(Environment.NewLine, missingFiles)}", "Missing drawings", MessageBoxButton.OK, MessageBoxImage.Warning);
                            }
                            IoC.Get<LayerComparisonApplicationViewModel>().GoToPage(ApplicationPage.MainPage);
                        }
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("A valid source drawing could not be determined. Please make sure the .lcomp file is accessible.", "No source drawing found.", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            //IoC.Get<LayerControllerApplicationViewModel>().GoToPage(LayerControllerApplicationPage.MainPage);
        }
        private void SaveRecentCommandClick()
        {
            DataAccess da = new DataAccess();
            da.SaveRecentComparisonList(SourceDrawing);
        }
        private void BuildRecentComparisons()
        {
            if (File.Exists(FilePaths.RecentComparisonFile))
            {
                var paths = File.ReadLines(FilePaths.RecentComparisonFile);

                RecentItems = new List<RecentItemViewModel>();

                foreach (var result in paths)
                {
                    if (!string.IsNullOrWhiteSpace(result))
                    {
                        RecentItemViewModel recentItem = new RecentItemViewModel();
                        {
                            LoadRecentItemCommand = new RelayCommand(() => LoadRecentItemsCommand(result));
                            recentItem.FileName = Path.GetFileName(result);
                            recentItem.Path = Path.GetDirectoryName(result);
                            recentItem.Command = LoadRecentItemCommand;
                            RecentItems.Add(recentItem);
                        };
                    }
                }
                RecentItems.Reverse();
            }
        }
        #endregion

        #region Navigation
        /// <summary>
        /// Sets the current page to the Start Page.
        /// </summary>
        private void NavigateToStartPage()
        {
            if (SourceDrawing != MessagesAndNotifications.SourceDrawingNotSet)
            {
                System.Windows.Forms.DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Returning to the start page will clear out your current comparison. Continue?", "Clear comparison?", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);
                switch (dialogResult)
                {
                    case System.Windows.Forms.DialogResult.Yes:
                        SourceDrawing = MessagesAndNotifications.SourceDrawingNotSet;
                        ClearComparison();
                        BuildRecentComparisons();
                        IoC.Get<LayerComparisonApplicationViewModel>().GoToPage(ApplicationPage.StartPage);
                        break;
                    case System.Windows.Forms.DialogResult.No:
                        break;
                    default:
                        break;
                }
            }
            else
            {
                SourceDrawing = MessagesAndNotifications.SourceDrawingNotSet;
                ClearComparison();
                BuildRecentComparisons();
                IoC.Get<LayerComparisonApplicationViewModel>().GoToPage(ApplicationPage.StartPage);
            }
        }
        private void NewComparisonCommandClick()
        {
            IoC.Get<LayerComparisonApplicationViewModel>().GoToPage(ApplicationPage.MainPage);
        }
        private void NavigateToAdvancedCommandClick()
        {
            IoC.Get<LayerComparisonApplicationViewModel>().GoToPage(ApplicationPage.AdvancedPage);
        }
        private void NavigateToMainCommandClick()
        {
            IoC.Get<LayerComparisonApplicationViewModel>().GoToPage(ApplicationPage.MainPage);
        }
        private void NavigateToReviewAndCompareCommandClick()
        {
            //MainPageTabIndex = 1;
        }
        private void NavigateToStyleCommandClick()
        {
            MainPageTabIndex = 0;
        }
        private void NavigateToConflictCommandClick()
        {
            MainPageTabIndex = 2;
        }
        #endregion

        #region Source Drawing Methods

        /// <summary>
        /// Call this method when you need to build a list of target drawings for a style to populate to the <see cref="MainPageViewModel.SourceDrawing"/> property.
        /// </summary>
        private void SetSourceDrawingClick()
        {
            if (SourceDrawing != null && SourceDrawing != MessagesAndNotifications.PleaseLoadSourceDrawing)
            {
                System.Windows.Forms.DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Setting a new source drawing will clear out your current comparison. Continue?", "Clear comparison?", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);
                switch (dialogResult)
                {
                    case System.Windows.Forms.DialogResult.Yes:
                        ClearComparison();
                        DataAccess da = new DataAccess();
                        SourceDrawing = da.GetSource();
                        GetSourceDrawingLayersClick();
                        GetSourceDrawingViewportLayersClick();
                        //BuildLayerItems();
                        break;
                    case System.Windows.Forms.DialogResult.No:
                        break;
                    default:
                        break;
                }
            }
            else
            {
                DataAccess da = new DataAccess();
                SourceDrawing = da.GetSource();
                GetSourceDrawingLayersClick();
                GetSourceDrawingViewportLayersClick();
            }
        }
        public void SetSourceDrawing(string source)
        {
            SourceDrawing = source;
        }

        /// <summary>
        /// Opens the target drawing selected by the user.
        /// </summary>
        private void OpenSourceDrawingClick()
        {
            if (SourceDrawing != null)
                try
                {
                    DatabaseHelpers dh = new DatabaseHelpers();
                    dh.OpenDrawing(SourceDrawing);
                }
                catch (System.Exception)
                {
                    MessageBox.Show("The source drawing could not be determined.");
                }
        }

        /// <summary>
        /// Builds the SourceDrawingLayers collection.
        /// </summary>
        public void GetSourceDrawingLayersClick()
        {
            if (!SourceDrawing.Contains(MessagesAndNotifications.SourceDrawingNotSet) && SourceDrawing != null)
            {
                string s = SourceDrawing;
                DataAccess da = new DataAccess();
                SourceDrawingLayers = da.BuildDrawingLayersCollection(s);
            }
            else
            {
                MessageBox.Show("A source drawing could not be found. Please use the 'Set Source' button to set the source drawing for this comparison.");
                return;
            }
        }

        private void GetSourceDrawingViewportLayersClick()
        {
            if (!SourceDrawing.Contains(MessagesAndNotifications.SourceDrawingNotSet) && SourceDrawing != null)
            {
                string s = SourceDrawing;
                DataAccess da = new DataAccess();
                SourceDrawingViewportLayers = da.BuildDrawingViewportLayersCollection(s);
            }
            else
            {
                MessageBox.Show("A source drawing could not be found. Please use the 'Set Source' button to set the source drawing for this comparison.");
                return;
            }
        }
        #endregion

        #region Target Drawing Methods

        /// <summary>
        /// Adds target drawings to an ObservableCollection of type <see cref="TargetDrawingModel"/>.
        /// </summary>
        private void AddTargetDrawingClick()
        {
            DataAccess da = new DataAccess();
            if (!SourceDrawing.Contains(MessagesAndNotifications.SourceDrawingNotSet) && SourceDrawing != null)
            {
                foreach (var result in da.AddtoTargetDrawingList())
                {
                    if (result != null && result.DrawingPath != "nothing added" && TargetDrawings.Any(p => p.DrawingPath == result.DrawingPath) == false)
                    {
                        TargetDrawings.Add(result);
                    }
                }
            }
            else
            {
                MessageBox.Show(MessagesAndNotifications.AddASourceBeforeAddingTargetsReminder, "No source drawing set", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            //this is a test.
        }

        /// <summary>
        /// Removes a target drawing from the list of target drawings.
        /// </summary>
        private void RemoveTargetDrawingClick()
        {
            TargetDrawings.Remove(SelectedTargetDrawing);
        }

        /// <summary>
        /// Opens the target drawing selected by the user.
        /// </summary>
        private void OpenSelectedTargetDrawingClick()
        {
            if (SelectedTargetDrawing != null)
                try
                {
                    DatabaseHelpers dh = new DatabaseHelpers();
                    dh.OpenDrawing(SelectedTargetDrawing.DrawingPath);
                }
                catch (System.Exception)
                {
                    MessageBox.Show("A valid target drawing could not be determined.");
                }
        }

        #endregion

        #region Comparison

        /// <summary>
        /// Populates the TargetDrawingLayers collection.
        /// </summary>        /// <summary>
        /// Runs the master comparison that compares every target drawing layer from every target drawing to every source drawing layer.
        /// </summary>
        private void CompareCommandClick()
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();


            //If the comparison has already been ran, we only want to clear out the conflicts, not the TargetDrawingLayers
            if (ComparisonRan == true)
            {
                PartialClearComparison();
            }
            ComparisonRan = true;

            //Get the layers and viewport layers for the comparison.
            GetTargetDrawingLayers();

            //Ensure the conflict layers are cleared out for a new comparison.
            if (AllConflictLayers != null && AllConflictLayers.Count != 0)
            {
                AllConflictLayers.Clear();
            }

            //This will create a large collection of every layer that is mis-matched from SourceDrawingLayers.
            var conflict = TargetDrawingLayers.Where(i => !SourceDrawingLayers.Contains(i)).ToList();
            ObservableCollection<LayerComparisonLayerModel> allConflicts = new ObservableCollection<LayerComparisonLayerModel>(conflict);
            AllConflictLayers = allConflicts;

            //these methods will run LINQ queries on the collections generated above to find the exact mis-matches.
            GetOnOffConflicts();
            GetFreezeConflicts();
            GetColorConflicts();
            GetLinetypeConflicts();
            GetLineweightConflicts();
            GetTransparencyConflicts();
            GetPlotConflicts();
            GetExtraLayerConflicts();
            GetMissingLayerConflicts();

            //Update the UI with the number of each type of conflict.
            OnOffConflictsCount = OnOffConflictLayers.Count.ToString() + MessagesAndNotifications.OnOffCountMessage;
            FreezeConflictsCount = FreezeConflictLayers.Count.ToString() + MessagesAndNotifications.FreezeCountMessage;
            ColorConflictsCount = ColorConflictLayers.Count.ToString() + MessagesAndNotifications.ColorCountMessage;
            LinetypeConflictsCount = LinetypeConflictLayers.Count.ToString() + MessagesAndNotifications.LinetypeCountMessage;
            LineweightConflictsCount = LineweightConflictLayers.Count.ToString() + MessagesAndNotifications.LineweightCountMessage;
            TransparencyConflictsCount = TransparencyConflictLayers.Count.ToString() + MessagesAndNotifications.TransparencyCountMessage;
            PlotConflictsCount = PlotConflictLayers.Count.ToString() + MessagesAndNotifications.PlotCountMessage;
            ExtraLayersCount = ExtraLayers.Count.ToString() + MessagesAndNotifications.ExtraLayersCountMessage;
            MissingLayersCount = MissingLayers.Count.ToString() + MessagesAndNotifications.MissingLayersCountMessage; TotalConflictsCount = AllConflictLayers.Count.ToString() + MessagesAndNotifications.AllConflictsCountMessage;

            //Navigate to the results
            MainPageTabIndex = 2;

            stopwatch.Stop();
            TimeSpan timeSpan = stopwatch.Elapsed;

            // Format and display the TimeSpan value. 
            ElapsedTime = "Comparison Run Time: " + String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
               timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds,
               timeSpan.Milliseconds / 10);
        }
        private void GetTargetDrawingLayers()
        {
            if (TargetDrawingLayers != null && TargetDrawingLayers.Count != 0)
            {
                TargetDrawingLayers.Clear();
            }
            DataAccess da = new DataAccess();
            foreach (TargetDrawingModel targetDrawingModel in TargetDrawings)
            {
                foreach (var result in da.BuildDrawingLayersCollection(targetDrawingModel.DrawingPath))
                {
                    TargetDrawingLayers.Add(result);
                }
            }
        }
        private void GetTargetDrawingViewportLayers()
        {
            if (TargetDrawingViewportLayers != null && TargetDrawingViewportLayers.Count != 0)
            {
                TargetDrawingViewportLayers.Clear();
            }
            DataAccess da = new DataAccess();
            foreach (TargetDrawingModel targetDrawingModel in TargetDrawings)
            {
                foreach (var result in da.BuildDrawingViewportLayersCollection(targetDrawingModel.DrawingPath))
                {
                    TargetDrawingViewportLayers.Add(result);
                }
            }
        }
        private void CompareViewportCommandClick()
        {
            //If the comparison has already been ran, we only want to clear out the conflicts, not the TargetDrawingLayers
            if (ComparisonRan == true)
            {
                PartialClearComparison();
            }
            ComparisonRan = true;
            if (SourceDrawingViewportLayers != null || SourceDrawingViewportLayers.Count != 0)
            {
                SourceDrawingViewportLayers.Clear();
            }
            if (TargetDrawingViewportLayers != null || TargetDrawingViewportLayers.Count != 0)
            {
                TargetDrawingViewportLayers.Clear();
            }
            GetSourceDrawingViewportLayersClick();
            GetTargetDrawingViewportLayers();
            GetViewportFreezeConflicts();
            GetViewportColorConflicts();
            GetViewportLinetypeConflicts();
            GetViewportLineWeightConflicts();
            GetViewportTransparencyConflicts();

            ViewportFreezeConflictsCount = ViewportFreezeConflictLayers.Count.ToString() + MessagesAndNotifications.ViewportFreezeCountMessage;
            ViewportColorConflictsCount = ViewportColorConflictLayers.Count.ToString() + MessagesAndNotifications.ViewportColorCountMessage;
            ViewportLinetypeConflictsCount = ViewportLinetypeConflictLayers.Count.ToString() + MessagesAndNotifications.ViewportLinetypeCountMessage;
            ViewportLineweightConflictsCount = ViewportLineweightConflictLayers.Count.ToString() + MessagesAndNotifications.ViewportLineweightCountMessage;
            ViewportTransparencyConflictsCount = ViewportTransparencyConflictLayers.Count.ToString() + MessagesAndNotifications.ViewportTransparencyCountMessage;


            SelectedAdvancedTabIndex = 0;
        }
        private void FindMissingLayersCommandClick()
        {
            ReturnMissingLayers(SourceDrawingLayers);
            MissingLayersCount = MissingLayers.Count.ToString() + MessagesAndNotifications.MissingLayersCountMessage;
        }

        #endregion

        #region Basic Layer Conflict Methods
        private void GetOnOffConflicts()
        {
            DataAccess da = new DataAccess();
            if (OnOffConflictLayers != null && OnOffConflictLayers.Count != 0)
            {
                OnOffConflictLayers.Clear();
            }
            var onOffQuery = from target in TargetDrawingLayers.AsParallel()
                             from source in SourceDrawingLayers.AsParallel()
                             where target.Name == source.Name && target.OnOff != source.OnOff
                             select target;

            ObservableCollection<LayerComparisonLayerModel> q = new ObservableCollection<LayerComparisonLayerModel>(onOffQuery);
            foreach (var result in q)
            {
                foreach (var sourceLayer in SourceDrawingLayers)
                {
                    if (result.Name == sourceLayer.Name)
                    {
                        result.SourceOnOff = sourceLayer.OnOff;
                    }
                }
            }
            OnOffConflictLayers = q;
        }
        private void GetFreezeConflicts()
        {
            DataAccess da = new DataAccess();
            if (FreezeConflictLayers != null && FreezeConflictLayers.Count != 0)
            {
                FreezeConflictLayers.Clear();
            }

            var freezeQuery = from target in TargetDrawingLayers.AsParallel()
                              from source in SourceDrawingLayers.AsParallel()
                              where target.Name == source.Name && target.Freeze != source.Freeze
                              select target;

            ObservableCollection<LayerComparisonLayerModel> q = new ObservableCollection<LayerComparisonLayerModel>(freezeQuery);
            foreach (var result in q)
            {
                foreach (var sourceLayer in SourceDrawingLayers)
                {
                    if (result.Name == sourceLayer.Name)
                    {
                        result.SourceFreeze = sourceLayer.Freeze;
                    }
                }
            }
            FreezeConflictLayers = q;
        }
        private void GetColorConflicts()
        {
            DataAccess da = new DataAccess();
            if (ColorConflictLayers != null && ColorConflictLayers.Count != 0)
            {
                ColorConflictLayers.Clear();
            }

            var colorQuery = from target in TargetDrawingLayers.AsParallel()
                             from source in SourceDrawingLayers.AsParallel()
                             where target.Name == source.Name && target.Color != source.Color
                             select target;

            ObservableCollection<LayerComparisonLayerModel> q = new ObservableCollection<LayerComparisonLayerModel>(colorQuery);
            foreach (var result in q)
            {
                foreach (var sourceLayer in SourceDrawingLayers)
                {
                    if (result.Name == sourceLayer.Name)
                    {
                        result.SourceColor = sourceLayer.Color;
                    }
                }
            }
            ColorConflictLayers = q;
        }
        private void GetLinetypeConflicts()
        {
            DataAccess da = new DataAccess();
            if (LinetypeConflictLayers != null && LinetypeConflictLayers.Count != 0)
            {
                LinetypeConflictLayers.Clear();
            }

            var linetypeQuery = from target in TargetDrawingLayers.AsParallel()
                                from source in SourceDrawingLayers.AsParallel()
                                where target.Name == source.Name && target.Linetype != source.Linetype
                                select target;

            ObservableCollection<LayerComparisonLayerModel> q = new ObservableCollection<LayerComparisonLayerModel>(linetypeQuery);
            foreach (var result in q)
            {
                foreach (var sourceLayer in SourceDrawingLayers)
                {
                    if (result.Name == sourceLayer.Name)
                    {
                        result.SourceLinetype = sourceLayer.Linetype;
                    }
                }
            }
            LinetypeConflictLayers = q;
        }
        private void GetLineweightConflicts()
        {
            DataAccess da = new DataAccess();
            if (LineweightConflictLayers != null && LineweightConflictLayers.Count != 0)
            {
                LineweightConflictLayers.Clear();
            }

            var lineweightQuery = from target in TargetDrawingLayers.AsParallel()
                                  from source in SourceDrawingLayers.AsParallel()
                                  where target.Name == source.Name && target.Lineweight != source.Lineweight
                                  select target;

            ObservableCollection<LayerComparisonLayerModel> q = new ObservableCollection<LayerComparisonLayerModel>(lineweightQuery);
            foreach (var result in q)
            {
                foreach (var sourceLayer in SourceDrawingLayers)
                {
                    if (result.Name == sourceLayer.Name)
                    {
                        result.SourceLineweight = sourceLayer.Lineweight;
                    }
                }
            }
            LineweightConflictLayers = q;
        }
        private void GetTransparencyConflicts()
        {
            DataAccess da = new DataAccess();
            if (TransparencyConflictLayers != null && TransparencyConflictLayers.Count != 0)
            {
                TransparencyConflictLayers.Clear();
            }

            var transparencyQuery = from target in TargetDrawingLayers.AsParallel()
                                    from source in SourceDrawingLayers.AsParallel()
                                    where target.Name == source.Name && target.Transparency != source.Transparency
                                    select target;

            ObservableCollection<LayerComparisonLayerModel> q = new ObservableCollection<LayerComparisonLayerModel>(transparencyQuery);
            foreach (var result in q)
            {
                foreach (var sourceLayer in SourceDrawingLayers)
                {
                    if (result.Name == sourceLayer.Name)
                    {
                        result.SourceTransparency = sourceLayer.Transparency;
                    }
                }
            }
            TransparencyConflictLayers = q;
        }
        private void GetPlotConflicts()
        {
            DataAccess da = new DataAccess();
            if (PlotConflictLayers != null && PlotConflictLayers.Count != 0)
            {
                PlotConflictLayers.Clear();
            }

            var plotQuery = from target in TargetDrawingLayers.AsParallel()
                            from source in SourceDrawingLayers.AsParallel()
                            where target.Name == source.Name && target.Plot != source.Plot
                            select target;

            ObservableCollection<LayerComparisonLayerModel> q = new ObservableCollection<LayerComparisonLayerModel>(plotQuery);
            foreach (var result in q)
            {
                foreach (var sourceLayer in SourceDrawingLayers)
                {
                    if (result.Name == sourceLayer.Name)
                    {
                        result.SourcePlot = sourceLayer.Plot;
                    }
                }
            }
            PlotConflictLayers = q;
        }
        private void GetExtraLayerConflicts()
        {
            DataAccess da = new DataAccess();
            if (ExtraLayers != null && ExtraLayers.Count != 0)
            {
                ExtraLayers.Clear();
            }
            var extraLayerQuery = TargetDrawingLayers.Where(p => !SourceDrawingLayers.Any(p2 => p2.Name == p.Name));
            ObservableCollection<LayerComparisonLayerModel> q = new ObservableCollection<LayerComparisonLayerModel>(extraLayerQuery);
            ExtraLayers = q;
        }
        public void GetMissingLayerConflicts()
        {
            if (MissingLayers != null && MissingLayers.Count != 0)
            {
                MissingLayers.Clear();
            }

            if (TargetDrawingLayers == null || TargetDrawingLayers.Count == 0)
            {
                System.Windows.Forms.DialogResult dialogResult = System.Windows.Forms.MessageBox.Show(MessagesAndNotifications.MissingLayersCantRun, "Get target drawing layers?", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);
                switch (dialogResult)
                {
                    case System.Windows.Forms.DialogResult.Yes:
                        GetTargetDrawingLayers();
                        ReturnMissingLayers(SourceDrawingLayers);
                        SelectedAdvancedTabIndex = 1;
                        break;


                    case System.Windows.Forms.DialogResult.No:
                        MessageBox.Show(MessagesAndNotifications.MissingLayersCancelled);
                        break;
                    default:
                        break;
                }
            }
            else
            {
                ReturnMissingLayers(SourceDrawingLayers);
                //ReturnMissingLayers(SourceDrawingLayers);
                SelectedAdvancedTabIndex = 1;
            }
        }
        #endregion

        #region Viewport Layer Conflict Methods
        private void GetViewportFreezeConflicts()
        {
            DataAccess da = new DataAccess();
            if (ViewportFreezeConflictLayers != null && ViewportFreezeConflictLayers.Count != 0)
            {
                ViewportFreezeConflictLayers.Clear();
            }

            var viewportFreezeQuery = from target in TargetDrawingViewportLayers.AsParallel()
                                      from source in SourceDrawingViewportLayers.AsParallel()
                                      where target.ViewportLayer == source.ViewportLayer && target.Name == source.Name && target.ViewportFreeze != source.ViewportFreeze
                                      select target;

            if (viewportFreezeQuery != null)
            {
                ObservableCollection<LayerComparisonViewportLayerModel> q = new ObservableCollection<LayerComparisonViewportLayerModel>(viewportFreezeQuery);
                foreach (var result in q)
                {
                    foreach (var sourceLayer in SourceDrawingViewportLayers)
                    {
                        if (result.Name == sourceLayer.Name)
                        {
                            result.SourceViewportFreeze = sourceLayer.ViewportFreeze;
                        }
                    }
                }
                ViewportFreezeConflictLayers = q;
            }
        }
        private void GetViewportColorConflicts()
        {
            DataAccess da = new DataAccess();
            if (ViewportColorConflictLayers != null && ViewportColorConflictLayers.Count != 0)
            {
                ViewportColorConflictLayers.Clear();
            }

            var viewportColorQuery = from target in TargetDrawingViewportLayers.AsParallel()
                                     from source in SourceDrawingViewportLayers.AsParallel()
                                     where target.Name == source.Name && target.ViewportLayer == source.ViewportLayer && target.ViewportColor != source.ViewportColor
                                     select target;

            if (viewportColorQuery != null)
            {
                ObservableCollection<LayerComparisonViewportLayerModel> q = new ObservableCollection<LayerComparisonViewportLayerModel>(viewportColorQuery);
                foreach (var result in q)
                {
                    foreach (var sourceLayer in SourceDrawingViewportLayers)
                    {
                        if (result.Name == sourceLayer.Name)
                        {
                            result.SourceViewportColor = sourceLayer.ViewportColor;
                        }
                    }
                }
                ViewportColorConflictLayers = q;
            }
        }
        private void GetViewportLinetypeConflicts()
        {
            DataAccess da = new DataAccess();
            if (ViewportLinetypeConflictLayers != null && ViewportLinetypeConflictLayers.Count != 0)
            {
                ViewportLinetypeConflictLayers.Clear();
            }

            var viewportLinetypeQuery = from target in TargetDrawingViewportLayers.AsParallel()
                                        from source in SourceDrawingViewportLayers.AsParallel()
                                        where target.Name == source.Name && target.ViewportLayer == source.ViewportLayer && target.ViewportLinetype != source.ViewportLinetype
                                        select target;

            if (viewportLinetypeQuery != null)
            {
                ObservableCollection<LayerComparisonViewportLayerModel> q = new ObservableCollection<LayerComparisonViewportLayerModel>(viewportLinetypeQuery);
                foreach (var result in q)
                {
                    foreach (var sourceLayer in SourceDrawingViewportLayers)
                    {
                        if (result.Name == sourceLayer.Name)
                        {
                            result.SourceViewportLinetype = sourceLayer.ViewportLinetype;
                        }
                    }
                }
                ViewportLinetypeConflictLayers = q;
            }
        }
        private void GetViewportLineWeightConflicts()
        {
            DataAccess da = new DataAccess();
            if (ViewportLineweightConflictLayers != null && ViewportLineweightConflictLayers.Count != 0)
            {
                ViewportLineweightConflictLayers.Clear();
            }

            var viewportLineweightQuery = from target in TargetDrawingViewportLayers.AsParallel()
                                          from source in SourceDrawingViewportLayers.AsParallel()
                                          where target.Name == source.Name && target.ViewportLayer == source.ViewportLayer && target.ViewportLineweight != source.ViewportLineweight
                                          select target;

            if (viewportLineweightQuery != null)
            {
                ObservableCollection<LayerComparisonViewportLayerModel> q = new ObservableCollection<LayerComparisonViewportLayerModel>(viewportLineweightQuery);
                foreach (var result in q)
                {
                    foreach (var sourceLayer in SourceDrawingViewportLayers)
                    {
                        if (result.Name == sourceLayer.Name)
                        {
                            result.SourceViewportLineweight = sourceLayer.ViewportLineweight;
                        }
                    }
                }
                ViewportLineweightConflictLayers = q;
            }
        }
        private void GetViewportTransparencyConflicts()
        {
            DataAccess da = new DataAccess();
            if (ViewportTransparencyConflictLayers != null && ViewportTransparencyConflictLayers.Count != 0)
            {
                ViewportTransparencyConflictLayers.Clear();
            }

            var viewportLineweightQuery = from target in TargetDrawingViewportLayers.AsParallel()
                                          from source in SourceDrawingViewportLayers.AsParallel()
                                          where target.Name == source.Name && target.ViewportLayer == source.ViewportLayer && target.ViewportTransparency != source.ViewportTransparency
                                          select target;

            if (viewportLineweightQuery != null)
            {
                ObservableCollection<LayerComparisonViewportLayerModel> q = new ObservableCollection<LayerComparisonViewportLayerModel>(viewportLineweightQuery);
                foreach (var result in q)
                {
                    foreach (var sourceLayer in SourceDrawingViewportLayers)
                    {
                        if (result.Name == sourceLayer.Name)
                        {
                            result.SourceViewportTransparency = sourceLayer.ViewportTransparency;
                        }
                    }
                }
                ViewportTransparencyConflictLayers = q;
            }
        }
        #endregion

        #region Resolve Basic Layer Conflicts
        private void ResolveOnOffConflictsClick()
        {
            if (OnOffConflictLayers != null && OnOffConflictLayers.Count != 0)
            {
                DataAccess da = new DataAccess();

                var conflictsAndSources = from conflict in OnOffConflictLayers
                                          from source in SourceDrawingLayers
                                          where conflict.Name == source.Name && conflict.OnOff != source.OnOff
                                          select new { Conflict = conflict, Source = source };

                var Groups = conflictsAndSources.GroupBy(x => x.Conflict.DrawingPath, x => x.Source);

                foreach (var group in Groups)
                {
                    da.FixGroupOfLayerConflicts(group.Key, "OnOff", group);
                }
                GetTargetDrawingLayers();
                GetOnOffConflicts();
                //OnOffConflictLayers.Clear();
                OnOffConflictsCount = OnOffConflictLayers.Count.ToString() + " Total On/Off conflicts with source drawing detected.";
                #region Original
                //foreach (var conflictAndSource in conflictsAndSources)
                //    da.FixLayerConflict(conflictAndSource.Conflict.DrawingPath, conflictAndSource.Conflict.Name, "OnOff", conflictAndSource.Source.OnOff);
                #endregion
            }
        }
        private void ResolveFreezeThawConflictsClick()
        {
            if (FreezeConflictLayers != null && FreezeConflictLayers.Count != 0)
            {
                DataAccess da = new DataAccess();

                var conflictsAndSources = from conflict in FreezeConflictLayers
                                          from source in SourceDrawingLayers
                                          where conflict.Name == source.Name && conflict.Freeze != source.Freeze
                                          select new { Conflict = conflict, Source = source };

                var Groups = conflictsAndSources.GroupBy(x => x.Conflict.DrawingPath, x => x.Source);

                foreach (var group in Groups)
                {
                    da.FixGroupOfLayerConflicts(group.Key, "FreezeThaw", group);
                }
                GetTargetDrawingLayers();
                GetFreezeConflicts();
                //FreezeConflictLayers.Clear();
                FreezeConflictsCount = FreezeConflictLayers.Count.ToString() + MessagesAndNotifications.FreezeCountMessage;
                #region Original
                //foreach (var conflictAndSource in conflictsAndSources)
                //    da.FixLayerConflict(conflictAndSource.Conflict.DrawingPath, conflictAndSource.Conflict.Name, "FreezeThaw", conflictAndSource.Source.Freeze);
                #endregion

            }
        }
        private void ResolveColorConflictsClick()
        {
            if (ColorConflictLayers != null && ColorConflictLayers.Count != 0)
            {
                DataAccess da = new DataAccess();

                var conflictsAndSources = from conflict in ColorConflictLayers
                                          from source in SourceDrawingLayers
                                          where conflict.Name == source.Name && conflict.Color != source.Color
                                          select new { Conflict = conflict, Source = source };


                var Groups = conflictsAndSources.GroupBy(x => x.Conflict.DrawingPath, x => x.Source);

                foreach (var group in Groups)
                {
                    da.FixGroupOfLayerConflicts(group.Key, "Color", group);
                }
                GetTargetDrawingLayers();
                GetColorConflicts();
                //ColorConflictLayers.Clear();
                ColorConflictsCount = ColorConflictLayers.Count.ToString() + MessagesAndNotifications.ColorCountMessage;
                #region Original
                //foreach (var conflictAndSource in conflictsAndSources)
                //    da.FixLayerConflict(conflictAndSource.Conflict.DrawingPath, conflictAndSource.Conflict.Name, "Color", conflictAndSource.Source.Color);
                #endregion
            }
        }
        private void ResolveLinetypeConflictsClick()
        {
            if (LinetypeConflictLayers != null && LinetypeConflictLayers.Count != 0)
            {
                DataAccess da = new DataAccess();

                var conflictsAndSources = from conflict in LinetypeConflictLayers
                                          from source in SourceDrawingLayers
                                          where conflict.Name == source.Name && conflict.Linetype != source.Linetype
                                          select new { Conflict = conflict, Source = source };


                var Groups = conflictsAndSources.GroupBy(x => x.Conflict.DrawingPath, x => x.Source);

                foreach (var group in Groups)
                {
                    da.FixGroupOfLayerConflicts(group.Key, "Linetype", group);
                }
                GetTargetDrawingLayers();
                GetLinetypeConflicts();
                //LinetypeConflictLayers.Clear();
                LinetypeConflictsCount = LinetypeConflictLayers.Count.ToString() + MessagesAndNotifications.LinetypeCountMessage;
                #region Original
                //foreach (var conflictAndSource in conflictsAndSources)
                //    da.FixLayerConflict(conflictAndSource.Conflict.DrawingPath, conflictAndSource.Conflict.Name, "Linetype", conflictAndSource.Source.Linetype);
                #endregion
            }
        }
        private void ResolveLineweightConflictsClick()
        {
            if (LineweightConflictLayers != null && LineweightConflictLayers.Count != 0)
            {
                DataAccess da = new DataAccess();

                var conflictsAndSources = from conflict in LineweightConflictLayers
                                          from source in SourceDrawingLayers
                                          where conflict.Name == source.Name && conflict.Lineweight != source.Lineweight
                                          select new { Conflict = conflict, Source = source };
                var Groups = conflictsAndSources.GroupBy(x => x.Conflict.DrawingPath, x => x.Source);

                foreach (var group in Groups)
                {
                    da.FixGroupOfLayerConflicts(group.Key, "Lineweight", group);
                }
                GetTargetDrawingLayers();
                GetLineweightConflicts();
                //LinetypeConflictLayers.Clear();
                LineweightConflictsCount = LineweightConflictLayers.Count.ToString() + MessagesAndNotifications.LineweightCountMessage;
                #region Original
                //foreach (var conflictAndSource in conflictsAndSources)
                //    da.FixLayerConflict(conflictAndSource.Conflict.DrawingPath, conflictAndSource.Conflict.Name, "Lineweight", conflictAndSource.Source.Lineweight); 
                #endregion
            }
        }
        private void ResolveTransparencyConflictsClick()
        {
            if (TransparencyConflictLayers != null && TransparencyConflictLayers.Count != 0)
            {
                DataAccess da = new DataAccess();

                var conflictsAndSources = from conflict in TransparencyConflictLayers
                                          from source in SourceDrawingLayers
                                          where conflict.Name == source.Name && conflict.Transparency != source.Transparency
                                          select new { Conflict = conflict, Source = source };

                var Groups = conflictsAndSources.GroupBy(x => x.Conflict.DrawingPath, x => x.Source);

                foreach (var group in Groups)
                {
                    da.FixGroupOfLayerConflicts(group.Key, "Transparency", group);
                }
                GetTargetDrawingLayers();
                GetTransparencyConflicts();
                //TransparencyConflictLayers.Clear();
                TransparencyConflictsCount = TransparencyConflictLayers.Count.ToString() + MessagesAndNotifications.TransparencyCountMessage;
                #region Original
                //foreach (var conflictAndSource in conflictsAndSources)
                //    da.FixLayerConflict(conflictAndSource.Conflict.DrawingPath, conflictAndSource.Conflict.Name, "Transparency", conflictAndSource.Source.Transparency); 
                #endregion
            }
        }
        private void ResolvePlotConflictsClick()
        {
            if (PlotConflictLayers != null && PlotConflictLayers.Count != 0)
            {
                DataAccess da = new DataAccess();

                var conflictsAndSources = from conflict in PlotConflictLayers
                                          from source in SourceDrawingLayers
                                          where conflict.Name == source.Name && conflict.Plot != source.Plot
                                          select new { Conflict = conflict, Source = source };

                var Groups = conflictsAndSources.GroupBy(x => x.Conflict.DrawingPath, x => x.Source);

                foreach (var group in Groups)
                {
                    da.FixGroupOfLayerConflicts(group.Key, "Plot", group);
                }
                GetTargetDrawingLayers();
                GetPlotConflicts();
                //PlotConflictLayers.Clear();
                PlotConflictsCount = PlotConflictLayers.Count.ToString() + MessagesAndNotifications.PlotCountMessage;
                #region Original
                //foreach (var conflictAndSource in conflictsAndSources)
                //    da.FixLayerConflict(conflictAndSource.Conflict.DrawingPath, conflictAndSource.Conflict.Name, "Plot", conflictAndSource.Source.Plot);
                #endregion
            }
        }
        #endregion

        #region Resolve Selected Basic Layer Conflicts
        private void ResolveSelectedConflictClick()
        {
            switch (SelectedConflictTabIndex)
            {
                case (int)SelectedConflictTab.OnOff:
                    if (SelectedConflict != null)
                    {
                        DataAccess da = new DataAccess();

                        var conflictsAndSources =
                                                  from source in SourceDrawingLayers
                                                  where SelectedConflict.Name == source.Name && SelectedConflict.OnOff != source.OnOff
                                                  select new { Conflict = SelectedConflict, Source = source };

                        foreach (var conflictAndSource in conflictsAndSources)
                            da.FixLayerConflict(conflictAndSource.Conflict.DrawingPath, conflictAndSource.Conflict.Name, "OnOff", conflictAndSource.Source.OnOff);

                        GetTargetDrawingLayers();
                        GetOnOffConflicts();
                        //OnOffConflictLayers.Remove(SelectedConflict);
                        OnOffConflictsCount = OnOffConflictLayers.Count.ToString() + MessagesAndNotifications.OnOffCountMessage;
                    }
                    break;
                case (int)SelectedConflictTab.Freeze:
                    if (SelectedConflict != null)
                    {
                        {
                            DataAccess da = new DataAccess();

                            var conflictsAndSources =
                                                      from source in SourceDrawingLayers
                                                      where SelectedConflict.Name == source.Name && SelectedConflict.Freeze != source.Freeze
                                                      select new { Conflict = SelectedConflict, Source = source };

                            foreach (var conflictAndSource in conflictsAndSources)
                                da.FixLayerConflict(conflictAndSource.Conflict.DrawingPath, conflictAndSource.Conflict.Name, "FreezeThaw", conflictAndSource.Source.Freeze);

                            GetTargetDrawingLayers();
                            GetFreezeConflicts();
                            //FreezeConflictLayers.Remove(SelectedConflict);
                            FreezeConflictsCount = FreezeConflictLayers.Count.ToString() + MessagesAndNotifications.FreezeCountMessage;
                        }
                    }
                    break;
                case (int)SelectedConflictTab.Color:
                    if (SelectedConflict != null)
                    {
                        {
                            DataAccess da = new DataAccess();

                            var conflictsAndSources =
                                                      from source in SourceDrawingLayers
                                                      where SelectedConflict.Name == source.Name && SelectedConflict.Color != source.Color
                                                      select new { Conflict = SelectedConflict, Source = source };

                            foreach (var conflictAndSource in conflictsAndSources)
                                da.FixLayerConflict(conflictAndSource.Conflict.DrawingPath, conflictAndSource.Conflict.Name, "Color", conflictAndSource.Source.Color);

                            GetTargetDrawingLayers();
                            GetColorConflicts();
                            //ColorConflictLayers.Remove(SelectedConflict);
                            ColorConflictsCount = ColorConflictLayers.Count.ToString() + MessagesAndNotifications.ColorCountMessage;
                        }
                    }
                    break;
                case (int)SelectedConflictTab.Linetype:
                    if (SelectedConflict != null)
                    {
                        {
                            DataAccess da = new DataAccess();

                            var conflictsAndSources =
                                                      from source in SourceDrawingLayers
                                                      where SelectedConflict.Name == source.Name && SelectedConflict.Linetype != source.Linetype
                                                      select new { Conflict = SelectedConflict, Source = source };

                            foreach (var conflictAndSource in conflictsAndSources)
                                da.FixLayerConflict(conflictAndSource.Conflict.DrawingPath, conflictAndSource.Conflict.Name, "Linetype", conflictAndSource.Source.Linetype);

                            GetTargetDrawingLayers();
                            GetLinetypeConflicts();
                            //LinetypeConflictLayers.Remove(SelectedConflict);
                            LinetypeConflictsCount = LinetypeConflictLayers.Count.ToString() + MessagesAndNotifications.LinetypeCountMessage;
                        }
                    }
                    break;
                case (int)SelectedConflictTab.Lineweight:
                    if (SelectedConflict != null)
                    {
                        {
                            DataAccess da = new DataAccess();

                            var conflictsAndSources =
                                                      from source in SourceDrawingLayers
                                                      where SelectedConflict.Name == source.Name && SelectedConflict.Lineweight != source.Lineweight
                                                      select new { Conflict = SelectedConflict, Source = source };

                            foreach (var conflictAndSource in conflictsAndSources)
                                da.FixLayerConflict(conflictAndSource.Conflict.DrawingPath, conflictAndSource.Conflict.Name, "Lineweight", conflictAndSource.Source.Lineweight);

                            GetTargetDrawingLayers();
                            GetLineweightConflicts();
                            //LineweightConflictLayers.Remove(SelectedConflict);
                            LineweightConflictsCount = LineweightConflictLayers.Count.ToString() + MessagesAndNotifications.LineweightCountMessage;
                        }
                    }
                    break;
                case (int)SelectedConflictTab.Transparency:
                    if (SelectedConflict != null)
                    {
                        {
                            DataAccess da = new DataAccess();

                            var conflictsAndSources =
                                                      from source in SourceDrawingLayers
                                                      where SelectedConflict.Name == source.Name && SelectedConflict.Transparency != source.Transparency
                                                      select new { Conflict = SelectedConflict, Source = source };

                            foreach (var conflictAndSource in conflictsAndSources)
                                da.FixLayerConflict(conflictAndSource.Conflict.DrawingPath, conflictAndSource.Conflict.Name, "Transparency", conflictAndSource.Source.Transparency);

                            GetTargetDrawingLayers();
                            GetTransparencyConflicts();
                            //TransparencyConflictLayers.Remove(SelectedConflict);
                            TransparencyConflictsCount = TransparencyConflictLayers.Count.ToString() + MessagesAndNotifications.TransparencyCountMessage;
                        }
                    }
                    break;
                case (int)SelectedConflictTab.Plot:
                    if (SelectedConflict != null)
                    {
                        {
                            DataAccess da = new DataAccess();

                            var conflictsAndSources =
                                                      from source in SourceDrawingLayers
                                                      where SelectedConflict.Name == source.Name && SelectedConflict.Plot != source.Plot
                                                      select new { Conflict = SelectedConflict, Source = source };

                            foreach (var conflictAndSource in conflictsAndSources)
                                da.FixLayerConflict(conflictAndSource.Conflict.DrawingPath, conflictAndSource.Conflict.Name, "Plot", conflictAndSource.Source.Plot);

                            GetTargetDrawingLayers();
                            GetPlotConflicts();
                            //PlotConflictLayers.Remove(SelectedConflict);
                            PlotConflictsCount = PlotConflictLayers.Count.ToString() + MessagesAndNotifications.PlotCountMessage;
                        }
                    }
                    break;
            }
        }
        #endregion

        #region Resolve Viewport Layer Conflicts
        private void ResolveViewportFreezeConflictsClick()
        {
            if (ViewportFreezeConflictLayers != null && ViewportFreezeConflictLayers.Count != 0)
            {
                DataAccess da = new DataAccess();

                var conflictsAndSources = from conflict in ViewportFreezeConflictLayers
                                          from source in SourceDrawingViewportLayers
                                          where conflict.Name == source.Name && conflict.ViewportFreeze != source.ViewportFreeze && conflict.ViewportPosition == source.ViewportPosition
                                          select new { Conflict = conflict, Source = source };

                var Groups = conflictsAndSources.GroupBy(x => x.Conflict.DrawingPath, x => x.Source);

                foreach (var group in Groups)
                {
                    da.FixGroupOfViewportLayerConflicts(group.Key, "FreezeThaw", group);
                }

                GetTargetDrawingViewportLayers();
                GetViewportFreezeConflicts();
                //ViewportFreezeConflictLayers.Clear();
                //OnOffConflictsCount = OnOffConflictLayers.Count.ToString() + " Total On/Off conflicts with source drawing detected.";
                #region Original
                //foreach (var conflictAndSource in conflictsAndSources)
                //    da.FixViewportLayerConflict(conflictAndSource.Conflict.DrawingPath, conflictAndSource.Conflict.Name, "FreezeThaw", conflictAndSource.Source.ViewportFreeze);
                #endregion
            }
        }
        private void ResolveViewportColorConflictsClick()
        {
            if (ViewportColorConflictLayers != null && ViewportColorConflictLayers.Count != 0)
            {
                DataAccess da = new DataAccess();

                var conflictsAndSources = from conflict in ViewportColorConflictLayers
                                          from source in SourceDrawingViewportLayers
                                          where conflict.Name == source.Name && conflict.ViewportColor != source.ViewportColor && conflict.ViewportPosition == source.ViewportPosition
                                          select new { Conflict = conflict, Source = source };

                var Groups = conflictsAndSources.GroupBy(x => x.Conflict.DrawingPath, x => x.Source);

                foreach (var group in Groups)
                {
                    da.FixGroupOfViewportLayerConflicts(group.Key, "Color", group);
                }

                GetTargetDrawingViewportLayers();
                GetViewportColorConflicts();
                //ViewportColorConflictLayers.Clear();
                //ColorConflictsCount = ColorConflictLayers.Count.ToString() + MessagesAndNotifications.ColorCountMessage;
                #region Original
                //foreach (var conflictAndSource in conflictsAndSources)
                //    da.FixViewportLayerConflict(conflictAndSource.Conflict.DrawingPath, conflictAndSource.Conflict.Name, "Color", conflictAndSource.Source.ViewportColor);
                #endregion
            }
        }
        private void ResolveViewportLinetypeConflictsClick()
        {
            if (ViewportLinetypeConflictLayers != null && ViewportLinetypeConflictLayers.Count != 0)
            {
                DataAccess da = new DataAccess();

                var conflictsAndSources = from conflict in ViewportLinetypeConflictLayers
                                          from source in SourceDrawingViewportLayers
                                          where conflict.Name == source.Name && conflict.ViewportLinetype != source.ViewportLinetype && conflict.ViewportPosition == source.ViewportPosition
                                          select new { Conflict = conflict, Source = source };

                var Groups = conflictsAndSources.GroupBy(x => x.Conflict.DrawingPath, x => x.Source);

                foreach (var group in Groups)
                {
                    da.FixGroupOfViewportLayerConflicts(group.Key, "Linetype", group);
                }

                GetTargetDrawingViewportLayers();
                GetViewportLinetypeConflicts();
                //ViewportLinetypeConflictLayers.Clear();
                //LinetypeConflictsCount = LinetypeConflictLayers.Count.ToString() + MessagesAndNotifications.LinetypeCountMessage;
                #region Original
                //foreach (var conflictAndSource in conflictsAndSources)
                //    da.FixViewportLayerConflict(conflictAndSource.Conflict.DrawingPath, conflictAndSource.Conflict.Name, "Linetype", conflictAndSource.Source.ViewportLinetype); 
                #endregion
            }
        }
        private void ResolveViewportLineweightConflictsClick()
        {
            if (ViewportLineweightConflictLayers != null && ViewportLineweightConflictLayers.Count != 0)
            {
                DataAccess da = new DataAccess();

                var conflictsAndSources = from conflict in ViewportLineweightConflictLayers
                                          from source in SourceDrawingViewportLayers
                                          where conflict.Name == source.Name && conflict.ViewportLineweight != source.ViewportLineweight && conflict.ViewportPosition == source.ViewportPosition
                                          select new { Conflict = conflict, Source = source };

                var Groups = conflictsAndSources.GroupBy(x => x.Conflict.DrawingPath, x => x.Source);

                foreach (var group in Groups)
                {
                    da.FixGroupOfViewportLayerConflicts(group.Key, "Lineweight", group);
                }

                GetTargetDrawingViewportLayers();
                GetViewportLineWeightConflicts();
                //ViewportLineweightConflictLayers.Clear();
                //LineweightConflictsCount = LineweightConflictLayers.Count.ToString() + MessagesAndNotifications.LineweightCountMessage;
                #region Original
                //foreach (var conflictAndSource in conflictsAndSources)
                //    da.FixViewportLayerConflict(conflictAndSource.Conflict.DrawingPath, conflictAndSource.Conflict.Name, "Lineweight", conflictAndSource.Source.ViewportLineweight); 
                #endregion
            }
        }
        private void ResolveViewportTransparencyConflictsClick()
        {
            if (ViewportTransparencyConflictLayers != null && ViewportTransparencyConflictLayers.Count != 0)
            {
                DataAccess da = new DataAccess();

                var conflictsAndSources = from conflict in ViewportTransparencyConflictLayers
                                          from source in SourceDrawingViewportLayers
                                          where conflict.Name == source.Name && conflict.ViewportTransparency != source.ViewportTransparency && conflict.ViewportPosition == source.ViewportPosition
                                          select new { Conflict = conflict, Source = source };

                var Groups = conflictsAndSources.GroupBy(x => x.Conflict.DrawingPath, x => x.Source);

                foreach (var group in Groups)
                {
                    da.FixGroupOfViewportLayerConflicts(group.Key, "Transparency", group);
                }

                GetTargetDrawingViewportLayers();
                GetViewportTransparencyConflicts();
                //ViewportTransparencyConflictLayers.Clear();
                //TransparencyConflictsCount = TransparencyConflictLayers.Count.ToString() + MessagesAndNotifications.TransparencyCountMessage;
                #region Original
                //foreach (var conflictAndSource in conflictsAndSources)
                //    da.FixViewportLayerConflict(conflictAndSource.Conflict.DrawingPath, conflictAndSource.Conflict.Name, "Transparency", conflictAndSource.Source.ViewportTransparency); 
                #endregion
            }
        }
        #endregion

        #region Resolve Selected Viewport Layer Conflicts
        private void ResolveSelectedViewportConflictClick()
        {
            switch (SelectedViewportConflictTabIndex)
            {
                case (int)SelectedViewportConflictTab.ViewportFreeze:
                    if (SelectedViewportConflict != null)
                    {
                        {
                            if(ViewportFreezeConflictLayers.Count <= 0 == false)
                            {
                                ViewportFreezeConflictLayers.Remove(SelectedViewportConflict);
                            }
                            ViewportFreezeConflictsCount = ViewportFreezeConflictLayers.Count.ToString() + MessagesAndNotifications.ViewportFreezeCountMessage;
                        }
                    }
                    break;
                case (int)SelectedViewportConflictTab.ViewportColor:
                    if (SelectedViewportConflict != null)
                    {
                        if (ViewportColorConflictLayers.Count <= 0 == false)
                        {
                        ViewportColorConflictLayers.Remove(SelectedViewportConflict);
                        }
                        ViewportColorConflictsCount = ViewportColorConflictLayers.Count.ToString() + MessagesAndNotifications.ViewportColorCountMessage;
                    }
                    break;
                case (int)SelectedViewportConflictTab.ViewportLinetype:
                    if (SelectedViewportConflict != null)
                    {
                        if (ViewportColorConflictLayers.Count <= 0 == false)
                        {
                            ViewportColorConflictLayers.Remove(SelectedViewportConflict);
                        }
                        ViewportLinetypeConflictsCount = ViewportLinetypeConflictLayers.Count.ToString() + MessagesAndNotifications.ViewportLinetypeCountMessage;
                    }
                    break;
                case (int)SelectedViewportConflictTab.ViewportLineweight:
                    if (SelectedViewportConflict != null)
                    {
                        if (ViewportLineweightConflictLayers.Count <= 0 == false)
                        {
                            ViewportLineweightConflictLayers.Remove(SelectedViewportConflict);
                        }
                        ViewportLineweightConflictsCount = ViewportLineweightConflictLayers.Count.ToString() + MessagesAndNotifications.ViewportLineweightCountMessage;
                    }
                    break;
                case (int)SelectedViewportConflictTab.ViewportTransparency:
                    if (SelectedViewportConflict != null)
                    {
                        if (ViewportTransparencyConflictLayers.Count <= 0 == false)
                        {
                            ViewportTransparencyConflictLayers.Remove(SelectedViewportConflict);
                        }
                        ViewportTransparencyConflictsCount = ViewportTransparencyConflictLayers.Count.ToString() + MessagesAndNotifications.ViewportTransparencyCountMessage;
                    }
                    break;
            }
        }

        private void RemoveSelectedViewportConflictClick()
        {
            switch (SelectedViewportConflictTabIndex)
            {
                case (int)SelectedViewportConflictTab.ViewportFreeze:
                    if (SelectedViewportConflict != null)
                    {
                        {
                            DataAccess da = new DataAccess();

                            var conflictsAndSources =
                                                      from source in SourceDrawingViewportLayers
                                                      where SelectedViewportConflict.Name == source.Name && SelectedViewportConflict.ViewportFreeze != source.ViewportFreeze
                                                      select new { Conflict = SelectedViewportConflict, Source = source };

                            foreach (var conflictAndSource in conflictsAndSources)
                                da.FixViewportLayerConflict(conflictAndSource.Conflict.DrawingPath, conflictAndSource.Conflict.Name, "FreezeThaw", conflictAndSource.Source.ViewportFreeze);

                            GetTargetDrawingViewportLayers();
                            GetViewportFreezeConflicts();
                            //ViewportFreezeConflictLayers.Remove(SelectedViewportConflict);
                            ViewportFreezeConflictsCount = ViewportFreezeConflictLayers.Count.ToString() + MessagesAndNotifications.ViewportColorCountMessage;
                        }
                    }
                    break;
                case (int)SelectedViewportConflictTab.ViewportColor:
                    if (SelectedViewportConflict != null)
                    {
                        {
                            DataAccess da = new DataAccess();

                            var conflictsAndSources =
                                                      from source in SourceDrawingViewportLayers
                                                      where SelectedViewportConflict.Name == source.Name && SelectedViewportConflict.ViewportColor != source.ViewportColor
                                                      select new { Conflict = SelectedViewportConflict, Source = source };

                            foreach (var conflictAndSource in conflictsAndSources)
                                da.FixViewportLayerConflict(conflictAndSource.Conflict.DrawingPath, conflictAndSource.Conflict.Name, "Color", conflictAndSource.Source.ViewportColor);

                            GetTargetDrawingViewportLayers();
                            GetViewportColorConflicts();
                            //ViewportColorConflictLayers.Remove(SelectedViewportConflict);
                            ViewportColorConflictsCount = ViewportColorConflictLayers.Count.ToString() + MessagesAndNotifications.ViewportColorCountMessage;
                        }
                    }
                    break;
                case (int)SelectedViewportConflictTab.ViewportLinetype:
                    if (SelectedViewportConflict != null)
                    {
                        {
                            DataAccess da = new DataAccess();

                            var conflictsAndSources =
                                                      from source in SourceDrawingViewportLayers
                                                      where SelectedViewportConflict.Name == source.Name && SelectedViewportConflict.ViewportLinetype != source.ViewportLinetype
                                                      select new { Conflict = SelectedViewportConflict, Source = source };

                            foreach (var conflictAndSource in conflictsAndSources)
                                da.FixViewportLayerConflict(conflictAndSource.Conflict.DrawingPath, conflictAndSource.Conflict.Name, "Linetype", conflictAndSource.Source.ViewportLinetype);

                            GetTargetDrawingViewportLayers();
                            GetViewportLinetypeConflicts();
                            //ViewportLinetypeConflictLayers.Remove(SelectedViewportConflict);
                            ViewportLinetypeConflictsCount = ViewportLinetypeConflictLayers.Count.ToString() + MessagesAndNotifications.ViewportLinetypeCountMessage;
                        }
                    }
                    break;
                case (int)SelectedViewportConflictTab.ViewportLineweight:
                    if (SelectedViewportConflict != null)
                    {
                        {
                            DataAccess da = new DataAccess();

                            var conflictsAndSources =
                                                      from source in SourceDrawingViewportLayers
                                                      where SelectedViewportConflict.Name == source.Name && SelectedViewportConflict.ViewportLineweight != source.ViewportLineweight
                                                      select new { Conflict = SelectedViewportConflict, Source = source };

                            foreach (var conflictAndSource in conflictsAndSources)
                                da.FixViewportLayerConflict(conflictAndSource.Conflict.DrawingPath, conflictAndSource.Conflict.Name, "Lineweight", conflictAndSource.Source.ViewportLineweight);

                            GetTargetDrawingViewportLayers();
                            GetViewportLineWeightConflicts();
                            //ViewportLineweightConflictLayers.Remove(SelectedViewportConflict);
                            ViewportLineweightConflictsCount = ViewportLineweightConflictLayers.Count.ToString() + MessagesAndNotifications.ViewportLineweightCountMessage;
                        }
                    }
                    break;
                case (int)SelectedViewportConflictTab.ViewportTransparency:
                    if (SelectedViewportConflict != null)
                    {
                        {
                            DataAccess da = new DataAccess();

                            var conflictsAndSources =
                                                      from source in SourceDrawingViewportLayers
                                                      where SelectedViewportConflict.Name == source.Name && SelectedViewportConflict.ViewportTransparency != source.ViewportTransparency
                                                      select new { Conflict = SelectedViewportConflict, Source = source };

                            foreach (var conflictAndSource in conflictsAndSources)
                                da.FixViewportLayerConflict(conflictAndSource.Conflict.DrawingPath, conflictAndSource.Conflict.Name, "Transparency", conflictAndSource.Source.ViewportTransparency);

                            GetTargetDrawingViewportLayers();
                            GetViewportTransparencyConflicts();
                            //ViewportTransparencyConflictLayers.Remove(SelectedViewportConflict);
                            ViewportTransparencyConflictsCount = ViewportTransparencyConflictLayers.Count.ToString() + MessagesAndNotifications.ViewportTransparencyCountMessage;
                        }
                    }
                    break;
            }
        }

        #endregion

        #region Helpers
        /// <summary>
        /// Clears out all collecetion and property data from the application.
        /// </summary>
        private void ClearComparison()
        {
            ComparisonRan = false;
            if (TargetDrawings != null && TargetDrawings.Count != 0)
            {
                TargetDrawings.Clear();
            }
            if (TargetDrawingLayers != null && TargetDrawingLayers.Count != 0)
            {
                TargetDrawingLayers.Clear();
            }
            if (TargetDrawingViewportLayers != null && TargetDrawingViewportLayers.Count != 0)
            {
                TargetDrawingViewportLayers.Clear();
            }
            if (DrawingModels != null && DrawingModels.Count != 0)
            {
                DrawingModels.Clear();
            }
            if (AllConflictLayers != null && AllConflictLayers.Count != 0)
            {
                AllConflictLayers.Clear();
            }
            if (OnOffConflictLayers != null && OnOffConflictLayers.Count != 0)
            {
                OnOffConflictLayers.Clear();
            }
            if (FreezeConflictLayers != null && FreezeConflictLayers.Count != 0)
            {
                FreezeConflictLayers.Clear();
            }
            if (ColorConflictLayers != null && ColorConflictLayers.Count != 0)
            {
                ColorConflictLayers.Clear();
            }
            if (LinetypeConflictLayers != null && LinetypeConflictLayers.Count != 0)
            {
                LinetypeConflictLayers.Clear();
            }
            if (LineweightConflictLayers != null && LineweightConflictLayers.Count != 0)
            {
                LineweightConflictLayers.Clear();
            }
            if (TransparencyConflictLayers != null && TransparencyConflictLayers.Count != 0)
            {
                TransparencyConflictLayers.Clear();
            }
            if (PlotConflictLayers != null && PlotConflictLayers.Count != 0)
            {
                PlotConflictLayers.Clear();
            }
            if (ExtraLayers != null && ExtraLayers.Count != 0)
            {
                ExtraLayers.Clear();
            }
            if (MissingLayers != null && MissingLayers.Count != 0)
            {
                MissingLayers.Clear();
            }
            if (ViewportFreezeConflictLayers != null && ViewportFreezeConflictLayers.Count != 0)
            {
                ViewportFreezeConflictLayers.Clear();
            }
            if (ViewportColorConflictLayers != null && ViewportColorConflictLayers.Count != 0)
            {
                ViewportColorConflictLayers.Clear();
            }
            if (ViewportLinetypeConflictLayers != null && ViewportLinetypeConflictLayers.Count != 0)
            {
                ViewportLinetypeConflictLayers.Clear();
            }
            if (ViewportTransparencyConflictLayers != null && ViewportTransparencyConflictLayers.Count != 0)
            {
                ViewportTransparencyConflictLayers.Clear();
            }
        }

        /// <summary>
        /// Clears out all of the comparisons without removing the <see cref="TargetDrawings"/>.
        /// </summary>
        private void PartialClearComparison()
        {
            if (TargetDrawingLayers != null && TargetDrawingLayers.Count != 0)
            {
                TargetDrawingLayers.Clear();
            }
            if (TargetDrawingViewportLayers != null && TargetDrawingViewportLayers.Count != 0)
            {
                TargetDrawingViewportLayers.Clear();
            }
            if (DrawingModels != null && DrawingModels.Count != 0)
            {
                DrawingModels.Clear();
            }
            if (AllConflictLayers != null && AllConflictLayers.Count != 0)
            {
                AllConflictLayers.Clear();
            }
            if (OnOffConflictLayers != null && OnOffConflictLayers.Count != 0)
            {
                OnOffConflictLayers.Clear();
            }
            if (FreezeConflictLayers != null && FreezeConflictLayers.Count != 0)
            {
                FreezeConflictLayers.Clear();
            }
            if (ColorConflictLayers != null && ColorConflictLayers.Count != 0)
            {
                ColorConflictLayers.Clear();
            }
            if (LinetypeConflictLayers != null && LinetypeConflictLayers.Count != 0)
            {
                LinetypeConflictLayers.Clear();
            }
            if (LineweightConflictLayers != null && LineweightConflictLayers.Count != 0)
            {
                LineweightConflictLayers.Clear();
            }
            if (TransparencyConflictLayers != null && TransparencyConflictLayers.Count != 0)
            {
                TransparencyConflictLayers.Clear();
            }
            if (PlotConflictLayers != null && PlotConflictLayers.Count != 0)
            {
                PlotConflictLayers.Clear();
            }
            if (ExtraLayers != null && ExtraLayers.Count != 0)
            {
                ExtraLayers.Clear();
            }
            if (MissingLayers != null && MissingLayers.Count != 0)
            {
                MissingLayers.Clear();
            }
            if (ViewportFreezeConflictLayers != null && ViewportFreezeConflictLayers.Count != 0)
            {
                ViewportFreezeConflictLayers.Clear();
            }
            if (ViewportColorConflictLayers != null && ViewportColorConflictLayers.Count != 0)
            {
                ViewportColorConflictLayers.Clear();
            }
            if (ViewportLinetypeConflictLayers != null && ViewportLinetypeConflictLayers.Count != 0)
            {
                ViewportLinetypeConflictLayers.Clear();
            }
            if (ViewportTransparencyConflictLayers != null && ViewportTransparencyConflictLayers.Count != 0)
            {
                ViewportTransparencyConflictLayers.Clear();
            }
        }

        /// <summary>
        /// Builds a list of type <see cref="DrawingModel"/> and compares the layers within them to the source drawings layers. Used to find layers that are missing from the <see cref="TargetDrawings"/>.
        /// </summary>
        /// <param name="sourceLayers"></param>
        /// <returns></returns>
        private ObservableCollection<MissingLayerModel> ReturnMissingLayers(ObservableCollection<LayerComparisonLayerModel> sourceLayers)
        {
            ObservableCollection<MissingLayerModel> missingLayers = new ObservableCollection<MissingLayerModel>();

            try
            {
                if (TargetDrawings != null && TargetDrawings.Count != 0)
                {
                    foreach (var drawing in TargetDrawings)
                    {
                        DataAccess da = new DataAccess();
                        DrawingModel drawingModel = da.CreateDrawingModel(drawing.DrawingPath);
                        {
                            DrawingModels.Add(drawingModel);
                        }
                    }

                    foreach (var sourceLayer in SourceDrawingLayers)
                    {
                        foreach (var drawingModel in DrawingModels)
                        {
                            DataAccess da = new DataAccess();
                            if (da.FindDrawingModelLayer(drawingModel, sourceLayer.Name) == false)
                            {
                                MissingLayerModel missingLayer = new MissingLayerModel();
                                missingLayer.TargetDrawingPath = drawingModel.DrawingPath;
                                missingLayer.SourceLayerName = sourceLayer.Name;
                                missingLayers.Add(missingLayer);
                                MissingLayers = missingLayers;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return missingLayers;
        }

        /// <summary>
        /// Loads the RecentItemsList with each item on the user's RecentComparisons.txt file.
        /// </summary>
        /// <param name="filePath"></param>
        private void LoadRecentItemsCommand(string filePath)
        {
            List<string> missingFiles = new List<string>();
            try
            {
                using (var fileStream = File.OpenRead(filePath))
                {
                    using (var streamReader = new StreamReader(fileStream.Name))
                    {
                        string source = streamReader.ReadLine() ?? "";
                        if (source != null && source.Contains(".dwg"))
                        {
                            SetSourceDrawing(source);
                            GetSourceDrawingLayersClick();
                        }
                        else
                        {
                            source = MessagesAndNotifications.SourceDrawingNotSet;
                        }
                        foreach (var line in File.ReadLines(filePath, System.Text.Encoding.GetEncoding(1250)))
                        {
                            if (File.Exists(line) == true)
                            {
                                TargetDrawingModel targetDrawingModel = new TargetDrawingModel();
                                if (line != null && line.Contains(".dwg"))
                                {
                                    targetDrawingModel.DrawingPath = line;
                                    TargetDrawings.Add(targetDrawingModel);
                                }
                                else
                                {
                                    MessageBox.Show(MessagesAndNotifications.TargetDrawingsNotSet, "Target Drawings Not Found", MessageBoxButton.OK, MessageBoxImage.Warning);
                                }
                            }
                            else
                            {
                                missingFiles.Add(line);
                            }
                        }
                        if (missingFiles.Count > 0)
                        {
                            MessageBox.Show($"The following drawings could not be located and weren't added\r\n{string.Join(Environment.NewLine, missingFiles)}", "Missing drawings", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                        IoC.Get<LayerComparisonApplicationViewModel>().GoToPage(ApplicationPage.MainPage);
                        MainPageTabIndex = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Enables the VISRETAIN system variable in all target drawings.
        /// </summary>
        private void VisretainCommandClick()
        {
            try
            {
                if (TargetDrawings != null && TargetDrawings.Count != 0)
                {
                    DataAccess da = new DataAccess();
                    foreach (var target in TargetDrawings)
                    {
                        da.EnableVisretain(target.DrawingPath);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ReconcileAllLayersCommandClick()
        {
            try
            {
                if (TargetDrawings != null && TargetDrawings.Count != 0)
                {
                    DataAccess da = new DataAccess();
                    da.ReconcileAllLayers(SourceDrawing);
                    foreach (var target in TargetDrawings)
                    {
                        da.ReconcileAllLayers(target.DrawingPath);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        private void BuildLayerItems()
        {
            TempDataAccess tempDataAccess = new TempDataAccess();
            LayerItemList = tempDataAccess.BuildLayerItem(SourceDrawing);

            foreach (var layerItem in LayerItemList)
            {
                LayerItemViewModel control = new LayerItemViewModel();
                control.Name = layerItem.Name;
                control.Drawing = layerItem.Drawing;
                control.On = layerItem.On;
                control.Freeze = layerItem.Freeze;
                control.Color = layerItem.Color;
                control.Linetype = layerItem.Linetype;
                control.Lineweight = layerItem.Lineweight;
                control.Transparency = layerItem.Transparency;
                control.Plot = layerItem.Plot;

                LayerControls.Add(control);
            }
        }
    }
}
