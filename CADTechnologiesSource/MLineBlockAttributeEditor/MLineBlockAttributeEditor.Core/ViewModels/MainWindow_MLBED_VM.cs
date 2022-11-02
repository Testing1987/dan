using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using CADTechnologiesSource.All.AutoCADHelpers;
using CADTechnologiesSource.All.Base;
using CADTechnologiesSource.All.Models;
using MLineBlockAttributeEditor.Core.Constants;
using MLineBlockAttributeEditor.Core.CoreLogic;
using System;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Input;

namespace MLineBlockAttributeEditor.Core.ViewModels
{
    public class MainWindow_MLBED_VM : BaseViewModel
    {
        /// <summary>
        /// A single instance of the view model to bind to.
        /// </summary>
        #region Singleton
        public static MainWindow_MLBED_VM Instance = new MainWindow_MLBED_VM();
        #endregion

        #region Public Properties
        //Set up our application's properties for binding.
        public bool TextBoxEnabled { get; set; } = false;
        public ObjectId SelectedBlockId { get; set; }
        public string SelectedBlockObjectIDString { get; set; }
        public string SelectedBlockName { get; set; }
        public int TagComboBoxSelectedIndex { get; set; }
        public string SelectedTagOnComboBox { get; set; }
        public string AttributeLayer { get; set; }
        public Color AttributeColor { get; set; }
        public string AttributeLinetype { get; set; }
        public LineWeight AttributeLineweight { get; set; }
        public string AttributeLineWeightString { get; set; }
        public AttachmentPoint AttributeJustification { get; set; }
        public string AttributeJustificationString { get; set; }
        public double AttributeTextHeight { get; set; } = 0;
        public double AttributeRotation { get; set; } = 0;
        public double AttributeWidthFactor { get; set; } = 0;
        public string AttributeMTextContent { get; set; }
        public string AttributeTextString { get; set; }
        public bool AttributeBackgroundFill { get; set; }
        public bool AttributeUseBackgroundColor { get; set; }
        public ObservableCollection<BlockAttributeModel> BlockAttributeData { get; set; }
        public ObservableCollection<string> DrawingLayers { get; set; }
        public ObservableCollection<string> DrawingLinetypes { get; set; }


        #endregion

        #region Public Commands
        /// <summary>
        /// Runs the <see cref="SelectBlockCommandClick"/> method.
        /// </summary>
        public ICommand SelectBlockCommand { get; set; }
        public ICommand AttributeSelectionChangedCommand { get; set; }
        public ICommand UpdateAttributeWidthFactorCommand { get; set; }
        public ICommand UpdateAttributeTextValuesCommand { get; set; }
        public ICommand ApplyBackgroundMaskCommand { get; set; }
        public ICommand ClearCommand { get; set; }
        public ICommand TestCommand { get; set; }
        #endregion

        #region Constructor
        /// <summary>
        /// Default constructor.
        /// </summary>
        public MainWindow_MLBED_VM()
        {
            #region Instantiate Collections
            // Instantiate the DrawingLayers collection and make it not null.
            if (DrawingLayers == null)
            {
                DrawingLayers = new ObservableCollection<string>();
            }
            // instatntiate the DrawingLinetypes collection and make it not null.
            if (DrawingLinetypes == null)
            {
                DrawingLinetypes = new ObservableCollection<string>();
            }
            //Instantiate the Block Attribute Collection and make it not null.
            if (BlockAttributeData == null)
            {
                BlockAttributeData = new ObservableCollection<BlockAttributeModel>();
            }
            #endregion

            #region CommandDefinitions
            //Prompt the user to select a block attribute.
            SelectBlockCommand = new RelayCommand(() => SelectBlockCommandClick());
            AttributeSelectionChangedCommand = new RelayCommand(() => AttributeSelectionChanged());
            UpdateAttributeWidthFactorCommand = new RelayCommand(() => UpdatedAttributeWidthFactorClick());
            UpdateAttributeTextValuesCommand = new RelayCommand(() => UpdateAttributeTextValuesTyping());
            ApplyBackgroundMaskCommand = new RelayCommand(() => ApplyBackgroundMaskCommandClick());
            ClearCommand = new RelayCommand(() => ClearCommandClick());
            TestCommand = new RelayCommand(() => TestCommandClick());
            #endregion
        }


        #endregion

        #region CommandMethods
        /// <summary>
        /// Prompts the user to select a <see cref="BlockReference"/> and populates <see cref="BlockAttributeData"/> from that selection.
        /// </summary>
        private void SelectBlockCommandClick()
        {
            try
            {
                //ClearCommandClick();
                //First, Clear the drawing layers collection...
                if (DrawingLayers != null && DrawingLayers.Count != 0)
                {
                    DrawingLayers.Clear();
                }

                // then re-build the drawing Layers list to capture any changes the user may have made to their layers since the last time the button was pressed
                DataAccess da = new DataAccess();
                DrawingLayers = da.GetDrawingLayerNames();

                // then re-build the drawing linetypes list to capture any changes the user may have made to the drawing since the last time the button was pressed
                if (DrawingLinetypes != null && DrawingLinetypes.Count != 0)
                {
                    DrawingLinetypes.Clear();
                }
                DrawingLinetypes = da.GetDrawingLinetypes();

                // Then clear the previous block attribute data
                if (BlockAttributeData != null && BlockAttributeData.Count != 0)
                {
                    BlockAttributeData.Clear();
                }

                // Set the user's screen focus to AutoCAD and access the needed database information
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Database thisDatabase = thisDrawing.Database;
                Editor thisEditor = thisDrawing.Editor;

                // Ask the user to select a block reference...
                PromptEntityOptions peo = new PromptEntityOptions("");
                peo.Message = "\nSelect a block reference.";
                peo.SetRejectMessage("\nYou must select a block reference.");
                peo.AddAllowedClass(typeof(BlockReference), false);
                peo.AllowNone = true;

                // Store the result of the selection
                PromptEntityResult per = thisDrawing.Editor.GetEntity(peo);
                if (per.Status == PromptStatus.OK)
                {
                    //Get the objectid and set the property
                    SelectedBlockId = per.ObjectId;
                    DatabaseHelpers dh = new DatabaseHelpers();
                    SelectedBlockName = dh.GetBlockName(thisDrawing, per.ObjectId);

                    //Access the data from the selected block
                    foreach (var result in da.BuildBlockData(per, thisDrawing, thisDatabase))
                    {
                        BlockAttributeData.Add(result);
                        TagComboBoxSelectedIndex = 0;
                        if (BlockAttributeData.Count == 0)
                        {
                            MessageBox.Show(ErrorMessages.NoBlockAttributesWereMultiLine);
                        }
                    }
                }
                TextBoxEnabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        /// <summary>
        /// Updates the UI elements when the user changes the selected attribute in the attribute selection combobox.
        /// </summary>
        private void AttributeSelectionChanged()
        {
            //MessageBox.Show($"You Changed The Combobox Selection. It is now {SelectedTagOnComboBox}");

            DataAccess da = new DataAccess();
            BlockAttributeModel blockAttributeModel = da.ReturnAttributeData(SelectedBlockId, SelectedTagOnComboBox);

            AttributeTextString = blockAttributeModel.TextString;
            AttributeMTextContent = blockAttributeModel.MtextAttributeContent;
            if (blockAttributeModel.MtextAttributeContent != null)
            {
                AttributeTextString = blockAttributeModel.MtextAttributeContent.Substring(blockAttributeModel.MtextAttributeContent.IndexOf(';') + 1).Replace("\\P", "\r\n");
            }
            AttributeLayer = blockAttributeModel.Layer;
            AttributeColor = blockAttributeModel.Color;
            AttributeLinetype = blockAttributeModel.Linetype;
            AttributeLineweight = blockAttributeModel.Lineweight;
            AttributeLineWeightString = blockAttributeModel.LineWeightString;
            AttributeJustification = blockAttributeModel.Justification;
            AttributeJustificationString = blockAttributeModel.JustificationString;
            AttributeTextHeight = blockAttributeModel.Height;
            AttributeRotation = Math.Round(blockAttributeModel.Rotation, 7);
            AttributeWidthFactor = blockAttributeModel.WidthFactor;
            //AttributeBackgroundFill = blockAttributeModel.BackgroundFill;
            //AttributeUseBackgroundColor = blockAttributeModel.UseBackgroundFill;
            //if(AttributeBackgroundFill == true)
            //{
            //    AttributeBackgroundFillColorString = blockAttributeModel.BackgroundFillColor.ToString();
            //    AttributeBackroundMaskScaleFactor = blockAttributeModel.BackgroundMaskScaleFactor;
            //}
            SelectedBlockObjectIDString = SelectedBlockId.ToString();
            SelectedBlockName = blockAttributeModel.BlockName;
        }

        /// <summary>
        /// updates the width factor of the selected block attribute
        /// </summary>
        private void UpdatedAttributeWidthFactorClick()
        {
            var thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.CurrentDocument;
            // Access the DataAccess class...
            DataAccess da = new DataAccess();

            // Create a new blockAttributeModel that takes in the data our user updated
            BlockAttributeModel updatedBlockAttributeModel = new BlockAttributeModel();

            // populate the updatedBlockAttributeModel with our user's input
            updatedBlockAttributeModel.Tag = SelectedTagOnComboBox;
            AttributeTextString = AttributeTextString.Substring(AttributeTextString.IndexOf(';') + 1).Replace("\\P", "\r\n").TrimEnd('}');
            updatedBlockAttributeModel.TextString = AttributeTextString;
            updatedBlockAttributeModel.TextString = "{\\W" + AttributeWidthFactor.ToString() + ";" + AttributeTextString + "}";
            AttributeTextString = updatedBlockAttributeModel.TextString;

            // Update the block
            da.UpdateBlockAttributeReferenceWidthFactorViaTextString(SelectedBlockId, updatedBlockAttributeModel);
            thisDrawing.Editor.UpdateScreen();
        }

        /// <summary>
        /// Updates the block attribute text as the user is typing.
        /// </summary>
        private void UpdateAttributeTextValuesTyping()
        {
            var thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.CurrentDocument;
            // Access the DataAccess class...
            DataAccess da = new DataAccess();

            // Create a new blockAttributeModel that takes in the data our user updated
            BlockAttributeModel updatedBlockAttributeModel = new BlockAttributeModel();

            // populate the updatedBlockAttributeModel with our user's input
            updatedBlockAttributeModel.Tag = SelectedTagOnComboBox;
            updatedBlockAttributeModel.MtextAttributeContent = AttributeMTextContent;
            updatedBlockAttributeModel.TextString = "{\\W" + AttributeWidthFactor.ToString() + AttributeTextString + "}";
            //updatedBlockAttributeModel.Layer = AttributeLayer;
            //updatedBlockAttributeModel.Color = AttributeColor;
            //updatedBlockAttributeModel.Linetype = AttributeLinetype;
            //updatedBlockAttributeModel.Lineweight = AttributeLineweight;
            //updatedBlockAttributeModel.Justification = AttributeJustification;
            updatedBlockAttributeModel.Height = AttributeTextHeight;
            updatedBlockAttributeModel.Rotation = AttributeRotation;
            updatedBlockAttributeModel.WidthFactor = AttributeWidthFactor;

            // Update the block
            da.UpdateBlockAttributeReference(SelectedBlockId, updatedBlockAttributeModel);
            thisDrawing.Editor.UpdateScreen();
            //Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();  
        }

        /// <summary>
        /// Applies a simple background mask to the selected attribute.
        /// </summary>
        private void ApplyBackgroundMaskCommandClick()
        {
            var thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.CurrentDocument;
            // Access the DataAccess class...
            DataAccess da = new DataAccess();

            // Create a new blockAttributeModel that takes in the data our user updated
            BlockAttributeModel updatedBlockAttributeModel = new BlockAttributeModel();
            updatedBlockAttributeModel.Tag = SelectedTagOnComboBox;


            // Update the block
            da.ApplyBackgroundMask(SelectedBlockId, updatedBlockAttributeModel);
            thisDrawing.Editor.UpdateScreen();
            //Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView(); 
        }

        
        /// <summary>
        /// Clear all collections and properties.
        /// </summary>
        private void ClearCommandClick()
        {
            if (SelectedBlockObjectIDString != null)
            {
                if (DrawingLayers != null)
                {
                    DrawingLayers.Clear();
                }
                if (DrawingLinetypes != null)
                {
                    DrawingLinetypes.Clear();
                }
                if (BlockAttributeData != null)
                {
                    BlockAttributeData.Clear();
                }
                SelectedBlockId = new ObjectId();
                SelectedBlockObjectIDString = "";
                SelectedBlockName = "";
                TagComboBoxSelectedIndex = 0;
                SelectedTagOnComboBox = "";
                AttributeLayer = "";
                AttributeColor = new Color();
                AttributeLinetype = "";
                AttributeLineweight = new LineWeight();
                AttributeLineWeightString = "";
                AttributeJustification = new AttachmentPoint();
                AttributeJustificationString = "";
                AttributeTextHeight = 0;
                AttributeRotation = 0;
                AttributeWidthFactor = 0;
                AttributeMTextContent = "";
                AttributeTextString = "";
                if (TextBoxEnabled == true)
                {
                    TextBoxEnabled = false;
                }
            }
        }


        private void TestCommandClick()
        {
            var thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.CurrentDocument;
            DataAccess da = new DataAccess();

            BlockAttributeModel updatedBlockAttributeModel = new BlockAttributeModel();

            updatedBlockAttributeModel.Tag = SelectedTagOnComboBox;

            updatedBlockAttributeModel.TextString = "{\\W0.5;THIS IS A TEST}";
            //updatedBlockAttributeModel.MtextAttributeContent = AttributeMTextContent;
            //updatedBlockAttributeModel.Layer = AttributeLayer;
            //updatedBlockAttributeModel.Color = AttributeColor;
            //updatedBlockAttributeModel.Linetype = AttributeLinetype;
            //updatedBlockAttributeModel.Lineweight = AttributeLineweight;
            //updatedBlockAttributeModel.Justification = AttributeJustification;
            updatedBlockAttributeModel.Height = AttributeTextHeight;
            updatedBlockAttributeModel.Rotation = AttributeRotation;
            updatedBlockAttributeModel.WidthFactor = AttributeWidthFactor;

            da.UpdateBlockAttributeReference(SelectedBlockId, updatedBlockAttributeModel);
            thisDrawing.Editor.UpdateScreen();
        }

        #endregion
    }
}
