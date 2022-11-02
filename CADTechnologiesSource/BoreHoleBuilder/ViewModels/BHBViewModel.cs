using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Windows.Data;
using BoreExcavationBuilder.Base;
using BoreExcavationBuilder.CoreLogic;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;

namespace BoreExcavationBuilder.ViewModels
{
    public class BHBViewModel : BaseViewModel
    {
        #region Singleton
        public static BHBViewModel bHBViewModel = new BHBViewModel();
        #endregion

        #region Public Properties
        public bool UseHatch { get; set; }

        public double Length { get; set; }

        public double Width { get; set; }

        public string SelectedHatch { get; set; }

        public string BoreHoleLayer { get; set; }

        public string HatchLayer { get; set; }

        public double HatchScale { get; set; }

        public double HatchAngle { get; set; }

        public List<string> AddedHatches { get; set; }

        #endregion

        #region Public Commands
        public ICommand DrawCommand { get; set; }
        #endregion

        #region Constructor
        public BHBViewModel()
        {
            DrawCommand = new RelayCommand(() => DrawCommandClick());
            Length = 60;
            Width = 15;
            BoreHoleLayer = "0";
            HatchLayer = "0";
            HatchScale = 1;
            HatchAngle = 0;
            if (AddedHatches == null)
            {
                AddedHatches = new List<string>();
            }
            //Add all hatch patterns in the drawing to the list
            Editor thisEditor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;

            if (AddedHatches != null || AddedHatches.Count != 0)
            {
                AddedHatches.Clear();
            }
            try
            {
                foreach (string hatchPatternName in HatchPatterns.Instance.PredefinedPatterns)
                {
                    AddedHatches.Add(hatchPatternName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        private void DrawCommandClick()
        {
            if (Length <= 0 || Width <= 0)
            {
                MessageBox.Show("Length and Width must be greater than 0");
                return;
            }

            DataAccess dataAccess = new DataAccess();
            dataAccess.DrawBoreHole(Length, Width, UseHatch, SelectedHatch, BoreHoleLayer, HatchLayer, HatchScale, HatchAngle);
        }
    }
}
