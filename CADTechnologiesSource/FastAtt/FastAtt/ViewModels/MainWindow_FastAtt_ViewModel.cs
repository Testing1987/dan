using Autodesk.AutoCAD.DatabaseServices;
using FastAtt.Base;
using FastAtt.CoreLogic;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace FastAtt.ViewModels
{
    public class MainWindow_FastAtt_ViewModel : BaseViewModel
    {
        #region Singleton
        public static MainWindow_FastAtt_ViewModel instance = new MainWindow_FastAtt_ViewModel();
        #endregion

        #region PublicProperties
        //Properties
        public bool Invisible { get; set; }
        public bool Constant { get; set; }
        public bool Verify { get; set; }
        public bool Preset { get; set; }
        public bool LockPosition { get; set; }
        public bool MultipleLines { get; set; }
        public bool SpecifyInsertionInApp { get; set; }
        public double InsertX { get; set; }
        public double InsertY { get; set; }
        public double InsertZ { get; set; }
        public string Tag { get; set; }
        public List<string> Justifications { get; set; }
        public string SelectedJustification { get; set; }
        public List<string> TextStyles { get; set; }
        public string SelectedTextStyle { get; set; }
        public bool Annotative { get; set; }
        public double TextHeight { get; set; }
        public double Rotation { get; set; }

        //Commands
        public ICommand RunCommand { get; set; }
        #endregion

        #region Constructor
        public MainWindow_FastAtt_ViewModel()
        {
            #region Text Styles
            //Make sure TextStyles isn't null
            if (TextStyles == null)
            {
                TextStyles = new List<string>();
            }
            if (TextStyles.Count() > 0)
            {
                TextStyles.Clear();
            }
            //and add the drawing's textstyles to the list
            CADAccess CA = new CADAccess();
            foreach (var result in CA.GetTextStylesFromCurrentDrawing())
            {
                TextStyles.Add(result);
            }
            SelectedTextStyle = TextStyles.First();
            #endregion

            #region Justifications
            //Make sure Justification isn't null
            if (Justifications == null)
            {
                Justifications = new List<string>();
            }
            if(Justifications.Count() > 0)
            Justifications.Clear();
            //Add Justifications from Autocad
            //foreach (var attachmentPoint in ((Autodesk.AutoCAD.DatabaseServices.AttachmentPoint[])Enum.GetValues(typeof(AttachmentPoint))).Distinct().ToString())
            //{
            //}

            //Justifications.Add(AttachmentPoint.BaseLeft);
            //Justifications.Add(AttachmentPoint.BaseCenter);
            //Justifications.Add(AttachmentPoint.BaseRight);
            //Justifications.Add(AttachmentPoint.BaseAlign);
            //Justifications.Add(AttachmentPoint.MiddleAlign);
            //Justifications.Add(AttachmentPoint.BaseFit);
            Justifications.Add(AttachmentPoint.TopLeft.ToString());
            Justifications.Add(AttachmentPoint.TopCenter.ToString());
            Justifications.Add(AttachmentPoint.TopRight.ToString());
            Justifications.Add(AttachmentPoint.MiddleLeft.ToString());
            Justifications.Add(AttachmentPoint.MiddleCenter.ToString());
            Justifications.Add(AttachmentPoint.MiddleRight.ToString());
            Justifications.Add(AttachmentPoint.BottomLeft.ToString());
            Justifications.Add(AttachmentPoint.BottomCenter.ToString());
            Justifications.Add(AttachmentPoint.BottomRight.ToString());
            SelectedJustification = Justifications.First();
            #endregion

            #region Everything Else
            //set up the rest of the default properties
            LockPosition = true;

            SpecifyInsertionInApp = false;
            InsertX = 0.00;
            InsertY = 0.00;
            InsertZ = 0.00;

            Tag = "REV";
            Annotative = false;
            TextHeight = 0.08;
            Rotation = 0;
            #endregion

            #region Commands
            RunCommand = new RelayCommand(() => RunCommandClick());
            #endregion

        }
        #endregion

        #region Command Methods
        private void RunCommandClick()
        {
            CADAccess CA = new CADAccess();
            CA.CreateBlockAttribute(Invisible, Constant, Verify, Preset, LockPosition, MultipleLines, SpecifyInsertionInApp, InsertX, InsertY, InsertZ, Tag, SelectedTextStyle, SelectedJustification, Annotative, TextHeight, Rotation);
        }
        #endregion

    }
}
