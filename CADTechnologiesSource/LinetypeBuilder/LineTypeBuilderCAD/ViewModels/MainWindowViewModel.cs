using LineTypeBuilderCAD.Base;
using System;
using System.Collections.Generic;
using System.IO;

namespace LineTypeBuilderCAD.Viewmodels
{
    public class MainWindowViewModel : BaseViewModel
    {
        #region Singleton
        public static MainWindowViewModel LTBuilderInstance = new MainWindowViewModel();
        #endregion

        #region Public Properties

        public string LinetypeName { get; set; }

        public string LinetypeDescription { get; set; }

        public List<string> LinetypeStyle { get; set; }

        public double SelectedLinetypeIndex { get; set; }

        public string SelectedLinetypeStyle { get; set; }

        public double DashLength { get; set; }

        public double TextMargin { get; set; }

        public string TextContents { get; set; }

        public string TextStyle { get; set; }

        public double TextScale { get; set; }

        public double TextRotation { get; set; }

        public double HorizontalOffset { get; set; }

        public double VerticalOffset { get; set; }

        public double DashSpacing { get; set; }

        public string CompletedLin { get; set; }

        #endregion

        #region Public Commands

        public RelayCommand CreateLinetypeCommand { get; set; }
        public RelayCommand CopyToClipboardCommand { get; set; }

        #endregion

        #region Constructor

        public MainWindowViewModel()
        {
            if (LinetypeStyle == null || LinetypeStyle.Count != 0)
            {
                LinetypeStyle = new List<string>();
                LinetypeStyle.Add("Aligned");
                LinetypeStyle.Add("Rotated");
            }
            SelectedLinetypeIndex = 0;
            TextStyle = "standard";
            TextScale = 0.08;
            TextContents = "contents";
            LinetypeName = "name";
            LinetypeDescription = "description";
            CreateLinetypeCommand = new RelayCommand(() => CreateLinetypeCommandClick());
            CopyToClipboardCommand = new RelayCommand(() => CopyToClipboardCommandClick());
        }

        #endregion

        #region CommandMethods
        private void CreateLinetypeCommandClick()
        {

        }

        private void CopyToClipboardCommandClick()
        {

            string userStyle = "";
            switch (SelectedLinetypeIndex)
            {
                case 0:
                    userStyle = "A";
                    break;
                case 1:
                    userStyle = "R";
                    break;
            }

            CompletedLin = "*" + LinetypeName + "," + LinetypeDescription + System.Environment.NewLine + userStyle + "," + DashLength.ToString() + "," + DashSpacing.ToString() + "," + "[" + '\u0022' + TextContents + '\u0022' +  "," + TextStyle + "," + "s=" + TextScale.ToString() + "," + "R=" + TextRotation.ToString() + "," + "X=" + HorizontalOffset + "," + "Y=" + VerticalOffset  + "]" + ","  + DashSpacing.ToString();
            if (string.IsNullOrEmpty(CompletedLin) == false)
            {
                var filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "completed_lintype.txt");
                System.IO.File.WriteAllText(filePath, "Paste the following onto a fresh line in your .lin file. I've already copied it to your clipboard." + System.Environment.NewLine + System.Environment.NewLine + CompletedLin);
                System.Windows.Clipboard.SetText(CompletedLin);
                System.Diagnostics.Process.Start(filePath);
            }
        }

        #endregion
    }
}
