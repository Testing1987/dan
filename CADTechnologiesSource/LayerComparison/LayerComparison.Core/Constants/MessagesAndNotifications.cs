using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LayerComparison.Core.Constants
{
    public class MessagesAndNotifications
    {
        public const string AllConflictsCountMessage = " Total layers found that conflict with the source drawing.";
        public const string OnOffCountMessage = " Total layers found where On/Off conflicts with the source drawing.";
        public const string FreezeCountMessage = " Total layers found where Freeze/Thaw conflicts with the source drawing.";
        public const string ColorCountMessage = " Total layers found where Color conflicts with the source drawing.";
        public const string LinetypeCountMessage = " Total layers found where Linetype conflicts with the source drawing.";
        public const string LineweightCountMessage = " Total layers found where Lineweight conflicts with the source drawing.";
        public const string TransparencyCountMessage = " Total layers found where Transparency conflicts with the source drawing.";
        public const string ViewportFreezeCountMessage = " Total layers found where Freeze/Thaw viewport override conflicts with the source drawing.";
        public const string ViewportColorCountMessage = " Total layers found where Color viewport override conflicts with the source drawing.";
        public const string ViewportLinetypeCountMessage = " Total layers found where Linetype viewport override conflicts with the source drawing.";
        public const string ViewportLineweightCountMessage = " Total layers found where Lineweight viewport override conflicts with the source drawing.";
        public const string ViewportTransparencyCountMessage = " Total layers found where Transparency viewport override conflicts with the source drawing.";
        public const string PlotCountMessage = " Total layers found where Plot conflicts with the source drawing.";
        public const string ExtraLayersCountMessage = " Total extra layers found in the indicated target drawings that don't exist in the source drawing.";
        public const string MissingLayersCountMessage = " Total layers that exist in the source drawing but weren't found in the indicated target drawings.";
        public const string PleaseLoadSourceDrawing = " Please load a source drawing.";
        public const string SourceDrawingNotSet = "A source drawing could not be determined or was not set";
        public const string LCompIsReadOnly = "Access to the .lcomp file is restricted. It may be read-only.";
        public const string TargetDrawingsNotSet = "No target drawings were loaded. Please check the .lcomp file and try again.";
        public const string AddASourceBeforeAddingTargetsReminder = "Please add a source drawing before trying to add target drawings.";
        public const string ViewportComparisonWarning = "This comparison will compare the layer override settings on target drawing viewports that share the same layer as any viewport in the source drawing. Continue?";
        public const string MissingLayersCantRun = "To find missing layers, LCOMP must build the target drawing layers collection. It's recommended that you run a basic comparison first to accomplish thatI u. Do you wish to build the target drawing layer collection without doing a basic comparison?";
        public const string MissingLayersCancelled = "Missing layers search has been cancelled.";
    }
}
