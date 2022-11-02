using Autodesk.AutoCAD.Colors;
using System;
using System.Collections.Generic;
using System.Text;

namespace LayerManager.Helpers.AutoCADHelpers.ColorHelpers
{
    public class ColorHelpers
    {
        /// <summary>
        /// Sets object color from AutoCAD's color index.
        /// </summary>
        /// <param name="target">the object to be colored</param>
        /// <param name="mColor">the color from the AutoCAD color index. NOTE: Use the ACAD_COLORS class</param>
        public static void SetObjectColorByACADIndex(object target, short mColor)
        {
            target = Color.FromColorIndex(ColorMethod.ByAci, mColor);
        }
    }
}
