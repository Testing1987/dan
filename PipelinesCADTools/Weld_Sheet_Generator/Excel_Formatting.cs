using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace weld_sheet
{
    class Excel_Formatting
    {
        public static void Border_Style_Bottom_Thin(Range range1)
        {
            range1.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
        }
        public static void Border_Style_Remove_RightSide(Range range1)
        {
            range1.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlLineStyleNone;
        }
        public static void Border_Style_Thin(Range range1)
        {
            range1.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlColorIndex.xlColorIndexAutomatic);
        }
        public static void Border_Style_Gray(Range range1)
        {
            range1.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlColorIndex.xlColorIndexAutomatic);
            range1.Borders.Color = Color.Gray;
        }
        public static void Border_Style_Thick(Range range1)
        {
            range1.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlThick, XlColorIndex.xlColorIndexAutomatic);
        }
        public static void Label_Color(Range range1)
        {
            range1.Font.Color = 8421504;
        }
    }
}
