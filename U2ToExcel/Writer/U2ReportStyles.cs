using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;

namespace U2ToExcel.Writer
{
    public class U2ReportStyles
    {
        public const int ColumnWidth = 30;

        private static SLStyle _columnsFilterStyle;
        private static SLStyle _stripStyle;
        private static SLStyle _moneyStyle;


        public static SLStyle GetColumnsFilterStyle(SLDocument sl)
        {
            if (_columnsFilterStyle != null) return _columnsFilterStyle;

            _columnsFilterStyle = sl.CreateStyle();
            _columnsFilterStyle.Font.Bold = true;
            _columnsFilterStyle.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.FromArgb(217, 225, 241), System.Drawing.Color.White);

            return _columnsFilterStyle;
        }


        public static SLStyle GetMoneyStyle(SLDocument sl)
        {
            if (_moneyStyle != null) return _moneyStyle;

            _moneyStyle = sl.CreateStyle();
            _moneyStyle.FormatCode = "#,##0.00";

         
    
            return _moneyStyle;
        }


        public static SLStyle GetStripStyle(SLDocument sl)
        {
            if (_stripStyle != null) return _stripStyle;

            _stripStyle = sl.CreateStyle();
            _stripStyle.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.FromArgb(243, 244, 255), System.Drawing.Color.White);

            return _stripStyle;
        }
    }
}