using System;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;

namespace Com.Ena.Timesheet.XlUtilUtil
{
    public static class XlUtil
    {
        public static string GetCellValue(ICell cell)
        {
            if (cell == null) return "";
            
            switch (cell.CellType)
            {
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Numeric:
                    return cell.NumericCellValue.ToString();
                case CellType.Boolean:
                    return cell.BooleanCellValue.ToString();
                case CellType.Error:
                    return "ERROR";
                case CellType.Formula:
                    return cell.CellFormula;
                default:
                    return "";
            }
        }

        public static bool IsEmpty(ICell cell)
        {
            return cell == null || cell.CellType == CellType.Blank;
        }

        public static string GetColumnLabel(int columnIndex)
        {
            var label = new StringBuilder();
            while (columnIndex >= 0)
            {
                label.Insert(0, (char)('A' + (columnIndex % 26)));
                columnIndex = (columnIndex / 26) - 1;
            }
            return label.ToString();
        }
    }
}
