using System;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Text;

namespace Com.Ena.Timesheet.XlUtilUtil
{
    public static class XlUtil
    {

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
