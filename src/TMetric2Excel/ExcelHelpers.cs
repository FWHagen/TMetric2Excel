using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TMetric2Excel
{
    public static class ExcelHelpers
    {
        public static string GetCellId(int row, int col)
        {
            return GetColId(col) + row.ToString();
        }
        public static string GetColId(int col)
        {
            // this will only work to col 26: "Z" //
            return ((char)(col + 64)).ToString();
        }

        public static void WriteCellHeader(this _Worksheet worksheet, string value, int col, int row = 1, double width = 0, int indent = 0)
        {
            var cellid = GetCellId(row, col);
            var colid = GetColId(col);
            var rngid = String.Concat(colid, ":", colid);
            worksheet.Range[cellid].Value = value;
            //worksheet.Range[cellid].Style.HorizontalAlignment = HorizontalAlignType.Left;
            worksheet.Range[cellid].IndentLevel = indent;
            worksheet.Range[cellid].Font.Bold = true;
            if (width > 0)
                worksheet.Columns[rngid].ColumnWidth = width;
        }

        public static void WriteCell(this _Worksheet worksheet, object value, int col, int row, int indent = 0)
        {
            var cellid = GetCellId(row, col);
            var colid = GetColId(col);
            var rngid = String.Concat(colid, ":", colid);
            worksheet.Range[cellid].Value = value;
            //worksheet.Range[cellid].Style.HorizontalAlignment = HorizontalAlignType.Left;
            //if (indent >= 0)
                worksheet.Range[cellid].IndentLevel = indent;
        }

    }
}
