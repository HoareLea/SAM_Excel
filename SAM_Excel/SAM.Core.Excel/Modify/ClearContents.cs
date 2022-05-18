using NetOffice.ExcelApi;
using System;

namespace SAM.Core.Excel
{
    public static partial class Modify
    {
        public static bool ClearContents(this Worksheet worksheet, Range<int> rowRange, Range<int> columnRange)
        {
            if (worksheet == null || rowRange == null || columnRange == null)
            {
                return false;
            }

            int rowIndex_Min = rowRange.Min;
            if(rowIndex_Min <= 0)
            {
                rowIndex_Min = 1;
            }

            int columnIndex_Min = columnRange.Min;
            if (columnIndex_Min <= 0)
            {
                columnIndex_Min = 1;
            }

            int rowIndex_Max = rowRange.Max;
            if (rowIndex_Max <= 0 || rowIndex_Max < rowIndex_Min)
            {
                return false;
            }

            if(rowIndex_Max == int.MaxValue)
            {
                rowIndex_Max = worksheet.UsedRange.Rows.Count;
            }

            int columnIndex_Max = columnRange.Max;
            if (columnIndex_Min <= 0 || columnIndex_Max < columnIndex_Min)
            {
                return false;
            }

            if (columnIndex_Max == int.MaxValue)
            {
                columnIndex_Max = worksheet.UsedRange.Columns.Count;
            }

            Range range = worksheet.Range(worksheet.Cells[rowIndex_Min, columnIndex_Min], worksheet.Cells[rowIndex_Max, columnIndex_Max]);

            if(range == null)
            {
                return false;
            }

            range.ClearContents();

            return true;
        }

        public static bool ClearContents(this Worksheet worksheet, int rowStart, int columnStart, int rowEnd, int columnEnd)
        {
            if(worksheet == null)
            {
                return false;
            }

            return ClearContents(worksheet, new Range<int>(rowStart, rowEnd), new Range<int>(columnStart, columnEnd));
        }

        public static bool ClearContents(this string path, string worksheetName, int rowStart, int columnStart, int rowEnd, int columnEnd)
        {
            Func<Workbook, bool> func = new Func<Workbook, bool>((Workbook workbook) =>
            {
                if (workbook == null)
                {
                    return false;
                }

                Worksheet worksheet = workbook.Worksheet(worksheetName);
                if(worksheet == null)
                {
                    return false;
                }

                return ClearContents(worksheet, rowStart, columnStart, rowEnd, columnEnd);
            });

            return Edit(path, func);
        }
    }
}
