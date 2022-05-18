using NetOffice.ExcelApi;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SAM.Core.Excel
{
    public static partial class Modify
    {
        public static bool Write(this Worksheet worksheet, object[,] values, int rowIndex = 1, int columnIndex = 1, ClearOption clearOption = ClearOption.None)
        {
            if (worksheet == null)
                return false;

            if (values == null)
            {
                worksheet.Cells.ClearContents();
                return true;
            }

            int rowCount = values.GetUpperBound(0) - values.GetLowerBound(0);
            int columnCount = values.GetUpperBound(1) - values.GetLowerBound(1);

            Range range_1 = null;
            Range range_2 = null;

            switch (clearOption)
            {
                case ClearOption.All:
                    worksheet.UsedRange.ClearContents();
                    break;

                case ClearOption.Data:
                    worksheet.UsedRange.ClearContents();
                    break;

                case ClearOption.Column:
                    int lastRowIndex = worksheet.Cells.SpecialCells(NetOffice.ExcelApi.Enums.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
                    worksheet.Range(worksheet.Cells[1, columnIndex], worksheet.Cells[lastRowIndex, columnIndex + columnCount]).Clear();
                    break;
            }

            range_1 = worksheet.Cells[rowIndex, columnIndex];
            range_2 = worksheet.Cells[rowIndex + rowCount, columnIndex + columnCount];

            Range range = worksheet.Range(range_1, range_2);

            range.Value = Query.Values(values);
            return true;
        }

        public static bool Write(this Workbook workbook, string worksheetName, object[,] values, int rowIndex = 1, int columnIndex = 1, ClearOption clearOption = ClearOption.None)
        {
            if (workbook == null || string.IsNullOrEmpty(worksheetName) || rowIndex < 1 || columnIndex < 1)
                return false;

            Worksheet worksheet = workbook.Worksheet(worksheetName);
            if (worksheet == null)
            {
                worksheet = workbook.Worksheets.Add() as Worksheet;
                if (worksheet != null)
                    worksheet.Name = worksheetName;
            }

            return Write(worksheet, values, rowIndex, columnIndex, clearOption);
        }

        public static bool Write(string path, string worksheetName, object[,] values, int rowIndex = 1, int columnIndex = 1, ClearOption clearOption = ClearOption.None)
        {
            if (string.IsNullOrEmpty(path) || string.IsNullOrEmpty(worksheetName) || rowIndex < 1 || columnIndex < 1)
                return false;

            List<bool> result = Write(path, new string[] { worksheetName }, new List<object[,]> { values }, new int[] { rowIndex }, new int[] { columnIndex }, clearOption);
            if (result == null || result.Count == 0)
                return false;

            return result[0];
        }

        public static bool Write(string path, string worksheetName, DelimitedFileTable delimitedFileTable, int rowIndex = 1, int columnIndex = 1, ClearOption clearOption = ClearOption.None)
        {
            if(string.IsNullOrWhiteSpace(path) || delimitedFileTable == null || string.IsNullOrWhiteSpace(worksheetName))
            {
                return false;
            }

            return Write(path, worksheetName, delimitedFileTable.GetValues(), rowIndex, columnIndex, clearOption);
        }

        public static List<bool> Write(string path, IEnumerable<string> worksheetNames, IEnumerable<object[,]> values, IEnumerable<int> rowIndexes = null, IEnumerable<int> columnIndexes = null, ClearOption clearOption = ClearOption.None)
        {
            if (string.IsNullOrEmpty(path) || worksheetNames == null || worksheetNames.Count() == 0 || values == null)
            {
                return null;
            }

            List<bool> result = new List<bool>();

            int count = values.Count();
            if (count == 0)
            {
                return result;
            }

            Func<Workbook, bool> func = new Func<Workbook, bool>((Workbook workbook) =>
            {
                if (workbook == null)
                {
                    return false;
                }

                string worksheetName = null;
                int rowIndex = 1;
                int columnIndex = 1;
                for (int i = 0; i < count; i++)
                {
                    if (i < worksheetNames.Count())
                        worksheetName = worksheetNames.ElementAt(i);

                    if (rowIndexes != null && i < rowIndexes.Count())
                        rowIndex = rowIndexes.ElementAt(i);

                    if (columnIndexes != null && i < columnIndexes.Count())
                        columnIndex = columnIndexes.ElementAt(i);

                    bool succeded = Write(workbook, worksheetName, values.ElementAt(i), rowIndex, columnIndex, clearOption);
                    result.Add(succeded);
                }

                return true;
            });

            bool edited = Edit(path, func);
            if(!edited)
            {
                return null;
            }

            return result;
        }
    }
}
