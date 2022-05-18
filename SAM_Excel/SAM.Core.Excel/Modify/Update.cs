using NetOffice.ExcelApi;
using System;

namespace SAM.Core.Excel
{
    public static partial class Modify
    {
        public static bool Update(this Worksheet worksheet, DelimitedFileTable delimitedFileTable, int headerIndex = 1, int headerCount = 0, ClearOption clearOption = ClearOption.None)
        {
            if (worksheet == null)
                return false;

            int dataRowIndex = headerIndex + headerCount + 1;

            Range range_1 = null;
            Range range_2 = null;

            int lastRowIndex = worksheet.Cells.SpecialCells(NetOffice.ExcelApi.Enums.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
            int lastColumnIndex = worksheet.Cells.SpecialCells(NetOffice.ExcelApi.Enums.XlCellType.xlCellTypeLastCell, Type.Missing).Column;

            if (clearOption == ClearOption.Data)
            {
                range_1 = worksheet.Cells[dataRowIndex, 1];
                range_2 = worksheet.Cells[lastRowIndex, lastColumnIndex];

                worksheet.Range(range_1, range_2).ClearContents();
            }
            else if(clearOption == ClearOption.All)
            {
                worksheet.UsedRange.ClearContents();
            }

            range_1 = worksheet.Cells[headerIndex, 1];
            range_2 = worksheet.Cells[headerIndex, lastColumnIndex];

            object[,]  values = worksheet.Range(range_1, range_2).Value as object[,];

            for(int i = values.GetLowerBound(1); i <= values.GetUpperBound(1); i++)
            {
                string name = values[1, i]?.ToString();
                if (string.IsNullOrEmpty(name))
                    continue;

                int index = delimitedFileTable.GetColumnIndex(name);
                if (index == -1)
                    continue;

                object[,] values_Column = delimitedFileTable.GetColumnValues(index).Transpose();

                range_1 = worksheet.Cells[dataRowIndex, i];
                range_2 = worksheet.Cells[dataRowIndex + delimitedFileTable.RowCount - 1, i];

                worksheet.Range(range_1, range_2).Value = values_Column;

                if(clearOption == ClearOption.Column)
                {
                    int firstRowIndex = dataRowIndex + delimitedFileTable.RowCount;
                    if(firstRowIndex < lastColumnIndex)
                    {
                        range_1 = worksheet.Cells[firstRowIndex, i];
                        range_2 = worksheet.Cells[lastRowIndex, i];

                        worksheet.Range(range_1, range_2).ClearContents();
                    }
                }
            }

            return true;
        }

        public static bool Update(this Workbook workbook, string worksheetName, DelimitedFileTable delimitedFileTable, int headerIndex = 1, int headerCount = 0, ClearOption clearOption = ClearOption.None)
        {
            if (workbook == null || string.IsNullOrWhiteSpace(worksheetName) || delimitedFileTable == null)
                return false;

            Worksheet worksheet = workbook.Worksheet(worksheetName);
            if (worksheet == null)
            {
                worksheet = workbook.Worksheets.Add() as Worksheet;
                if (worksheet != null)
                    worksheet.Name = worksheetName;
            }

            return Update(worksheet, delimitedFileTable, headerIndex, headerCount, clearOption);
        }

        public static bool Update(string path, string worksheetName, DelimitedFileTable delimitedFileTable, int headerIndex = 1, int headerCount = 0, ClearOption clearOption = ClearOption.None)
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

                return Update(worksheet, delimitedFileTable, headerIndex, headerCount, clearOption);
            });

            return Edit(path, func);
        }
    }
}
