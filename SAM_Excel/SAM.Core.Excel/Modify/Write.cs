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
                worksheet.Cells.Clear();
                return true;
            }

            int rowCount = values.GetUpperBound(0) - values.GetLowerBound(0);
            int columnCount = values.GetUpperBound(1) - values.GetLowerBound(1);

            Range range_1 = null;
            Range range_2 = null;

            switch (clearOption)
            {
                case ClearOption.All:
                    worksheet.UsedRange.Clear();
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
            List<bool> result = new List<bool>();

            int count = values.Count();
            if (count == 0)
                return result;

            if (worksheetNames.Count() == 0)
                return result;

            Dictionary<string, Func<Worksheet, bool>> funcs = new Dictionary<string, Func<Worksheet, bool>>();

            for(int i = 0; i < count; i++)
            {
                string worksheetName = null;
                if (i < worksheetNames.Count())
                    worksheetName = worksheetNames.ElementAt(0);

                int rowIndex = 1;
                if (rowIndexes != null && i < rowIndexes.Count())
                    rowIndex = rowIndexes.ElementAt(0);

                int columnIndex = 1;
                if (columnIndexes != null && i < columnIndexes.Count())
                    columnIndex = columnIndexes.ElementAt(0);

                Func<Worksheet, bool> func = new Func<Worksheet, bool>((Worksheet worksheet) =>
                {
                    if (worksheet == null)
                    {
                        return false;
                    }

                    return Write(worksheet, values.ElementAt(i), rowIndex, columnIndex, clearOption);
                });

                funcs[worksheetName] = func;
            }
            

            if(funcs == null || funcs.Count == 0)
            {
                return result;
            }

            return Write(path, funcs);
        }

        public static bool Write(string path, string worksheetName, Func<Worksheet, bool> func)
        {
            if(string.IsNullOrEmpty(path) || string.IsNullOrEmpty(worksheetName) || func == null)
            {
                return false;
            }

            Dictionary<string, Func<Worksheet, bool>> funcs = new Dictionary<string, Func<Worksheet, bool>>();
            funcs[worksheetName] = func;

            List<bool> results = Write(path, funcs);

            return results != null && results.Count != 0 && results[0];
        }

        public static List<bool> Write(string path, Dictionary<string, Func<Worksheet, bool>> funcs)
        {
            if (string.IsNullOrEmpty(path) || funcs == null || funcs.Count == 0)
                return null;

            List<bool> result = new List<bool>();

            Application application = null;

            bool screenUpdating = false;
            bool displayStatusBar = false;
            bool enableEvents = false;

            try
            {
                application = new Application();
                application.DisplayAlerts = false;
                application.Visible = false;

                screenUpdating = application.ScreenUpdating;
                displayStatusBar = application.DisplayStatusBar;
                enableEvents = application.EnableEvents;

                application.ScreenUpdating = false;
                application.DisplayStatusBar = false;
                application.EnableEvents = false;

                Workbook workbook = null;
                if (System.IO.File.Exists(path))
                    workbook = application.Workbooks.Open(path);
                else
                    workbook = application.Workbooks.Add();

                foreach(KeyValuePair<string, Func<Worksheet, bool>> keyValuePair in funcs)
                {
                    result.Add(false);

                    if (string.IsNullOrEmpty(keyValuePair.Key) || keyValuePair.Value == null)
                    {
                        continue;
                    }

                    Worksheet worksheet = workbook.Worksheet(keyValuePair.Key);
                    if (worksheet == null)
                    {
                        continue;
                    }

                    result[result.Count - 1] = keyValuePair.Value.Invoke(worksheet);
                }

                if(result.Contains(true))
                {
                    workbook.SaveAs(path);
                }

                workbook.Close(false);
            }
            catch (Exception exception)
            {

            }
            finally
            {
                if (application != null)
                {
                    application.ScreenUpdating = screenUpdating;
                    application.DisplayStatusBar = displayStatusBar;
                    application.EnableEvents = enableEvents;

                    application.Quit();
                    application.Dispose();
                }
            }

            return result;
        }
    }
}
