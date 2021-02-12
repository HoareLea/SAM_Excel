using NetOffice.ExcelApi;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SAM.Core.Excel
{
    public static partial class Modify
    {
        public static bool Write(this Worksheet worksheet, object[,] values, int rowIndex = 1, int columnIndex = 1)
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

            Range range_1 = worksheet.Cells[rowIndex, columnIndex];
            Range range_2 = worksheet.Cells[rowIndex + rowCount, columnIndex + columnCount];

            Range range = worksheet.Range(range_1, range_2);

            range.Value = values;
            return true;
        }

        public static bool Write(this Workbook workbook, string worksheetName, object[,] values, int rowIndex = 1, int columnIndex = 1)
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

            return Write(worksheet, values, rowIndex, columnIndex);
        }

        public static bool Write(string path, string worksheetName, object[,] values, int rowIndex = 1, int columnIndex = 1)
        {
            if (string.IsNullOrEmpty(path) || string.IsNullOrEmpty(worksheetName) || rowIndex < 1 || columnIndex < 1)
                return false;

            List<bool> result = Write(path, new string[] { worksheetName }, new List<object[,]> { values }, new int[] { rowIndex }, new int[] { columnIndex });
            if (result == null || result.Count == 0)
                return false;

            return result[0];
        }

        public static List<bool> Write(string path, IEnumerable<string> worksheetNames, IEnumerable<object[,]> values, IEnumerable<int> rowIndexes = null, IEnumerable<int> columnIndexes = null)
        {
            if (string.IsNullOrEmpty(path) || worksheetNames == null || values == null)
                return null;

            List<bool> result = new List<bool>();

            int count = values.Count();
            if (count == 0)
                return result;

            if (worksheetNames.Count() == 0)
                return result;

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

                string worksheetName = null;
                for (int i = 0; i < count; i++)
                {
                    if (i < worksheetNames.Count())
                        worksheetName = worksheetNames.ElementAt(0);

                    int rowIndex = 1;
                    if (rowIndexes != null && i < rowIndexes.Count())
                        rowIndex = rowIndexes.ElementAt(0);

                    int columnIndex = 1;
                    if (columnIndexes != null && i < columnIndexes.Count())
                        columnIndex = columnIndexes.ElementAt(0);

                    bool succeded = Write(workbook, worksheetName, values.ElementAt(i), rowIndex, columnIndex);
                    result.Add(succeded);
                }

                if (result != null && result.Contains(true))
                    workbook.SaveAs(path);

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
