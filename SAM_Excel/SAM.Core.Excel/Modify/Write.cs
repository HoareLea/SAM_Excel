using NetOffice.ExcelApi;
using System;

namespace SAM.Core.Excel
{
    public static partial class Modify
    {
        public static bool Write(this Worksheet worksheet, object[,] values)
        {
            if (worksheet == null)
                return false;

            if (values == null)
            {
                worksheet.Cells.Clear();
                return true;
            }

            int lowerBound_1 = values.GetLowerBound(0);
            if (lowerBound_1 == 0)
                lowerBound_1 += 1;

            int lowerBound_2 = values.GetLowerBound(1);
            if (lowerBound_2 == 0)
                lowerBound_2 += 1;

            Range range_1 = worksheet.Cells[lowerBound_1, lowerBound_2];
            Range range_2 = worksheet.Cells[values.GetUpperBound(0) + lowerBound_1, values.GetUpperBound(1) + lowerBound_2];

            Range range = worksheet.Range(range_1, range_2);

            range.Value = values;
            return true;
        }

        public static bool Write(this Workbook workbook, string worksheetName, object[,] values)
        {
            if (workbook == null || string.IsNullOrEmpty(worksheetName))
                return false;

            Worksheet worksheet = workbook.Worksheet(worksheetName);
            if (worksheet == null)
            {
                worksheet = workbook.Worksheets.Add() as Worksheet;
                if (worksheet != null)
                    worksheet.Name = worksheetName;
            }

            return Write(worksheet, values);
        }

        public static bool Write(string path, string worksheetName, object[,] values)
        {
            if (string.IsNullOrEmpty(path) || string.IsNullOrEmpty(worksheetName))
                return false;

            Application application = null;
            bool result = false;

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

                Workbook workbook = null;
                if(System.IO.File.Exists(path))
                    workbook = application.Workbooks.Open(path);
                else
                    workbook = application.Workbooks.Add();

                result = Write(workbook, worksheetName, values);
                if(result)
                    workbook.SaveAs(path);

                workbook.Close(false);
            }
            catch (Exception exception)
            {
                result = false;
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
