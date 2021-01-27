using NetOffice.ExcelApi;
using System;
using System.Collections.Generic;

namespace SAM.Core.Excel
{
    public static partial class Query
    {
        public static object[,] Values(string path, string worksheetName)
        {
            if (string.IsNullOrWhiteSpace(path) || string.IsNullOrEmpty(worksheetName))
                return null;

            if (!System.IO.File.Exists(path))
                return null;

            Application application = null;
            object[,] result = null;

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

                Workbook workbook = application.Workbooks.Open(path, true, true);
                if(workbook != null)
                {
                    Worksheet worksheet = workbook.Worksheet(worksheetName);
                    if (worksheet != null)
                        result = worksheet.Values();

                    workbook.Close(false);
                }

            }
            catch(Exception exception)
            {

            }
            finally
            {
                if(application != null)
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

        public static object[,] Values(this Worksheet worksheet)
        {
            if (worksheet == null)
                return null;

            return worksheet.UsedRange?.Value as object[,];
        }

        public static Dictionary<string, object[,]> Values(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return null;

            if (!System.IO.File.Exists(path))
                return null;

            Application application = null;
            Dictionary<string, object[,]> result = null;

            try
            {
                application = new Application(true);
                application.DisplayAlerts = false;

                Workbook workbook = application.Workbooks.Open(path);

                int count = workbook.Worksheets.Count;

                result = new Dictionary<string, object[,]>();
                for (int i = 0; i < count; i++)
                {
                    Worksheet worksheet = workbook.Worksheets[i + 1] as Worksheet;
                    if (worksheet != null)
                        result[worksheet.Name] = worksheet?.UsedRange?.Value as object[,];
                }
            }
            catch (Exception exception)
            {

            }
            finally
            {
                if (application != null)
                {
                    application.Quit();
                    application.Dispose();
                }
            }

            return result;
        }
    }
}